import json
import locale
import logging
from typing import List, Dict, Callable, Any
from files import PDFManager
from db import Database
from cloud import CloudServices
import utils as utils
from enum import Enum
from datetime import datetime

locale.setlocale(locale.LC_ALL, 'en_US.UTF8')


class StateFlag(Enum):
    UNINITIALIZED = 0
    CONNECTED = 1
    DISCONNECTED = 2


class EmesDispatch:

    ZONES = \
        [
            "Norte",
            "Sur",
            "Oriente",
            "Occidente",
            "Regiones"
        ]

    def __init__(self) -> None:
        """
        Constructor
        """
        # Initialize object attibutes
        self._params = None
        self._pdf = None
        self._db = None
        self._cloud = None

        # State flag variable
        self._state = StateFlag.UNINITIALIZED

        # Get config data
        with open('config/config.json') as f:
            self.__config_data = json.load(f)

        self._users = self.__config_data["users"]

        # Get initial couriers and packers names
        self.__names = self.get_users()

    @property
    def state(self) -> StateFlag:
        return self._state

    @property
    def params(self) -> Dict[str, Any]:
        return self._params

    @params.setter
    def params(self, in_params: Dict[str, Any]) -> None:
        self._params = in_params

    @property
    def users(self) -> Dict[str, Any]:
        return self._users

    @users.setter
    def users(self, in_users: Dict[str, Any]) -> None:
        self._users = in_users

    @property
    def names(self) -> Dict[str, List[str]]:
        return self.__names

    @names.setter
    def names(self, in_names: Dict[str, List[str]]) -> None:
        self.__names = in_names

    def get_credentials(self, user: str) -> Dict:
        """
        Converts .json file into properly dict info

        Args:
            user (str): User name

        Returns:
            Dict: User specified data dict
        """
        user_data = self._users.get(user, {})

        params = \
            {
                key: self.__config_data["general"][key]
                for key in self.__config_data["general"]
            }

        params.update(user_data)

        return params

    def __is_duplicated(self, order_id: int) -> bool:
        """
        Find and delete order duplicates

        Args:
            order_id (int): Order ID number

        Returns:
            bool: True if order is duplicated
        """
        if self._db and self._db.exists("factura", order_id):
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M")

            # Delete repeated file
            self._pdf.delete_file(f"{order_id}.pdf")

            # Log a message if the invoice ID already exists in the database
            logging.info(
                f"La orden {order_id} no se pudo ingresar porque está repetida (Hora: {current_time})"
            )
            return True

        return False

    def pdf_to_db(self) -> bool:
        """
        Update database with data extracted from the PDFs

        Returns:
            bool: True if convert PDF data into DB data was successful
        """
        try:
            # Check if there if pdf object was created
            if not self._pdf:
                return False

            # Extract data from pdf object
            pdf_data = self._pdf.update()

            if pdf_data is None:
                return False

            # Loop through pdf data and update database for each item
            for name, data in pdf_data.items():
                order_id = int(data["remision"])

                # Check if invoice ID already exists in the database
                if self.__is_duplicated(order_id):
                    continue

                # Get client name and pharmacy address
                tienda = self._cloud.get_client_name(data["tienda"])
                farmacia = tienda + " " + data["direccion"]

                # Convert the value of the invoice to an integer
                valor = utils.str2int(data["valor"])

                # Convert the number of items to an integer
                items = int(locale.atof(data["items"]))

                # Insert values into database
                self._db.insert_values(
                    items=items,
                    factura=order_id,
                    farmacia=farmacia,
                    valor=valor,
                    facturador=data["facturador"],
                    forma_pago=data["forma"]
                )

                # Delete the processed file
                self._pdf.delete_file(name)

            return True

        except (KeyError, ValueError, TypeError):
            logging.error(
                "Error al insertar nuevos documentos en la base de datos",
                exc_info=True
            )
            return False

    def is_uninitialized(self) -> bool:
        """
        Check if flag is equal to StateFlag.UNINITIALIZED
        """
        return self._state.value == StateFlag.UNINITIALIZED.value

    def is_connected(self) -> bool:
        """
        Check if flag is equal to StateFlag.CONNECTED
        """
        return self._state.value == StateFlag.CONNECTED.value

    def is_disconnected(self) -> bool:
        """
        Check if flag is equal to StateFlag.DISCONNECTED
        """
        return self._state.value == StateFlag.DISCONNECTED.value

    def is_non_connected(self) -> bool:
        """
        Check if flag is equal to StateFlag.UNINITIALIZED or StateFlag.DISCONNECTED
        """
        return self._state.value != StateFlag.CONNECTED

    def initialize(self, user: str) -> None:
        """
        Initialize APIs

        Args:
            params (str): User name
        """
        try:
            self._username = user
            self._params = self.get_credentials(user)

            self._pdf = PDFManager(
                self._params['folder_id_remisiones']
            )

            self._db = Database(
                self._params['json_path_db'],
                self._params['db_url']
            )

            self._cloud = CloudServices(
                self._params['folder_id_planillas'],
                self._params['json_path_gs'],
                EmesDispatch.ZONES
            )

            # Create Spreadsheet (Google Sheets)
            self._cloud.init(self.__names)

            # Set flag to INITIALIZED
            self._state = StateFlag.CONNECTED

        except Exception as e:
            logging.error(
                f"Error al inicializar el servidor >>> {e}",
                exc_info=True
            )

    def disconnect(self) -> None:
        """
        Disconnect listener and delete database ref
        """
        try:
            # Close DB connection
            self._db.stop_connection()

            # Close Google Sheets session
            self._cloud.client.session.close()

            # Delete .rpt files
            self._pdf.delete_all_files()

            # Delete all objects and stop strong references
            self._db = None
            self._pdf = None
            self._cloud = None

            # Set flag to DISCONNECTED
            self._state = StateFlag.DISCONNECTED

        except Exception as e:
            logging.error(
                f"Error al desconectar el servidor >>> {e}",
                exc_info=True
            )

    def generate_report(
            self,
            type: str,
            filename: str) -> None:
        """
        Generate speficied report

        Args:
            type (str): Type report
            filename (str): File name where is the data
        """
        try:
            if type == "treasury":
                self._cloud.create_treasury_report(filename)
            elif type == "dispatches":
                self._cloud.create_dispatches_report(filename)
            else:
                logging.info(f"No existe el reporte de tipo {type}")

        except Exception as e:
            logging.error(
                f"Error al generar el reporte de {type} >>> {e}",
                exc_info=True
            )

    def find_item(self, order_id: int) -> Dict[str, Any] | None:
        """
        Find item by code

        Args:
            order_id (int): Order ID number

        Returns:
            Dict[str, Any]: Item information
        """
        return self._db.search_code(order_id)

    def get_zones(self) -> List[str]:
        """
        Get distribution zones/regions

        Returns:
            List[str]: Zones where Emes delivers 
        """
        return self._cloud.ZONES

    def get_courier_names(self) -> List[str]:
        """
        Get names of all couriers (Mensajeros)

        Returns:
            List[str]: Courier names
        """
        return self._cloud.names["couriers"]

    def get_packer_names(self) -> List[str]:
        """
        Get names of all packers (Empacadores)

        Returns:
            List[str]: Packer names
        """
        return self._cloud.names["packers"]

    def insert_rows_in_sheet(self, values: List[Any]) -> bool:
        """
        Insert new row ino GS file

        Args:
            values (List[Any]): Row data

        Returns:
            bool: True if was successful
        """
        return self._cloud.insert_rows(values)

    def get_empty_cols(self) -> List[str]:
        """
        Get Google Sheets file sheets empty cols
        """
        return self._cloud.EMPTY_COLS

    def set_listener_callback(self, callback: Callable) -> None:
        """
        Set Database listener callback
        """
        self._db.start_listener(callback)

    def db_query(
            self,
            item: str,
            value: Any) -> Dict[str, Any] | None:
        """
        Returns the value of specified item

        Args:
            item (str): Item name (eg. factura)
            value (Any): Value (eg. 548492)

        Returns:
            Dict[str, Any] | None: Specified value
        """
        try:
            if not self._db:
                return None

            # Get data from DB based on query
            data = self._db._where(item, value)
            return data

        except Exception as e:
            logging.error(
                f"Error al solicitar los datos a la base de datos del item: {item} >>> {e}",
                exc_info=True
            )
            return None

    def get_values_by_key(self, key: Any) -> List[Any]:
        """
        Get key value of Database

        Args:
            key (Any): Document key (uid)

        Returns:
            List[Any]: Database key items
        """
        return self._db.ref.child(key).get()

    def update_state_by_key(self, key: Any, state: int) -> None:
        """
        Update DB state value

        Args:
            key (Any): Document key
            state (Any): State value
        """
        self._db.ref.child(key).update(
            {
                "estado": state
            }
        )

    def update_state_by_keys(
            self,
            data: Dict[str, Dict[str, Any]]) -> None:
        """
        Update multiple DB items in one API call

        Args:
            data (Dict[str, Dict[str, Any]]): Format equivalent to [key, {"estado": value}]
        """
        self._db.ref.update(data)

    def update_users(self) -> Any:
        """
        Update users (Mensajeros/Empacadores)
        """
        # Update member variable
        self.__names = self.get_users()

        # Update GS sheet
        res = self._cloud.update_users(self.__names)

        return res

    def get_users(self) -> Dict[str, List[str]] | None:
        """
        Get users from shared Excel file

        Returns:
            Dict[str, List[str]]: Names of couriers and packers
        """
        try:
            from pandas import read_excel

            ID = self.__config_data["general"]["folder_id_remisiones"]
            PATH = fr"G:\.shortcut-targets-by-id\{ID}\Despachos\users.xlsx"

            sheet1 = read_excel(PATH, sheet_name="Mensajeros")
            sheet2 = read_excel(PATH, sheet_name="Empacadores")

            data = \
                {
                    "couriers": [
                        list(e.values())[0] for e in sheet1.to_dict("records")
                    ],
                    "packers": [
                        list(e.values())[0] for e in sheet2.to_dict("records")
                    ]
                }

            return data

        except Exception as e:
            logging.error(
                f"Error al obtener los nombres de los mensajeros y empacadores >>> {e}",
                exc_info=True
            )
            return None
