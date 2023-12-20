import logging
import firebase_admin
from firebase_admin import db
from datetime import datetime
from typing import Callable, Any, Dict


class Database:

    def __init__(
            self,
            json_path: str,
            db_url: str) -> None:
        """
        Constructor

        Args:
            json_path (str): firebase json path
            db_url (str): database url

        Raises:
            Exception: unauthorized firebase admin session
        """
        self.cred_firebase = firebase_admin.credentials.Certificate(
            cert=json_path
        )

        if not firebase_admin._apps:
            self.__app = firebase_admin.initialize_app(
                self.cred_firebase,
                {
                    'databaseURL': db_url
                }
            )

            if not self.__app:
                raise Exception("Error creating Firebase object")
            else:
                self.ref = db.reference("/empaques/")
                self.listener = None

    def insert_values(
            self,
            items: int,
            factura: int,
            farmacia: str,
            valor: int,
            facturador: str,
            forma_pago: str) -> None:
        """
        Args:
            items (int): total products
            factura (int): invoice number
            farmacia (str): client name
            valor (int): total invoice value
            facturador (str): Facturador name
            forma_pago (str): Forma de pago (Contado / A xx días)
        """
        try:
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M")

            self.ref.push().set(
                {
                    "articulos": items,
                    "bolsas": 0,
                    "cajas": 0,
                    "empacador": "",
                    "estado": 0,
                    "factura": factura,
                    "farmacia": farmacia,
                    "fecha": current_time,
                    "lios": 0,
                    "region": "",
                    "valor": valor,
                    "facturador": facturador,
                    "forma": forma_pago,
                    "uid": ""
                }
            )

        except Exception as e:
            logging.error(
                f"Error al insertar nuevos documentos en la base de datos >>> {e}",
                exc_info=True
            )

    def exists(
            self,
            item: str,
            value: Any) -> bool:
        """
        Checks if there is an item with the specified value

        Args:
            item (str): Item name (eg. factura)
            value (Any): Value (eg. 548492)

        Returns:
            bool: True if item code exists
        """
        try:
            query = self._where(item, value)
            values = list(query.values())[0] if query else None

            return False if not values else values["estado"] != -1

        except Exception as e:
            logging.error(
                f"Error al revisar si existe una order en la base de datos >>> {e}",
                exc_info=True
            )
            return False

    def _where(
            self,
            item: str,
            value: Any) -> Dict[str, Any] | None:
        """
        Returns the value of specified item

        Args:
            item (str): Item name (eg. factura)
            value (Any): Value (eg. 548492)

        Returns:
            Dict[str, Any] | None: Document data
        """
        try:
            query = self.ref.order_by_child(item).equal_to(value)

            # Get query values
            results = query.get()

            if len(results) > 0:
                return \
                    {
                        key: value
                        for key, value in results.items()
                    }

            return None

        except Exception as e:
            logging.error(
                f"Error al revisar si existe un item especificado con un valor determinado >>> {e}",
                exc_info=True
            )
            return None

    def search_code(self, code: int) -> Dict[str, Any] | None:
        """
        Search document by specified code

        Args:
            code (int): Invoice code

        Returns:
            Dict[Any]: Document data dict
        """
        try:
            data = self._where("factura", code)
            return data

        except Exception as e:
            logging.error(
                f"Error al buscar un documento por número de orden >>> {e}",
                exc_info=True
            )
            return None

    def search_uid(self, uid: str) -> Dict[str, Any] | None:
        """
        Args:
            uid (str): Document key

        Returns:
            Dict[str, Any] | None: Document values
        """
        try:
            query = self.ref.order_by_key().equal_to(uid)

            # Get query values
            results = query.get()

            if len(results) > 0:
                return \
                    {
                        key: value
                        for key, value in results.items()
                    }

            return None

        except Exception as e:
            logging.error(
                f"Error al revisar si existe una order con un uid especificado >>> {e}",
                exc_info=True
            )
            return None

    def start_listener(self, callback: Callable) -> None:
        """
        Starts the listener to listen for real-time events from Firebase.

        Args:
            callback (Callable): Function to be executed when events are received.
        """
        self.listener = self.ref.listen(callback)

    def _stop_listener(self) -> None:
        """
        Stops the listener.
        """
        self.listener.close()

    def stop_connection(self) -> None:
        """
        Stops the connection to Firebase and deletes the app.
        """
        # Stop the listener first to ensure the connection is closed properly
        self._stop_listener()

        # Delete the Firebase app to clean up resources and terminate the connection
        firebase_admin.delete_app(self.__app)
