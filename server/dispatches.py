from server.files import PDFManager
from server.db import Database
from server.cloud import CloudServices
import logging
import locale
import server.utils as utils
from firebase_admin.db import ListenerRegistration

locale.setlocale(locale.LC_ALL, 'en_US.UTF8')


class EmesDispatch:

    def __init__(
            self,
            db_url: str,
            folder_id_planillas: str,
            folder_id_remisiones: str,
            json_path_db: str,
            json_path_gs: str) -> None:
        """
        Constructor

        Create all necessary objects
        """
        self.pdf = PDFManager(
            folder_id_remisiones
        )

        self.db = Database(
            json_path_db,
            db_url
        )

        # create listener (searching for Firebase updates)
        if self.db is not None:
            self._listen = self.db.ref.listen(self.__callback)
        else:
            self._listen = None

        self.cloud = CloudServices(
            folder_id_planillas,
            json_path_gs
        )

    @property
    def listen(self) -> ListenerRegistration:
        return self._listen

    def __update_db(self) -> None:
        """
        Update database with data extracted from the pdfs
        """
        pdf_data = self.pdf.update()

        try:
            for name, data in pdf_data.items():
                farmacia_id = data['tienda'] + " " + data['direccion']
                factura_id = int(data['remision'])

                if data['valor'] == '':
                    valor = '$0'
                else:
                    valor = utils.str2int(
                        data['valor']
                    )

                items = int(locale.atof(
                    data['items']
                ))

                self.db.insert_values(
                    items=items,
                    factura=factura_id,
                    farmacia=farmacia_id,
                    valor=valor
                )

                self.pdf.delete_file(name)

        except (KeyError, ValueError, TypeError):
            logging.error(
                "Error al insertar nuevos documentos en la base de datos",
                exc_info=True
            )

    def __document_to_row(
            self,
            key: str,
            data: dict) -> list:
        """
        Convert document data into Sheets row values

        Args:
            key (str): document uid
            data (dict): database document

        Returns:
            list: row values (Google Sheets)
        """
        return [
            data['farmacia'],
            data['factura'],
            data['cajas'],
            data['bolsas'],
            data['lios'],
            data['valor'],
            '',
            data['empacador'],
            data['fecha'][11:],
            '',
            '',
            '',
            key
        ]

    def _event_insert(
            self,
            event,
            key: str,
            value: dict) -> None:
        """
        Insert new row into sheets and update to state = 2

        Args:
            event (Any): listener event
            key (str): database uid
            value (dict): document data
        """
        if value['estado'] == 1:
            values = self.__document_to_row(
                key=key,
                data=value
            )

            self.cloud.insert_row(
                values,
                event.data['region']
            )

            self.db.ref.child(key).update(
                {
                    "estado": 2
                }
            )

    def _event_update(
            self,
            event,
            key: str,
            value: dict) -> None:
        """
        Call if there is an update in the database

        Args:
            event (Any): database listener event
            key (str): database uid
            value (dict): document data
        """
        if value['estado'] == 2:
            values = self.__document_to_row(
                key=key,
                data=value
            )

            if 'region' in event.data.keys():
                del_values = self.cloud.delete_row(
                    key=key
                )

                # copy empty col values
                for i in [6, 9, 10, 11]:
                    values[i] = del_values[i]

                self.cloud.insert_row(
                    values=values,
                    zone=value['region']
                )
            else:
                self.cloud.update_row(
                    zone=value['region'],
                    key=key,
                    values=[values]
                )

    def __callback(self, event) -> None:
        """
        Callback function when some change occurs in database

        Args:
            event (Any): database listener event
        """
        try:
            split_path = event.path.split('/')

            # callback logic - update or set event
            if len(split_path) == 2 and split_path[1] != '':
                key = split_path[1]
                value = self.db.ref.child(key).get()

                if utils.all_dict(value, event.data):
                    if event.event_type == 'patch':
                        self._event_update(
                            event,
                            key,
                            value
                        )
                    else:  # put event
                        self._event_insert(
                            event,
                            key,
                            value
                        )

        except Exception as e:
            logging.error(
                f'Error {e} al actualizar el archivo de Google Sheets',
                exc_info=True
            )

        # else:
        #     print(event.event_type)  # 'put' or 'patch'
        #     print(event.path)  # key (reference)
        #     print(event.data)

    def _run(self) -> None:
        """
        Main loop function
        """
        try:
            while self.able_to_run:
                # step 1: read Drive pdf
                # step 2 (set): write in Firebase db
                self.__update_db()

        except KeyboardInterrupt:
            print('Main loop stopped...')

    def connect(self) -> None:
        """
        Initialize main infinite loop for update db
        """
        self.able_to_run = True
        self._run()

    def generate_report(
            self,
            type: str,
            filename: str) -> None:
        """
        Generate speficied report

        Args:
            type (str): report type
        """
        if type == "main":
            self.cloud.create_couriers_report(filename)
            self.cloud.create_packers_report(filename)
        elif type == "treasury":
            self.cloud.create_treasury_report(filename)
        elif type == "dispatches":
            self.cloud.create_dispatches_report(filename)

    def disconnect(self) -> None:
        """
        Disconnect listener and delete database ref
        """
        self.able_to_run = False

        # stop listener
        self._listen.close()

        # stop strong reference
        del self.db
        self.db = None

        # close Google Sheets session
        self.cloud.client.session.close()

        if self._listen._thread.is_alive:
            print("Listener is still alive...")
        else:
            print("Listener has been stopped and closed...")
