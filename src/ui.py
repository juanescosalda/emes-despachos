import sys
import logging
import requests
from firebase_admin.db import Event
from typing import Dict, Callable, Any, Tuple

from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *

import utils
from worker import Worker
from server import EmesDispatch

from ui_config import ConfigUi
from ui_auth import LoginUi
from ui_services import ServicesUi, EditUpdater
from ui_updaters import TableUpdater, RowUpdater
from ui_packed import PackedUi
from ui_invoiced import InvoiceUi
from ui_msg_box import MessageBox

logging.basicConfig(
    filename='despachos.log',
    format='%(asctime)s - %(message)s',
    datefmt='%d-%b-%y %H:%M:%S'
)


class FileProcessingThread(QThread):

    def __init__(self, server: EmesDispatch):
        """
        Constructor

        Args:
            server (EmesReception): Server object
        """
        super().__init__()

        self.server = server

    def run(self) -> None:
        """
        Entry point for the file processing thread.
        """
        while True:
            # Perform file search and processing
            self.server.pdf_to_db()

            # Pause the loop for 1000 milliseconds
            self.msleep(1000)

            # Verifica periódicamente si se ha solicitado detener el hilo
            if self.isInterruptionRequested():
                break


class MainWindow(QMainWindow):

    def __init__(self, server: EmesDispatch | None = None):
        """
        Constructor

        Args:
            emes (EmesDispatch | None, optional): Server object. Defaults to None.
        """
        super().__init__()

        # Intial load calls counter
        self._is_first_call = True

        # Server object
        self._server = server

        # Set main window settings
        self._set_window_settings()

        # Init threadpool
        self._threadpool = QThreadPool()

        # Create form tabs
        self.auth_tab = LoginUi(self)
        self.services_tab = ServicesUi(self)
        self.config_tab = ConfigUi(self)
        self.packed_tab = PackedUi(
            parent=self,
            couriers=self._server.names["couriers"],
            zones=self._server.ZONES
        )
        self.invoiced_tab = InvoiceUi(self)

        # Initialize User Interface (UI)
        self.init_UI()

        # Intialize table updaters
        self._initialize_updaters()

        # Set timer to check every 10 seconds
        self.timer = QTimer()
        self.timer.timeout.connect(self.__check_internet_connection)
        self.timer.start(10000)

        # Set timer to find new files and process
        self.file_processing_thread = None

        # Message boxes object
        self.msg = MessageBox(self)

    @property
    def user(self) -> str:
        """
        Get username specified in Auth tab
        """
        return self.auth_tab.username_edit.text()

    @property
    def password(self) -> str:
        """
        Get password specified in Auth tab
        """
        return self.auth_tab.password_edit.text()

    def _set_window_settings(self) -> None:
        """
        Define main window widget settings
        """
        # Set window icon
        self.setWindowIcon(QIcon("icons/logo_icon.ico"))

        # Get screen dimensions
        self.screen = QGuiApplication.primaryScreen()
        self.screen_size = self.screen.availableSize()
        self.sidebar_width = int(self.screen_size.width() / 9)
        self.sidebar_height = int(self.screen_size.height())

        # Set the title of main window
        self.setWindowTitle("Emes: Desconectado")
        self.setWindowState(Qt.WindowState.WindowMaximized)

    def _initialize_updaters(self) -> None:
        """
        Initialize table and row updaters
        """
        # Set table updater for PackedUi
        self._table_updater = TableUpdater(self.packed_tab)
        self._table_updater.update.connect(self._table_updater.update_table)
        self._table_updater.reset.connect(self._table_updater.reset_table)
        self._table_updater.add.connect(self._table_updater.add_row)

        # Set table updater for InvoicedUi
        self._invoiced_table_updater = TableUpdater(self.invoiced_tab)
        self._invoiced_table_updater.update.connect(
            self._invoiced_table_updater.update_table)
        self._invoiced_table_updater.reset.connect(
            self._invoiced_table_updater.reset_table)

        # Set row updater for PackedUi
        self._row_updater = RowUpdater(self.packed_tab)
        self._row_updater.update.connect(self._row_updater.update_row)

        # Set edit updater
        self._edit_updater = EditUpdater(self.services_tab)
        self._edit_updater.update.connect(self._edit_updater.update_edit)

    def __create_main_button(
            self,
            text: str,
            func: Callable,
            icon_path: str) -> QPushButton:
        """
        Create left widget button

        Args:
            text (str): Button text
            func (Callable): Click event function
            icon_path (str): Icon path

        Returns:
            QPushButton: Left widget button styled
        """
        # Create icon
        icon = QIcon(icon_path)

        # Create button
        button = QToolButton(self)
        button.setObjectName('left_button')
        button.clicked.connect(func)
        button.setSizePolicy(QSizePolicy.Policy.Expanding,
                             QSizePolicy.Policy.Expanding)
        button.setIcon(icon)
        button.setIconSize(button.fontMetrics().boundingRect("A").size() * 11)
        button.setText(text)
        button.setToolButtonStyle(
            Qt.ToolButtonStyle.ToolButtonTextUnderIcon)

        return button

    def create_left_buttons(self) -> Tuple[QToolButton]:
        """
        Create left main sidebar buttons

        Returns:
            Tuple[QToolButton]: Tuple with all left tab buttons
        """
        btn_1 = self.__create_main_button(
            text="LOGIN",
            func=self.click_button1,
            icon_path="icons/user_icon.svg")

        btn_2 = self.__create_main_button(
            text="FACTURADOS",
            func=self.click_button2,
            icon_path="icons/order_icon.svg")

        btn_3 = self.__create_main_button(
            text="EMPACADOS",
            func=self.click_button3,
            icon_path="icons/edit_icon.svg")

        btn_4 = self.__create_main_button(
            text="SERVICIOS",
            func=self.click_button4,
            icon_path="icons/services_icon.svg")

        btn_5 = self.__create_main_button(
            text="CONFIGURACIÓN",
            func=self.click_button5,
            icon_path="icons/config_icon.svg")

        return btn_1, btn_2, btn_3, btn_4, btn_5

    def create_right_tabs(self) -> Tuple[QWidget]:
        """
        Create main right tabs
        """
        tab_1 = self.auth_tab.get_form()
        tab_2 = self.invoiced_tab.get_form()
        tab_3 = self.packed_tab.get_form()
        tab_4 = self.services_tab.get_form()
        tab_5 = self.config_tab.get_form()

        return (tab_1, tab_2, tab_3, tab_4, tab_5)

    def __start_file_processing(self) -> None:
        """
        Start the file processing thread when the button is clicked.
        """
        try:
            if self.file_processing_thread is not None and \
                    self.file_processing_thread.isRunning():
                return

            # Create a new instance of the file processing thread
            self.file_processing_thread = FileProcessingThread(self._server)
            self.file_processing_thread.start()

        except Exception as e:
            logging.error(
                f"Error al iniciar el threading que busca nuevos PDFs >>> {e}",
                exc_info=True
            )

    def __stop_file_processing(self) -> bool:
        """
        Stop file processing threading

        Returns:
            bool: True if threading was successfully stopped
        """
        try:
            if self.user == "despachos":
                if self.file_processing_thread is not None:
                    self.file_processing_thread.requestInterruption()
                    self.file_processing_thread.quit()
                    self.file_processing_thread.wait()

                    return self.file_processing_thread.isFinished()
                else:  # Is main account but threading is None
                    return False

            return True

        except Exception as e:
            logging.error(
                f"Error al finalizar el threading que busca nuevos PDFs >>> {e}",
                exc_info=True
            )
            return False

    def init_UI(self) -> None:
        """
        Initialize user interface
        """
        # Create main buttons
        left_buttons = self.create_left_buttons()
        right_tabs = self.create_right_tabs()

        # Create left sidebar layout
        left_layout = QVBoxLayout()
        left_layout.addWidget(left_buttons[0])
        left_layout.addWidget(left_buttons[1])
        left_layout.addWidget(left_buttons[2])
        left_layout.addWidget(left_buttons[3])
        left_layout.addWidget(left_buttons[4])

        # Create left widget
        left_widget = QWidget()
        left_widget.setLayout(left_layout)
        left_widget.setFixedWidth(self.sidebar_width)
        left_widget.setObjectName('left_widget')
        left_widget.setAutoFillBackground(True)

        self.right_widget = QTabWidget()
        self.right_widget.tabBar().setObjectName("main_tab")
        self.right_widget.addTab(right_tabs[0], '')
        self.right_widget.addTab(right_tabs[1], '')
        self.right_widget.addTab(right_tabs[2], '')
        self.right_widget.addTab(right_tabs[3], '')
        self.right_widget.addTab(right_tabs[4], '')
        self.right_widget.setCurrentIndex(0)

        main_layout = QHBoxLayout()
        main_layout.addWidget(left_widget)
        main_layout.addWidget(self.right_widget)
        main_layout.setStretch(0, 40)
        main_layout.setStretch(1, 200)

        main_widget = QWidget()
        main_widget.setLayout(main_layout)

        # Define central main widgets
        self.setCentralWidget(main_widget)

        # Set button events
        self._set_button_events()

    def __is_dispatches_account(self) -> bool:
        """
        Checks if is a dispatches user 

        Returns:
            bool: True if is "despachos", "invitado" or "supervisor" account
        """
        return self.user in ("despachos", "invitado", "supervisor")

    def __check_internet_connection(self) -> None:
        """
        Check if are connected to the Internet
        """
        try:
            requests.get("http://www.google.com", timeout=5)
        except (requests.ConnectionError, requests.Timeout):
            self.setWindowTitle("Emes: Sin conexión a internet")
            self.show_no_internet_message()

    def show_no_internet_message(self) -> None:
        """
        Create new message if no internet
        """
        msg_box = QMessageBox()
        msg_box.setWindowTitle("Sin conexión a Internet")
        msg_box.setIcon(QMessageBox.Icon.Critical)
        msg_box.addButton("Verificar", QMessageBox.ButtonRole.AcceptRole)
        msg_box.setText(
            "No hay conexión a Internet. Verifique la conexión y vuelva a intentarlo.")
        msg_box.exec()

        # Check connection when press the button
        if msg_box.clickedButton().text() == "Verificar":
            self.setWindowTitle("Emes: Conectado")
            self.__check_internet_connection()

    def _set_button_events(self) -> None:
        """
        Assign button methods to widget events
        """
        self.auth_tab.login_function(
            self.click_connect)
        self.auth_tab.logout_function(
            self.click_disconnect)
        self.services_tab.find_function(
            self.click_find)
        self.services_tab.create_treasury_report_function(
            self.click_generate_treasury_report)
        self.services_tab.create_dispatches_report_function(
            self.click_generate_dispatches_report)
        self.packed_tab.send_button_function(
            self.click_upload)
        self.invoiced_tab.delete_button_function(
            self.click_delete_invoiced)
        self.invoiced_tab.update_files_button_function(
            self.click_update_files)
        self.invoiced_tab.update_button_function(
            self.click_update_invoiced)
        self.packed_tab.delete_button_function(
            self.click_delete_packed)
        self.packed_tab.update_button_function(
            self.click_update_packed)
        self.packed_tab.update_users_button_function(
            self.click_update_users)
        self.packed_tab.pending_order_button_function(
            self.click_set_as_pending)
        self.config_tab.update_button_function(
            self.click_edit_config)

    def click_button1(self) -> None:
        """
        Switches the current index of the right widget to 0
        """
        self.right_widget.setCurrentIndex(0)

    def click_button2(self) -> None:
        """
        Switches the current index of the right widget to 1
        """
        if self._server.is_connected():
            # if self.__is_dispatches_account():
            #     self.msg.unavailable_service()
            # else:
            self.right_widget.setCurrentIndex(1)
        else:
            self.msg.non_connection_msg()

    def click_button3(self) -> None:
        """
        Checks if the user is a dispatches account
        If the user is a supervisor, displays an unavailable service message
        """
        if self._server.is_connected():
            if self.__is_dispatches_account():
                self.right_widget.setCurrentIndex(2)
            else:
                self.msg.unavailable_service()
        else:
            self.msg.non_connection_msg()

    def click_button4(self) -> None:
        """
        Checks if the user is a dispatches account
        If the user is a supervisor, displays an unavailable service message
        """
        if self._server.is_connected():
            if self.__is_dispatches_account():
                self.right_widget.setCurrentIndex(3)
            else:
                self.msg.unavailable_service()
        else:
            self.msg.non_connection_msg()

    def click_button5(self) -> None:
        """
        Switches the current index of the right widget to 4
        """
        self.right_widget.setCurrentIndex(4)

    def progress_fn(self, value: Any):
        """
        Prints the progress as a percentage

        Args:
            value (int): The progress value
        """
        print("%d%% done" % value)

    def print_output(self, output: str):
        """
        Prints the given output

        Args:
            output (str): The output to be printed
        """
        print(output)

    def thread_complete(self):
        """
        Prints a message indicating that the thread has completed
        """
        print("Thread complete!")

    def __is_packed(
            self,
            key: str,
            data: Dict[str, Any] | str) -> bool:
        """
        Check is state is 1 (packed )or 3 (pending)

        Args:
            key (str): Order uid
            data (Dict[str, Any] | str): Order data

        Returns:
            bool: True if order is packed
        """
        if isinstance(data, dict):
            return data.get("estado") in (1, 3)
        else:
            logging.warning(f"El dato no es un dict: {data}. Key: {key}")
            return False

    def _event_db_initial_load(
            self,
            data: Dict[str, Dict[str, Any]]) -> None:
        """
        Load data into QTableView if is first load

        Args:
            data (Dict[str, Dict[str, Any]]): Order data
        """
        try:
            if data and self._is_first_call:
                # Get only packed orders
                table_data = \
                    {
                        key: value
                        for key, value in data.items()
                        if self.__is_packed(key, value)
                    }

                # Emit signal to upload initial orders
                self._table_updater.update.emit(table_data)

                # Set first call into false
                self._is_first_call = False

        except Exception as e:
            logging.error(
                f"Error al realizar la carga inicial de datos >>> {e}",
                exc_info=True
            )

    def _event_db_packing(
            self,
            key: str,
            data: Dict[str, Any]) -> None:
        """
        Insert new row into sheets and update to state 2 (packed)

        Args:
            key (str): Database uid
            data (Dict[str, Any]): Document data
        """
        try:
            if isinstance(data, dict) and data.get("estado") == 1:
                self._table_updater.add.emit({key: data})

        except Exception as e:
            logging.error(
                f"Error al realizar el evento de carga de datos en Google Sheets >>> {e}",
                exc_info=True
            )

    def _event_db_deleted_by_other(
            self,
            data: Dict[str, Any]) -> None:
        """
        Delete row in current instance if other delete it
        """
        try:
            if data:
                for key, value in data.items():
                    uid = key.split("/")[0]

                    # If was deleted by another account or dispatched
                    if value in (-1, 2):
                        self.packed_tab.delete_row_by_uid(uid)
                    elif value == 3:  # Set as pending order in other account
                        self.packed_tab.set_as_pending_by_uid(uid)

        except Exception as e:
            logging.error(
                f"Error al borrar item eliminado/despachado por otra instancia >>> {e}",
                exc_info=True
            )

    def _event_db_update(
            self,
            event: Event,
            uid: str,
            data: Dict[str, Any]) -> None:
        """
        Call if there is an update in the database
        Only packaged (Status 1) can be updated from the tablet

        Args:
            event (Any): Database listener event
            uid (str): Database uid
            data (Dict[str, Any]): Document data
        """
        try:
            if data.get("estado") == 1:
                for key, value in event.data.items():
                    data[key] = value

                table_data = {uid: data.copy()}

                # Emit signal
                self._row_updater.update.emit(data["factura"], table_data)

        except Exception as e:
            logging.error(
                f"Error al actualizar una orden por otra instancia >>> {e}",
                exc_info=True
            )

    def __event_connect(self) -> None:
        """
        Connect server event
        """
        try:
            if self._server.is_disconnected():
                self._server = EmesDispatch()

            self._server.initialize(self.user)

            def callback(event: Event):
                path: list = event.path.split('/')

                if len(path) == 2:
                    key = path[1]

                    if not path[1]:
                        if event.event_type == "put":
                            self._event_db_initial_load(event.data)
                        else:  # In case of patch event
                            self._event_db_deleted_by_other(event.data)
                    else:
                        # Callback logic put (packing) or patch (update from packing)
                        values = self._server.get_values_by_key(key)

                        if utils.all_dict(values, event.data):
                            if event.event_type == "put":
                                self._event_db_packing(key, values)
                            else:  # In case of patch event
                                self._event_db_update(event, key, values)

            # Create listener (searching for Firebase updates)
            self._server.set_listener_callback(callback)

            if self._server.is_connected():
                self.setWindowTitle("Emes: Conectado")

                # Initalize infinite loop to find new files if is main account
                if self.user == "despachos":
                    self.__start_file_processing()

        except Exception as e:
            logging.error(
                f'{e}. Error al conectarse a la base de datos y/o la API de GS',
                exc_info=True
            )

    def __event_disconnect(self) -> bool:
        """
        Disconnect server event
        """
        try:
            if self._server is not None:
                self._server.disconnect()

                # Stop thread
                if self.__stop_file_processing():
                    # Update main window title
                    self.setWindowTitle("Emes: Desconectado")
                    return True

            return False

        except Exception as e:
            logging.error(
                f"Error deteniendo el bot >>> {e}",
                exc_info=True
            )
            return False

    def __event_update_invoiced(self) -> None:
        """
        Update invoiced orders
        """
        try:
            data = self._server.db_query("estado", 0)

            if data and isinstance(data, dict):
                table_data = \
                    {
                        key: value
                        for key, value in data.items()
                    }

                # Emit signal to upload initial orders
                self._invoiced_table_updater.reset.emit(table_data)

        except Exception as e:
            logging.error(
                f"Error al actualizar las órdenes facturadas: {e}",
                exc_info=True
            )

    def __event_update_packed(self) -> None:
        """
        Update packed (1) and pending orders (3)
        """
        try:
            # Get packed and pending orders
            packed_data, pending_data = \
                (
                    self._server.db_query("estado", state)
                    for state in (1, 3)
                )

            # Merged dict data object
            data = {}

            # Check if both objects are dicts
            if utils.all_dict(packed_data, pending_data):
                data = {**pending_data, **packed_data}

            if data:
                table_data = \
                    {
                        key: value
                        for key, value in data.items()
                    }

                # Emit signal to upload initial orders
                self._table_updater.reset.emit(table_data)

        except Exception as e:
            logging.error(
                f"Error al actualizar las órdenes empacadas y pendientes >>> {e}",
                exc_info=True
            )

    def __event_update_files(self) -> bool:
        """
        Updates PDF files click event
        """
        try:
            res = False

            # Check if server is connected
            if self._server and self._server.is_connected():
                res = self._server.pdf_to_db()

            return res

        except Exception as e:
            logging.error(
                f"Error al actualizar la lista de nuevas remisiones facturadas >>> {e}",
                exc_info=True
            )

    def __event_delete_invoiced(self) -> bool:
        """
        Delete invoiced orders event
        """
        try:
            # Delete item from table
            _, list_data = self.invoiced_tab.delete_rows()

            if list_data:
                # Proper format to one API call in DB
                data = \
                    {
                        f'{d[-1]}/estado': -1
                        for d in list_data if d[-1] != ""
                    }

                # Set state to -1 in database
                if data and self._server and self._server.is_connected():
                    self._server.update_state_by_keys(data)

            return True

        except Exception as e:
            logging.error(
                f"Error al eliminar órdenes facturadas >>> {e}",
                exc_info=True
            )
            return False

    def __event_delete_packed(self) -> bool:
        """
        Delete packed orders event
        """
        try:
            # Delete item from table
            _, list_data = self.packed_tab.delete_rows()

            if list_data:
                # Proper format to one API call in DB
                data = \
                    {
                        f'{d[-1]}/estado': -1
                        for d in list_data if d[-1] != ""
                    }

                # Set state to -1 in database
                if data and self._server and self._server.is_connected():
                    self._server.update_state_by_keys(data)

            return True

        except Exception as e:
            logging.error(
                f"Error al eliminar órdenes empacadas/pendientes >>> {e}",
                exc_info=True
            )
            return False

    def __event_update_users(self) -> bool:
        """
        Updates combo box user lists
        """
        try:
            if self._server:
                res = self._server.update_users()
                return bool(res)

            return False

        except Exception as e:
            logging.error(
                f"Error al actualizar los usarios >>> {e}",
                exc_info=True
            )
            return False

    def __event_set_as_pending(self) -> None:
        """
        Set selected order as pending
        """
        try:
            # Get indices and values of selected range
            _, rows_data = self.packed_tab.get_all_values()

            # Convert to DB format
            db_data = \
                {
                    f'{d[-1]}/estado': 3
                    for d in rows_data
                    if d[-1] != ""
                }

            # Set state to 3 (Pendiente)
            if db_data:
                self._server.update_state_by_keys(db_data)

        except Exception as e:
            logging.error(
                f"Error al poner la orden seleccionada como pendiente >>> {e}",
                exc_info=True
            )

    def __event_treasury_report(self) -> bool:
        """
        Generate treasury report
        """
        try:
            filename = self.services_tab.get_date(1)

            # Change button state
            self.change_button_state(False)

            if self._server:
                self._server.generate_report(
                    type="treasury",
                    filename=filename
                )
                return True

            return False

        except Exception as e:
            logging.error(
                f"Error al crear el reporte de Tesorería >>> {e}",
                exc_info=True
            )
            return False

    def __event_dispatches_report(self) -> bool:
        """
        Generate dispatches report
        """
        try:
            filename = self.services_tab.get_date(2)

            if self._server:
                self._server.generate_report(
                    type="dispatches",
                    filename=filename
                )
                return True

            return False

        except Exception as e:
            logging.error(
                f"Error al crear el reporte de Despachos >>> {e}",
                exc_info=True
            )
            return False

    def __event_find(self) -> None:
        """
        Find element event
        """
        try:
            if self._server:
                # Get code and request the query
                code = utils.safe_int(self.services_tab.get_code())

                db_item = self._server.find_item(code)

                if db_item is not None:
                    data: dict = list(db_item.values())[0]

                    # Update date and client text edit
                    self._edit_updater.update.emit(data)

        except Exception as e:
            logging.error(
                f"Error al buscar el elemento solicitado. Revise si está conectado al servidor >>> {e}",
                exc_info=True
            )

    def __event_edit_config(self) -> bool:
        """
        Edit server event
        """
        try:
            # Set modified params
            self._server.params = \
                {
                    "db_url": self.config_tab.edit_1.text(),
                    "folder_id_planillas": self.config_tab.edit_2.text(),
                    "folder_id_remisiones": self.config_tab.edit_3.text(),
                    "json_path_db": self.config_tab.edit_4.text(),
                    "json_path_gs": self.config_tab.edit_5.text()
                }

            res = not any(v is None for v in self._server.params.values())

            return res

        except Exception as e:
            logging.error(
                f'Error editando variables del bot >>> {e}',
                exc_info=True
            )
            return False

    def __event_upload(self) -> bool:
        """
        Sends data to Google Sheets file from QTableWidget UI
        """
        try:
            # Get indices and values of selected range
            index, rows_data = self.packed_tab.get_all_values()

            if not all((index, rows_data)):
                return False

            # Verify if row have valid courier and zone
            valid_rows = \
                {
                    index[i]: row_data
                    for i, row_data in enumerate(rows_data)
                    if row_data[2] and row_data[3]
                }

            # Split keys and values
            keys, values = map(list, zip(*valid_rows.items()))

            # Convert to DB format
            db_data = {}

            for d in values:
                if d[-1] != "" and d[3] != "":
                    db_data[f'{d[-1]}/estado'] = 2
                    db_data[f'{d[-1]}/region'] = d[3]

            # Copy row into GS file
            if self._server.insert_rows_in_sheet(values):
                self.packed_tab.delete_rows_by_index(keys)

                # Set state to 2 (Despachado)
                if db_data:
                    self._server.update_state_by_keys(db_data)

                return True

            return False

        except Exception as e:
            logging.error(
                f"Error al despachar la orden >>> {e}",
                exc_info=True
            )
            return False

    def click_connect(self) -> None:
        """
        Button click method to connect server
        """
        if not self._server.is_connected():
            # Get username and password from edit widgets
            username = self.user
            password = self.password

            if (username in self._server.users and
                    password == self._server.users[username]["password"]):
                if self._server.is_non_connected():
                    self.__thread_connect()
            else:
                self.msg.invalid_account()
        else:
            self.msg.already_logged()

    def click_disconnect(self) -> None:
        """
        Button click method to disconnect server
        """
        if self._server.is_connected():
            self.__thread_disconnect()

    def click_delete_invoiced(self) -> None:
        """
        Button click method to delete items in pending orders table
        """
        res = self.msg.verify_deletion()

        if res == QMessageBox.StandardButton.Ok:
            self.__thread_delete_invoiced()

    def click_delete_packed(self) -> None:
        """
        Button click method to delete items in packed orders table
        """
        res = self.msg.verify_deletion()

        if res == QMessageBox.StandardButton.Ok:
            self.__thread_delete_packed()

    def click_update_invoiced(self) -> None:
        """
        Button click method to update data in InvoicedUi table
        """
        self.__thread_update_invoiced()

    def click_update_packed(self) -> None:
        """
        Button click method to delete items in pending orders table
        """
        res = self.msg.verify_update_orders()

        if res == QMessageBox.StandardButton.Ok:
            self.__thread_update_packed()

    def click_update_files(self) -> None:
        """
        Button click method to update PDF files
        """
        self.__thread_update_files()

    def click_update_users(self) -> None:
        """
        Button click method to update users
        """
        res = self.msg.verify_update_users()

        if res == QMessageBox.StandardButton.Ok:
            if self._server.is_connected():
                # Update GS combo box
                self.__thread_update_users()

                # Update table item widgets
                names = self._server.get_users()

                if names is not None:
                    self.packed_tab.update_users(names)
            else:
                self.msg.non_connection_msg()

    def click_set_as_pending(self) -> None:
        """
        Button click to able pending mode on specified orders
        """
        try:
            self.packed_tab.set_as_pending()

            if self._server.is_connected():
                self.__thread_set_as_pending()
            else:
                self.msg.non_connection_msg()

        except Exception as e:
            logging.error(
                f"Error en el evento click para definir orden como pendiente >>> {e}",
                exc_info=True
            )

    def click_generate_treasury_report(self) -> None:
        """
        Button click to generate treasury report
        """
        if self._server.is_connected():
            self.__thread_generate_treasury_report()
        else:
            self.msg.non_connection_msg()

    def click_generate_dispatches_report(self) -> None:
        """
        Button click to generate dispatches report
        """
        if self._server.is_connected():
            self.__thread_generate_dispatches_report()
        else:
            self.msg.non_connection_msg()

    def click_find(self) -> None:
        """
        Button click method to find documents by code
        """
        if self._server.is_connected():
            self.__thread_find()
        else:
            self.msg.non_connection_msg()

    def click_edit_config(self) -> None:
        """
        Button click method to modify configuration
        """
        self.__thread_edit_config()

    def click_upload(self) -> None:
        """
        Button click method to modify configuration
        """
        res = self.msg.verify_upload()

        if res == QMessageBox.StandardButton.Ok:
            if self._server.is_connected():
                self.__thread_upload()
            else:
                self.msg.non_connection_msg()

    def __thread_connect(self) -> None:
        """
        Add thread to create report Emes
        """
        worker = Worker(self.__event_connect)
        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # Start thread
        self._threadpool.start(worker)

    def __thread_disconnect(self) -> None:
        """
        Add thread to create report Emes
        """
        worker = Worker(self.__event_disconnect)
        worker.signals.result.connect(self.msg.logout_response)
        worker.signals.progress.connect(self.progress_fn)
        worker.signals.finished.connect(self.thread_complete)

        # Start thread
        self._threadpool.start(worker)

    def __thread_upload(self) -> None:
        """
        Add thread to dispatch order
        """
        worker = Worker(self.__event_upload)
        worker.signals.result.connect(self.msg.upload_response)
        worker.signals.progress.connect(self.progress_fn)
        worker.signals.finished.connect(self.thread_complete)

        # start thread
        self._threadpool.start(worker)

    def __thread_delete_invoiced(self) -> None:
        """
        Add thread to create report Emes
        """
        # worker
        worker = Worker(self.__event_delete_invoiced)
        worker.signals.result.connect(self.msg.delete_response)
        worker.signals.progress.connect(self.progress_fn)
        worker.signals.finished.connect(self.thread_complete)

        # start thread
        self._threadpool.start(worker)

    def __thread_delete_packed(self) -> None:
        """
        Add thread to create report Emes
        """
        # worker
        worker = Worker(self.__event_delete_packed)
        worker.signals.result.connect(self.msg.delete_response)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # Start thread
        self._threadpool.start(worker)

    def __thread_set_as_pending(self) -> None:
        """
        Add thread to set order as pending
        """
        worker = Worker(self.__event_set_as_pending)
        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # Start thread
        self._threadpool.start(worker)

    def __thread_update_invoiced(self) -> None:
        """
        Add thread to update invoiced orders table
        """
        # worker
        worker = Worker(self.__event_update_invoiced)
        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # Start thread
        self._threadpool.start(worker)

    def __thread_update_packed(self) -> None:
        """
        Add thread to create report Emes
        """
        # worker
        worker = Worker(self.__event_update_packed)
        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # Start thread
        self._threadpool.start(worker)

    def __thread_update_files(self) -> None:
        """
        Add thread to create report Emes
        """
        # worker
        worker = Worker(self.__event_update_files)
        worker.signals.result.connect(self.msg.update_files_response)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # Start thread
        self._threadpool.start(worker)

    def __thread_update_users(self) -> None:
        """
        Add thread to create report Emes
        """
        worker = Worker(self.__event_update_users)
        worker.signals.result.connect(self.msg.update_users_response)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # Start thread
        self._threadpool.start(worker)

    @pyqtSlot(bool)
    def update_button_state(self, state: bool):
        """
        Update enable button state
        """
        self.services_tab._report_ui.button_1.setEnabled(state)

    def change_button_state(self, state: bool) -> None:
        """
        Update button enable state from main thread

        Args:
            state (bool): Enabled/Disabled signal
        """
        QMetaObject.invokeMethod(
            self,
            "update_button_state",
            Qt.ConnectionType.QueuedConnection,
            Q_ARG(bool, state)
        )

    def __thread_generate_treasury_report(self) -> None:
        """
        Add thread to create report Emes
        """
        def finish_report_thread():
            self.thread_complete()
            self.change_button_state(True)

        worker = Worker(self.__event_treasury_report)
        worker.signals.result.connect(self.msg.report_response)
        worker.signals.finished.connect(finish_report_thread)
        worker.signals.progress.connect(self.progress_fn)

        # Start thread
        self._threadpool.start(worker)

    def __thread_generate_dispatches_report(self) -> None:
        """
        Add thread to create report Emes
        """
        worker = Worker(self.__event_dispatches_report)
        worker.signals.result.connect(self.msg.report_response)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # Start thread
        self._threadpool.start(worker)

    def __thread_find(self) -> None:
        """
        Add thread to create report Emes
        """
        worker = Worker(self.__event_find)
        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # Start thread
        self._threadpool.start(worker)

    def __thread_edit_config(self) -> None:
        """
        Add thread to create report Emes
        """
        worker = Worker(self.__event_edit_config)
        worker.signals.result.connect(self.msg.config_response)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # Start thread
        self._threadpool.start(worker)


def run_ui(server: EmesDispatch) -> None:
    """
    Main UI function to create MainWindow

    Args:
        emes (EmesDispatch): Server object
    """
    app = QApplication(sys.argv)
    window = MainWindow(server)

    with open("public/styles.css", "r") as file:
        style = file.read()

    # Apply QSS style
    app.setStyleSheet(style)

    # Resize main window
    window.resize(
        window.screen_size.width(),
        window.screen_size.height()
    )

    window.show()
    sys.exit(app.exec())
