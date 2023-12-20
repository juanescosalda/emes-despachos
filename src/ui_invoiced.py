import utils
from typing import Callable
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from ui_table import MainTable


class InvoiceUi(MainTable):

    EDIT_HEIGHT = 25
    BUTTON_HEIGHT = 30

    HEADERS = \
        {
            "Factura": 100,
            "Cliente": 450,
            "Facturador": 220,
            "Hora facturación": 140,
            "Artículos": 90,
            "Forma de pago": 200,
            "Valor a cobrar": 120,
            "uid": 0
        }

    def __init__(
            self,
            parent: QWidget | None):
        """
        Constructor

        Args:
            parent (QWidget | None): MainWidget parent
        """
        super().__init__(parent, InvoiceUi.HEADERS)

        self._parent = parent

        # Create all window buttons
        self.create_buttons()

    def create_buttons(self) -> None:
        """
        Override base class method

        Create tab widget buttons
        """
        # Create update users button
        self.update_button = self.create_top_button(
            icon_path="icons/refresh_icon.svg",
            tooltip="Actualizar lista"
        )

        # Create delete row button
        self.delete_button = self.create_top_button(
            icon_path="icons/delete_icon.svg",
            tooltip="Eliminar fila"
        )

        # Create update users button
        self.update_users_button = self.create_top_button(
            icon_path="icons/add_users_icon.svg",
            tooltip="Actualizar empacadores"
        )

        # Create update pdf files button
        self.update_files_button = self.create_top_button(
            icon_path="icons/pdf_icon.svg",
            tooltip="Actualizar PDFs facturados"
        )

    def update_button_function(self, func: Callable) -> None:
        """
        Assign send button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self.update_button.clicked.connect(func)

    def update_users_button_function(self, func: Callable) -> None:
        """
        Assign update button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self.update_users_button.clicked.connect(func)

    def update_files_button_function(self, func: Callable) -> None:
        """
        Assign update button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self.update_files_button.clicked.connect(func)

    def delete_button_function(self, func: Callable) -> None:
        """
        Delete row button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self.delete_button.clicked.connect(func)

    def fill_table(self, db_data: dict) -> None:
        """
        Fill table values

        Args:
            db_data (dict): Database values as dict
        """
        for row, (key, data) in enumerate(db_data.items()):
            # Remisión / Factura
            item = QStandardItem()
            item.setData(data["factura"], 0)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 0, item)

            # Cliente
            item = QStandardItem(str(data["farmacia"]))
            self.model.setItem(row, 1, item)

            # Nombre facturador
            item = QStandardItem(str(data["facturador"]))
            self.model.setItem(row, 2, item)

            # Hora de empaque
            item = QStandardItem(str(data["fecha"]))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 3, item)

            # Artículos
            item = QStandardItem(str(data["articulos"]))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 4, item)

            # Forma de pago
            item = QStandardItem(str(data["forma"]))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 5, item)

            # Valor de la factura
            item = QStandardItem()
            value = utils.to_currency(data["valor"])
            item.setData(value, 0)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 6, item)

            # uid
            item = QStandardItem(key)
            self.model.setItem(row, 7, item)

    def update_row(self, code: int, db_data: dict) -> None:
        """
        Args:
            code (int): Invoice code
            db_data (dict): Database modified values

        Nota:
            Se pone dict en lugar de Dict[] por la firma de la función en el slot
        """
        row = self.__find_row_by_value("Factura", code)
        data = list(db_data.values())[0]

        # Hora de empaque
        item = QStandardItem(str(data["fecha"]))
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.model.setItem(row, 3, item)

    def get_form(self) -> QWidget:
        """
        Get main custom widget

        Returns:
            QWidget: Created widget
        """
        # Upper layout
        top_layout = QHBoxLayout()
        top_layout.addWidget(self.update_button)
        top_layout.addWidget(self.delete_button)
        top_layout.addWidget(self.update_users_button)
        top_layout.addWidget(self.update_files_button)
        top_layout.setSpacing(7)
        top_layout.addStretch()
        top_layout.addWidget(self.query)

        # Bottom layout
        main_layout = QVBoxLayout()
        main_layout.addLayout(top_layout)
        main_layout.addWidget(self.table_view)

        main = QWidget()
        main.setLayout(main_layout)
        main.setObjectName("pending_tab_widget")

        return main
