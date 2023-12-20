import utils
import logging
from typing import Any, List, Dict, Callable
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from ui_table import MainTable


class ComboBoxDelegate(QStyledItemDelegate):

    def __init__(
            self,
            parent: QWidget | None = None,
            items: List[Any] = [],
            index: int = 0):
        """
        Constructor

        Args:
            parent (QWidget | None, optional): Widget parent. Defaults to None.
            items (List[Any], optional): Item list of combo box. Defaults to [].
            index (int, optional): Current index of combo box
        """
        super().__init__(parent)

        self._items = items
        self._index = index

    def createEditor(self, parent: QWidget, items: Any, index: Any):
        editor = QComboBox(parent)
        editor.addItems(self._items)
        editor.setCurrentIndex(self._index)
        return editor

    def setEditorData(self, editor, index):
        text = index.data()
        editor.setCurrentText(text)

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText())

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)

    def removeEditor(self, editor, index):
        editor.deleteLater()


class PackedUi(MainTable):

    EDIT_HEIGHT = 25
    BUTTON_HEIGHT = 30

    HEADERS = \
        {
            "Factura": 70,
            "Cliente": 300,
            "Mensajero": 220,
            "Zona": 90,
            "Forma de pago": 150,
            "Valor a cobrar": 120,
            "Cajas": 50,
            "Bolsas": 50,
            "Lios": 50,
            "Facturador": 180,
            "Empacador": 180,
            "Hora empaque": 90,
            "uid": 0
        }

    def __init__(
            self,
            parent: QWidget | None,
            couriers: List[str] = [],
            zones: List[str] = []):
        """
        Constructor

        Args:
            parent (QWidget | None): MainWidget parent
            couriers (List[str]): Courier names
            zones (List[str]): Regions where to distribute the goods
        """
        super().__init__(parent, PackedUi.HEADERS)

        # Member attributes
        self._couriers = couriers
        self._zones = zones

        # Create all window buttons
        self.create_buttons()

        # Set combo box delegates
        combo_box_couriers = ComboBoxDelegate(self, couriers)
        self.table_view.setItemDelegateForColumn(2, combo_box_couriers)

        combo_box_regions = ComboBoxDelegate(self, zones)
        self.table_view.setItemDelegateForColumn(3, combo_box_regions)

    @property
    def couriers(self) -> List:
        return self._couriers

    @couriers.setter
    def couriers(self, in_couriers: List) -> None:
        self._couriers = in_couriers

    @property
    def zones(self) -> List:
        return self._zones

    @zones.setter
    def zones(self, in_zones: List) -> None:
        self._zones = in_zones

    def create_buttons(self) -> None:
        """
        Override base class method

        Create tab widget buttons
        """
        # Create upload data button
        self.send_button = self.create_top_button(
            icon_path="icons/share_icon.svg",
            tooltip="Cargar despacho"
        )

        # Create add row button
        self.add_button = self.create_top_button(
            icon_path="icons/add_icon.svg",
            tooltip="Insertar nueva fila"
        )
        self.add_button.clicked.connect(self.add_empty_row)

        # Create delete row button
        self.delete_button = self.create_top_button(
            icon_path="icons/delete_icon.svg",
            tooltip="Eliminar fila"
        )

        # Create update users button
        self.update_users_button = self.create_top_button(
            icon_path="icons/add_users_icon.svg",
            tooltip="Actualizar usuarios"
        )

        # Create refresh orders button
        self.update_button = self.create_top_button(
            icon_path="icons/refresh_icon.svg",
            tooltip="Actualizar lista"
        )

        # Create setting pending orders button
        self.pending_order_button = self.create_top_button(
            icon_path="icons/pending_icon.svg",
            tooltip="Seleccionar como pendiente"
        )

    def __set_vheader_format(self, row: int):
        """
        Set vertical header format
        """
        icon = QIcon("icons/pace_icon.svg")
        header_item = QStandardItem()
        header_item.setIcon(icon)

        size = QSize(
            icon.pixmap(10).width(),
            icon.pixmap(10).height()
        )
        header_item.setSizeHint(size)
        header_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

        # Set item
        self.model.setVerticalHeaderItem(row, header_item)

    def __set_item_format_by_index(self, row: int, col: int) -> None:
        """
        Args:
            row (int): Row index
            col (int): Column index
        """
        item = self.model.item(row, col)

        if item:
            item.setForeground(QColor(192, 192, 192))
            item.setFont(
                QFont("Helvetica, sans-serif", 8, QFont.Weight.Normal)
            )

    def __set_item_format(self, item: QStandardItem) -> None:
        """
        Args:
            row (int): Row index
            col (int): Column index
        """
        if item:
            item.setForeground(QColor(192, 192, 192))
            item.setFont(
                QFont("Helvetica, sans-serif", 8, QFont.Weight.Normal)
            )

    def set_as_pending(self) -> None:
        """
        Change rows color based on pending orders
        """
        try:
            columns = self.model.columnCount()
            rows = self.selected_rows()

            for row in rows:
                self.__set_vheader_format(row)

                for col in range(columns):
                    self.__set_item_format_by_index(row, col)

        except Exception as e:
            logging.error(
                f"Error al definir un despacho como pendiente  >>> {e}",
                exc_info=True
            )

    def set_as_pending_by_uid(self, uid: str) -> None:
        """
        Change rows color based on pending orders
        """
        try:
            columns = self.model.columnCount()
            row = self._find_row_by_value("uid", uid)

            if row >= 0:
                self.__set_vheader_format(row)

                for col in range(columns):
                    self.__set_item_format_by_index(row, col)

        except Exception as e:
            logging.error(
                f"Error al definir un despacho como pendiente basado en el uid  >>> {e}",
                exc_info=True
            )

    def update_users(self, names: Dict[str, List[str]]) -> None:
        """
        Update users in QComboBox delegates
        """
        combo_box_couriers = ComboBoxDelegate(self, names["couriers"])

        self.table_view.setItemDelegateForColumn(2, combo_box_couriers)

    def send_button_function(self, func: Callable) -> None:
        """
        Assign send button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self.send_button.clicked.connect(func)

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

    def delete_button_function(self, func: Callable) -> None:
        """
        Delete row button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self.delete_button.clicked.connect(func)

    def pending_order_button_function(self, func: Callable) -> None:
        """
        Delete row button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self.pending_order_button.clicked.connect(func)

    def add_row(
            self,
            row_data: Dict[str, Any],
            row: int | None) -> None:
        """
        Fill specified row with given data

        Args:
            dict_data (Dict[str, Any]): Row data
            row (int | None): Row index. None if add in last row
        """
        if not row_data:
            return None

        row = self.proxy_model.rowCount() if row is None else row

        key, data = next(iter(row_data.items()))

        if data and isinstance(data, dict):
            # Remisión / Factura
            item = QStandardItem()
            item.setData(data["factura"], 0)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 0, item)

            # Cliente
            item = QStandardItem(str(data["farmacia"]))
            self.model.setItem(row, 1, item)

            # Mensajeros (QComboBox delegate)
            item = QStandardItem()
            self.model.setItem(row, 2, item)

            # Zona / Región (QComboBox delegate)
            item = QStandardItem()
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 3, item)

            # Forma de pago
            item = QStandardItem(str(data["forma"]))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 4, item)

            # Valor de la factura
            item = QStandardItem()
            value = utils.to_currency(data["valor"])
            item.setData(value, 0)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 5, item)

            # Número de cajas
            item6 = QStandardItem()
            item6.setData(data["cajas"], 0)
            item6.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 6, item6)

            # Número de bolsas
            item = QStandardItem()
            item.setData(data["bolsas"], 0)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 7, item)

            # Número de lios
            item = QStandardItem()
            item.setData(data["lios"], 0)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 8, item)

            # Nombre facturador
            item = QStandardItem(str(data["facturador"]))
            self.model.setItem(row, 9, item)

            # Nombre empacador
            item = QStandardItem(str(data["empacador"]))
            self.model.setItem(row, 10, item)

            # Hora de empaque
            item = QStandardItem(str(data["fecha"]))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.setItem(row, 11, item)

            # uid
            item = QStandardItem(key)
            self.model.setItem(row, 12, item)

    def fill_table(self, db_data: dict) -> None:
        """
        Fill table values

        Args:
            db_data (dict): Database values as dict
        """
        try:
            for row, (key, data) in enumerate(db_data.items()):
                self.add_row({key: data}, row)

                if data["estado"] == 3:
                    self.__set_vheader_format(row)

                    for col in range(self.model.columnCount()):
                        item = self.model.item(row, col)
                        self.__set_item_format(item)

        except Exception as e:
            logging.error(
                f"Error al llenar la tabla de órdenes empacadas  >>> {e}",
                exc_info=True
            )

    def update_row(self, code: int, db_data: dict) -> None:
        """
        Args:
            code (int): Invoice code
            db_data (dict): Database modified values

        Nota:
            Se pone dict en lugar de Dict[] por la firma de la función en el slot
        """
        try:
            row = self._find_row_by_value("Factura", code)

            if row >= 0:
                data = list(db_data.values())[0]

                # Número de cajas
                item = QStandardItem()
                item.setData(data["cajas"], 0)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.model.setItem(row, 6, item)

                # Número de bolsas
                item = QStandardItem()
                item.setData(data["bolsas"], 0)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.model.setItem(row, 7, item)

                # Número de lios
                item = QStandardItem()
                item.setData(data["lios"], 0)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.model.setItem(row, 8, item)

                # Nombre empacador
                item = QStandardItem(str(data["empacador"]))
                self.model.setItem(row, 10, item)

                # Hora de empaque
                item = QStandardItem(str(data["fecha"]))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.model.setItem(row, 11, item)

        except Exception as e:
            logging.error(
                f"Error al actualizar la fila especificada  >>> {e}",
                exc_info=True
            )

    def add_empty_row(self) -> None:
        """
        Create new empty row to add new data
        """
        row = self.model.rowCount()

        # Remisión / Factura
        item = QStandardItem()
        item.setData(0, 0)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.model.setItem(row, 0, item)

        # Cliente
        item = QStandardItem("")
        self.model.setItem(row, 1, item)

        # Mensajeros (QComboBox delegate)
        item = QStandardItem()
        self.model.setItem(row, 2, item)

        # Zona / Región (QComboBox delegate)
        item = QStandardItem()
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.model.setItem(row, 3, item)

        # Forma de pago
        item = QStandardItem("")
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.model.setItem(row, 4, item)

        # Valor de la factura
        item = QStandardItem()
        value = utils.to_currency(0)
        item.setData(value, 0)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.model.setItem(row, 5, item)

        # Número de cajas
        item = QStandardItem()
        item.setData(0, 0)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.model.setItem(row, 6, item)

        # Número de bolsas
        item = QStandardItem()
        item.setData(0, 0)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.model.setItem(row, 7, item)

        # Número de lios
        item = QStandardItem()
        item.setData(0, 0)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.model.setItem(row, 8, item)

        # Nombre facturador
        item = QStandardItem("")
        self.model.setItem(row, 9, item)

        # Nombre empacador
        item = QStandardItem("")
        self.model.setItem(row, 10, item)

        # Hora de empaque
        item = QStandardItem("")
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.model.setItem(row, 11, item)

        # uid
        item = QStandardItem("")
        self.model.setItem(row, 12, item)

    def get_form(self) -> QWidget:
        """
        Get main custom widget

        Returns:
            QWidget: Created widget
        """
        # Upper layout
        top_layout = QHBoxLayout()
        top_layout.addWidget(self.send_button)
        top_layout.addWidget(self.add_button)
        top_layout.addWidget(self.delete_button)
        top_layout.addWidget(self.update_button)
        top_layout.addWidget(self.update_users_button)
        top_layout.addWidget(self.pending_order_button)
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
