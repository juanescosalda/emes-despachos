import utils
from typing import Any, List, Dict, Tuple
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *


class MainTable(QWidget):

    EDIT_HEIGHT = 25
    BUTTON_HEIGHT = 30

    def __init__(
            self,
            parent: QWidget | None,
            settings: Dict[str, int]):
        """
        Constructor

        Args:
            parent (QWidget | None): MainWidget parent
        """
        super().__init__()

        # Get general settings
        self._headers, widths = map(list, zip(*settings.items()))

        # Member attributes
        self._parent = parent
        self._cols = len(self._headers)

        # Create  graphical objects
        self.table_view = self.__create_table_view()
        self.query = self.__create_query_bar()

        # Create data managing objects
        self.model = self.__create_model()
        self.proxy_model = self.__create_sort_proxy_model()

        # Set main proxy model
        self.table_view.setModel(self.proxy_model)

        # Set table settings
        self.__set_settings(widths)

        # Format currency cell after edit event
        self.model.dataChanged.connect(self.handle_item_changed)

    def create_buttons(self) -> None:
        """
        Virtual method
        """
        raise NotImplementedError("Subclasses must implement this method.")

    def fill_table(self, db_data: dict) -> None:
        """
        Fill table values

        Args:
            db_data (dict): Database values as dict
        """
        raise NotImplementedError(
            f"Subclasses must implement this method. Data: {db_data}")

    def update_row(self, code: int, db_data: dict) -> None:
        """
        Fill table values

        Args:
            code (int): Invoice code
            db_data (dict): Database values as dict
        """
        raise NotImplementedError(
            f"Subclasses must implement this method. Code: {code} | Data: {db_data}")

    def __create_model(self) -> QStandardItemModel:
        """
        Create standard item model to manage table view data

        Returns:
            QStandardItemModel: Model to storing and manage data
        """
        model = QStandardItemModel(0, self._cols)
        model.setHorizontalHeaderLabels(self._headers)

        return model

    def __create_table_view(self) -> QTableView:
        """
        Create main table view object

        Returns:
            QTableView: Table view object
        """
        table_view = QTableView()
        table_view.resizeColumnsToContents()
        table_view.setShowGrid(False)
        table_view.setSizeAdjustPolicy(
            QAbstractScrollArea.SizeAdjustPolicy.AdjustToContents
        )

        return table_view

    def __create_sort_proxy_model(self) -> QSortFilterProxyModel:
        """
        Create sort proxy model

        Returns:
            QSortFilterProxyModel: Model for sorting and filtering data
        """
        proxy_model = QSortFilterProxyModel()
        proxy_model.setSourceModel(self.model)
        proxy_model.setDynamicSortFilter(True)

        return proxy_model

    def __create_query_bar(self) -> QLineEdit:
        """
        Create QLineEdit query bar

        Returns:
            QLineEdit: Query bar for filter purposes
        """
        query = QLineEdit(self)
        query.setPlaceholderText("Buscar...")
        query.setObjectName("pending_edit")
        query.textChanged.connect(self.__filter_table)
        query.setFixedHeight(25)
        query.setClearButtonEnabled(True)
        query.addAction(
            QIcon("icons/search_mi.png"),
            QLineEdit.ActionPosition.LeadingPosition
        )

        return query

    def __set_settings(self, widths: List[int]):
        """
        Set table visual settings
        """
        # Set headers settings
        self.__set_headers_settings()

        # Set columns width
        self.__set_columns_width(widths)

    def __set_headers_settings(self) -> None:
        """
        Set vertical and horizontal table view settings
        """
        v_header = self.table_view.verticalHeader()
        v_header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        v_header.setObjectName("pending_vertical_header")
        v_header.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )

        # Set horizontal header format
        h_header = self.table_view.horizontalHeader()
        h_header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        h_header.setObjectName("pending_horizontal_header")

    def __set_column_width(self, col: int, width: int) -> None:
        """
        Set specified column width

        Args:
            col (int): Column index
            width (int): Width size
        """
        self.table_view.setColumnWidth(col, width)

    def __set_columns_width(self, widths: List[int]) -> None:
        """
        Set all columns width

        Args:
            widths (List[int]): List with widths values
        """
        for col in range(self._cols):
            self.__set_column_width(col, widths[col])

    def selected_rows(self) -> List[int]:
        """
        Get current selected row

        Returns:
            List[int]: Row indices
        """
        selection_model = self.table_view.selectionModel()

        # Check if selection model is not None
        if selection_model:
            indices = selection_model.selectedRows()
        else:
            indices = []

        # Get the row numbers of the selected rows
        selected_rows = \
            [
                index.row()
                for index in indices
            ]

        return selected_rows

    def set_text(self, row: int, col: int, value: Any) -> None:
        """

        """
        item = self.model.item(row, col)

        if item:
            item.setText(str(value))

    def __filter_table(self, text: Any, col: int = 0):
        """
        Filter table by specified text

        Args:
            text (Any): Text to find
            col (int): Column index
        """
        regex = QRegularExpression(
            text,
            QRegularExpression.PatternOption.CaseInsensitiveOption
        )
        self.proxy_model.setFilterRegularExpression(regex)
        self.proxy_model.setFilterKeyColumn(col)

    def create_top_button(
            self,
            icon_path: str,
            tooltip: str) -> QToolButton:
        """
        Create top buttons

        Args:
            icon_path (str): Path to button icon
            tooltip (str): Tooltip text

        Returns:
            QToolButton: Top main buttons
        """
        tool_button = QToolButton(self)
        tool_button.setFixedSize(
            50, MainTable.BUTTON_HEIGHT)
        tool_button.setObjectName("pending_button")
        tool_button.setIcon(QIcon(icon_path))
        tool_button.setToolTip(tooltip)
        tool_button.setIconSize(
            tool_button.fontMetrics().boundingRect("A").size() * 3)

        return tool_button

    def handle_item_changed(self, index: QModelIndex):
        """
        Item changed event function (for currency cell)
        """
        row = index.row()
        col = index.column()

        item = self.model.item(row, col)

        if item:
            # Verify if column is the fifth (Valor de factura)
            if col == 5:
                text = item.text()

                if text.find("$") != -1:
                    value = utils.from_currency(text)
                else:
                    value = utils.safe_int(text)

                if value >= 0:
                    formatted_value = utils.to_currency(value)
                    item.setText(formatted_value)

    def _find_row_by_value(
            self,
            key: str,
            value_to_compare: Any) -> int:
        """
        Find row number by specified invoice code

        Args:
            key (str): Column name
            value_to_compare (Any): Value to compare

        Returns:
            int: Row that has the specified value
        """
        code_row = -1
        code_col = self._headers.index(key)

        for row in range(self.model.rowCount()):
            data = self.item_data(row, code_col)

            if data is not None and data == str(value_to_compare):
                code_row = row
                break

        return code_row

    def get_selected_row_data(
            self,
            row: int | None = None) -> List[Any]:
        """
        Get data of specified row
        """
        if row is None:
            rows = self.selected_rows()
            row = rows[0] if len(rows) > 0 else -1

        # Get data of selected row
        data = self.__get_row_data(row, self._cols) if row != -1 else []

        return data

    def delete_row_by_uid(self, uid: str) -> None:
        """
        Deletes specified row by given uid
        """
        row = self._find_row_by_value("uid", uid)

        # Delete selected row
        self.__delete_row(row)

    def delete_row_by_order(self, order: str) -> None:
        """
        Deletes specified row by given order id
        """
        row = self._find_row_by_value("Orden", order)

        # Delete selected row
        self.__delete_row(row)

    def move_row(self, db_data: dict) -> None:
        """
        Args:
            db_data (dict): Database modified values

        Nota:
            Se pone dict en lugar de Dict[] por la firma de la función en el slot
        """
        row = self._find_row_by_value("Orden", db_data["order"])

        if row >= 0:
            self.__delete_row(row)

    def delete_all_rows(self) -> None:
        """
        Delete all table widget rows
        """
        self.model.setRowCount(0)

    def __delete_row(self, row: int) -> None:
        """
        Delete specified row
        """
        self.proxy_model.removeRow(row)

    def delete_rows(self) -> Tuple[List[int], List[List[Any]]]:
        """
        Delete selected rows
        """
        cols = self.model.columnCount()

        # Get all the selected rows
        selected_rows = self.selected_rows()

        # Get values of all rows
        data = \
            [
                self.__get_row_data(row, cols)
                for row in selected_rows
            ]

        # Delete selected rows
        self.delete_rows_by_index(selected_rows)

        return selected_rows, data

    def delete_rows_by_index(
            self,
            indices: List[int]) -> None:
        """
        Delete selected rows based on indices
        """
        for row in sorted(indices, reverse=True):
            self.__delete_row(row)

    def item_data(self, row: int, col: int) -> Any | None:
        """
        Get item data

        Args:
            row (int): Row index
            col (int): Column index

        Returns:
            Any | None: Item data
        """
        index = self.proxy_model.index(row, col)

        res = None

        if index:
            data = self.proxy_model.itemData(index)

            if data and isinstance(data, dict):
                values = list(data.values())

                if values is not None and len(values) > 0:
                    res = values[0]

        return res

    def __get_row_data(
            self,
            row: int,
            cols: int) -> List[Any]:
        """
        Get row data

        Args:
            row (int): Row index
            cols (int): Column index

        Returns:
            List[Any]: Row data as list
        """
        row_data = []

        # Iterate through all columns
        for col in range(cols):
            data = self.item_data(row, col)

            if data is not None:
                row_data.append(data)

        return row_data

    def get_all_values(self) -> Tuple[List[int], List[List]]:
        """
        Get all selected values
        """
        cols = self.model.columnCount()

        # Get all the selected rows
        selected_rows = self.selected_rows()

        # Get values of all rows
        data = \
            [
                self.__get_row_data(row, cols)
                for row in selected_rows
            ]

        return selected_rows, data
