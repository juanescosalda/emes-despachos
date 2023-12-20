from PyQt6.QtWidgets import *
from PyQt6.QtCore import *


class TableUpdater(QObject):

    update = pyqtSignal(dict)
    reset = pyqtSignal(dict)
    add = pyqtSignal(dict)

    def __init__(self, ui: QWidget):
        super().__init__()
        self.ui = ui

    @pyqtSlot(dict)
    def add_row(self, data: dict):
        """
        Update date in table

        Args:
            data (dict): Database values as dict
        """
        self.ui.add_row(data, None)

    @pyqtSlot(dict)
    def update_table(self, data: dict):
        """
        Update date in table

        Args:
            data (dict): Database values as dict
        """
        self.ui.fill_table(data)

    @pyqtSlot(dict)
    def reset_table(self, data: dict):
        """
        Update date in table

        Args:
            data (dict): Database values as dict
        """
        # Delete al QTableWidgetItem
        self.ui.delete_all_rows()

        # Fill table with new values
        self.ui.fill_table(data)


class RowUpdater(QObject):

    update = pyqtSignal(int, dict)

    def __init__(self, ui: QWidget):
        super().__init__()
        self.ui = ui

    @pyqtSlot(int, dict)
    def update_row(self, code: int, data: dict):
        """
        Update row data in table

        Args:
            data (dict): Database values as dict
        """
        self.ui.update_row(code, data)
