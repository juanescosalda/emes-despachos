from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from ui_reports import ReportUi
from ui_find import FindUi
from typing import Callable, Dict, Any


class ServicesUi(QWidget):

    def __init__(self, parent: QWidget):
        """
        Constructor

        Args:
            parent (QWidget): Main window widget
        """
        super().__init__()

        self._parent = parent

        # Create Find order UI
        self._find_ui = FindUi(self)
        self.component_1 = self._find_ui.get_form()
        self.component_1.setFixedHeight(300)
        self.component_1.setObjectName("gbox_services")

        # Create Report order UI
        self._report_ui = ReportUi(self)
        self.component_2 = self._report_ui.get_form()
        self.component_2.setFixedHeight(300)
        self.component_2.setObjectName("gbox_services")

    def get_date(self, report: int) -> str:
        """
        Get date as file name

        Args:
            report (int): Report type (1: Treasury, 2: Dispatches)

        Returns:
            str: File name
        """
        dates = \
            {
                1: self._report_ui.date_edit_1.date().toPyDate(),
                2: self._report_ui.date_edit_2.date().toPyDate()
            }

        if report not in dates:
            raise ValueError(f"Invalid report type: {report}")

        return dates[report].strftime('%Y-%m-%d')

    def get_code(self) -> str:
        """
        Get invoice code
        """
        code = self._find_ui.edit_1.text()

        return code

    def set_edit_by_query(self, data: Dict[str, Any]) -> None:
        """
        Enter the values according to the query

        Args:
            data (Dict[str, Any]): Database data
        """
        if data:
            self._find_ui.edit_2.setText(data["fecha"])
            self._find_ui.edit_3.setText(data["farmacia"])
        else:
            self._find_ui.edit_2.setText("")
            self._find_ui.edit_3.setText("")

    def find_function(self, func: Callable) -> None:
        """
        Find code button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self._find_ui.update_button.clicked.connect(func)

    def create_treasury_report_function(self, func: Callable) -> None:
        """
        Create treasury button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self._report_ui.button_1.clicked.connect(func)

    def create_dispatches_report_function(self, func: Callable) -> None:
        """
        Create treasury button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self._report_ui.button_2.clicked.connect(func)

    def get_form(self) -> QWidget:
        """
        Returns:
            QWidget: Form widget
        """
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.component_1)
        main_layout.addWidget(self.component_2)
        main_layout.setContentsMargins(180, 100, 180, 100)
        main_layout.setSpacing(50)

        main = QWidget()
        main.setLayout(main_layout)
        main.setObjectName("services_tab_widget")

        return main


class EditUpdater(QObject):

    update = pyqtSignal(dict)

    def __init__(self, ui: ServicesUi):
        super().__init__()
        self.ui = ui

    @pyqtSlot(dict)
    def update_edit(self, data: dict):
        """
        Update row data in table

        Args:
            data (dict): Database values as dict
        """
        self.ui.set_edit_by_query(data)
