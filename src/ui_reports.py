from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *


class ReportUi(QWidget):

    WIDGET_HEIGHT = 30
    EDIT_WIDTH = 450

    def __init__(self, parent: QWidget):
        """
        Constructor

        Args:
            parent (QWidget): Main window widget
        """
        super().__init__()

        self._parent = parent

        # Create labels
        self.label_1 = QLabel("Tesorería", parent=self)
        self.label_1.setFixedSize(60, ReportUi.WIDGET_HEIGHT)
        self.label_1.setObjectName("services_label")

        self.label_2 = QLabel("Despachos", parent=self)
        self.label_2.setFixedSize(60, ReportUi.WIDGET_HEIGHT)
        self.label_2.setObjectName("services_label")

        # Get current date object
        today = QDate.currentDate()

        # Create QLineEdit for "Tesorería" report
        self.date_edit_1 = QDateEdit(parent=self)
        self.date_edit_1.setDisplayFormat("yyyy-MM-dd")
        self.date_edit_1.setCalendarPopup(True)
        self.date_edit_1.setFixedSize(
            ReportUi.EDIT_WIDTH,
            ReportUi.WIDGET_HEIGHT
        )
        self.date_edit_1.setDate(today)
        self.date_edit_1.setObjectName("date_edit_services")
        self.date_edit_1.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.date_edit_1.calendarWidget().setFixedHeight(300)

        # Create QLineEdit for "Despachos" report
        self.date_edit_2 = QDateEdit(parent=self)
        self.date_edit_2.setDisplayFormat("yyyy-MM-dd")
        self.date_edit_2.setCalendarPopup(True)
        self.date_edit_2.setFixedSize(
            ReportUi.EDIT_WIDTH,
            ReportUi.WIDGET_HEIGHT
        )
        self.date_edit_2.setDate(today)
        self.date_edit_2.setObjectName("date_edit_services")
        self.date_edit_2.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.date_edit_2.calendarWidget().setFixedHeight(300)

        # Create request "Tesorería" button
        self.button_1 = QPushButton("Solicitar", self)
        self.button_1.setFixedSize(80, ReportUi.WIDGET_HEIGHT)
        self.button_1.setObjectName("services_button")

        # Create request "Despachos" button
        self.button_2 = QPushButton("Solicitar", self)
        self.button_2.setFixedSize(80, ReportUi.WIDGET_HEIGHT)
        self.button_2.setObjectName("services_button")

    def get_form(self) -> QGroupBox:
        """
        Returns:
            QWidget: Reports widget
        """
        component = QGroupBox("Reportes")

        # Horizontal "Código" layout
        hbox_1 = QHBoxLayout()
        hbox_1.addStretch()
        hbox_1.addWidget(self.label_1)
        hbox_1.addWidget(self.date_edit_1)
        hbox_1.addWidget(self.button_1)
        hbox_1.addStretch()

        # Horizontal "Fecha" layout
        hbox_2 = QHBoxLayout()
        hbox_2.addStretch()
        hbox_2.addWidget(self.label_2)
        hbox_2.addWidget(self.date_edit_2)
        hbox_2.addWidget(self.button_2)
        hbox_2.addStretch()

        # Set main layout
        main_layout = QVBoxLayout()
        main_layout.addLayout(hbox_1)
        main_layout.addLayout(hbox_2)

        component.setLayout(main_layout)

        return component
