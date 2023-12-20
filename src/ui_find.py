from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *


class FindUi(QWidget):

    EDIT_HEIGHT = 30

    def __init__(self, parent: QWidget):
        """
        Constructor

        Args:
            parent (QWidget): Main window widget
        """
        super().__init__()

        self._parent = parent

        # Create labels
        self.label_1 = QLabel('Código')
        self.label_2 = QLabel('Fecha')
        self.label_3 = QLabel('Cliente')

        # Set width
        self.label_1.setFixedWidth(50)
        self.label_2.setFixedWidth(50)
        self.label_3.setFixedWidth(50)

        # Set object name
        self.label_1.setObjectName("services_label")
        self.label_2.setObjectName("services_label")
        self.label_3.setObjectName("services_label")

        # Create QLineEdit for "Código"
        self.edit_1 = QLineEdit(parent=self)
        self.edit_1.setFixedHeight(FindUi.EDIT_HEIGHT)
        self.edit_1.setObjectName("services_edit")
        self.edit_1.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.edit_1.setPlaceholderText(
            "Ingrese el código/factura")

        # Create QLineEdit for "Fecha"
        self.edit_2 = QLineEdit(parent=self)
        self.edit_2.setFixedHeight(FindUi.EDIT_HEIGHT)
        self.edit_2.setObjectName("services_edit")
        self.edit_2.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.edit_2.setFocusPolicy(Qt.FocusPolicy.ClickFocus)

        # Create QLineEdit for "Cliente"
        self.edit_3 = QLineEdit(parent=self)
        self.edit_3.setFixedHeight(FindUi.EDIT_HEIGHT)
        self.edit_3.setObjectName("services_edit")
        self.edit_3.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.edit_3.setFocusPolicy(Qt.FocusPolicy.ClickFocus)

        # Create find button
        self.update_button = QPushButton('Buscar', self)
        self.update_button.setFixedHeight(FindUi.EDIT_HEIGHT)
        self.update_button.setFixedSize(80, 30)
        self.update_button.setObjectName("services_button")

    def get_form(self) -> QWidget:
        """
        Returns:
            QWidget: Form widget
        """
        component = QGroupBox("Búsqueda")

        # Horizontal "Código" layout
        hbox_1 = QHBoxLayout()
        hbox_1.addWidget(self.label_1)
        hbox_1.addWidget(self.edit_1)
        hbox_1.addWidget(self.update_button)

        # Horizontal "Fecha" layout
        hbox_2 = QHBoxLayout()
        hbox_2.addWidget(self.label_2)
        hbox_2.addWidget(self.edit_2)

        # Horizontal "Cliente" layout
        hbox_3 = QHBoxLayout()
        hbox_3.addWidget(self.label_3)
        hbox_3.addWidget(self.edit_3)

        # Set find layout
        main_layout = QVBoxLayout()
        main_layout.addLayout(hbox_1)
        main_layout.addLayout(hbox_2)
        main_layout.addLayout(hbox_3)

        component.setLayout(main_layout)

        return component
