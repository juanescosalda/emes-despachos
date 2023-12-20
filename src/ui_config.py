from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from typing import Callable


class ConfigUi(QWidget):

    EDIT_HEIGHT = 35

    def __init__(self, parent: QWidget):
        """
        Constructor

        Args:
            parent (QWidget): MainWindow widget
        """
        super().__init__()

        self._parent = parent

        # Create update button
        self.update_button = QPushButton('Actualizar', self)
        self.update_button.setFixedSize(150, ConfigUi.EDIT_HEIGHT)
        self.update_button.setObjectName("config_button")

        # Create edit text
        self.edit_1 = self.__create_line_edit(
            placeholder="URL base de datos")
        self.edit_2 = self.__create_line_edit(
            placeholder="ID carpeta Planillas")
        self.edit_3 = self.__create_line_edit(
            placeholder="ID carpeta Remisiones")
        self.edit_4 = self.__create_line_edit(
            placeholder="Ruta .json Base de datos")
        self.edit_5 = self.__create_line_edit(
            placeholder="Ruta .json API Gspread")

    def __create_line_edit(self, placeholder: str) -> QLineEdit:
        """
        Create edit widgets

        Args:
            placeholder (str): Placeholder text

        Returns:
            QLineEdit: Object
        """
        # Create line edit widget
        edit = QLineEdit(parent=self)
        edit.setFixedHeight(ConfigUi.EDIT_HEIGHT)
        edit.setPlaceholderText(placeholder)
        edit.setObjectName("config_edit")
        edit.setFocusPolicy(Qt.FocusPolicy.ClickFocus)

        return edit

    def update_button_function(self, func: Callable) -> None:
        """
        update settings button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self.update_button.clicked.connect(func)

    def get_form(self) -> QWidget:
        """
        Returns:
            QWidget: Configuration widget
        """
        component = QGroupBox()

        panel = QVBoxLayout()
        panel.addWidget(self.edit_1)
        panel.addWidget(self.edit_2)
        panel.addWidget(self.edit_3)
        panel.addWidget(self.edit_4)
        panel.addWidget(self.edit_5)
        panel.setSpacing(40)

        hbox = QHBoxLayout()
        hbox.addStretch()
        hbox.addWidget(self.update_button)
        hbox.addStretch()

        vbox = QVBoxLayout()
        vbox.addLayout(panel)
        vbox.addLayout(hbox)

        component.setLayout(vbox)

        main_layout = QVBoxLayout()
        main_layout.addWidget(component)
        main_layout.setContentsMargins(200, 75, 200, 75)

        main = QWidget()
        main.setLayout(main_layout)
        main.setObjectName("config_tab_widget")

        return main
