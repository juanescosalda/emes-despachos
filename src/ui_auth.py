from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from typing import Callable
import logo_emes


class LoginUi(QWidget):

    EDIT_HEIGHT = 35
    BUTTON_HEIGHT = 30

    def __init__(self, parent: QWidget | None = None):
        """
        Constructor

        Args:
            parent (QWidget): Main window widget
        """
        super().__init__()

        self._parent = parent

        # Logo label
        self.logo_label = QLabel(self)
        self.logo_label.setFixedSize(400, 150)
        self.logo_label.setObjectName("label_auth_logo")

        # Create user component
        self.username_edit = QLineEdit(self)
        self.username_edit.setPlaceholderText('Ingrese el usuario')
        self.username_edit.setObjectName("auth_user_edit")
        self.username_edit.setFixedHeight(LoginUi.EDIT_HEIGHT)
        self.username_edit.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.username_edit.addAction(
            QIcon("icons/person_icon.png"),
            QLineEdit.ActionPosition.LeadingPosition
        )

        # Create password component
        self.password_edit = QLineEdit(self)
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_edit.setPlaceholderText('Ingrese la constraseña')
        self.password_edit.setObjectName("auth_password_edit")
        self.password_edit.setFixedHeight(LoginUi.EDIT_HEIGHT)
        self.password_edit.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.password_edit.addAction(
            QIcon("icons/lock_icon.png"),
            QLineEdit.ActionPosition.LeadingPosition
        )

        # Create connect button
        self.login_button = QPushButton("Conectar", self)
        self.login_button.setFixedHeight(LoginUi.BUTTON_HEIGHT)
        self.login_button.setObjectName("auth_login_button")

        # Create disconnect button
        self.logout_button = QPushButton("Desconectar", self)
        self.logout_button.setFixedHeight(LoginUi.BUTTON_HEIGHT)
        self.logout_button.setObjectName("auth_logout_button")

        # Hide/Show password action
        icon_on = QIcon('icons/eye_on.png')
        self.show_password_action = QAction(
            icon_on, 'Show password', self.logout_button)

        self.password_edit.addAction(
            self.show_password_action,
            QLineEdit.ActionPosition.TrailingPosition
        )

        self.show_password_action.setCheckable(True)
        self.show_password_action.toggled.connect(self.show_password)

    def login_function(self, func: Callable) -> None:
        """
        Connect button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self.login_button.clicked.connect(func)

    def logout_function(self, func: Callable) -> None:
        """
        Disconnect button clicked event function

        Args:
            func (Callable): Function to execute
        """
        self.logout_button.clicked.connect(func)

    def get_form(self) -> QWidget:
        """
        Returns:
            QWidget: Login widget
        """
        component = QGroupBox()

        #  Horizontal layout
        hbox = QHBoxLayout()
        hbox.addWidget(self.login_button)
        hbox.addWidget(self.logout_button)

        # Vertical layout
        vbox = QVBoxLayout()
        vbox.setSpacing(30)
        vbox.addWidget(self.logo_label, 0, Qt.AlignmentFlag.AlignHCenter)
        vbox.addWidget(self.username_edit)
        vbox.addWidget(self.password_edit)
        vbox.addLayout(hbox)
        vbox.setAlignment(Qt.AlignmentFlag.AlignCenter)

        component.setLayout(vbox)

        main_layout = QVBoxLayout()
        main_layout.addWidget(component)
        main_layout.setContentsMargins(350, 150, 350, 150)

        main = QWidget(parent=self._parent)
        main.setLayout(main_layout)
        main.setObjectName("auth_tab_widget")

        return main

    def show_password(self, show):
        self.password_edit.setEchoMode(
            QLineEdit.EchoMode.Normal if show else QLineEdit.EchoMode.Password
        )
