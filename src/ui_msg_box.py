from typing import Any
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *


class MessageBox:

    def __init__(self, parent: QWidget | None) -> None:
        """
        Constructor
        """
        self._parent = parent

    def __create_msg_box(
            self,
            text: str,
            title: str,
            icon: QMessageBox.Icon = QMessageBox.Icon.Information,
            buttons: QMessageBox.StandardButton | None = None) -> int:
        """
        Create and show QMessageBox

        Args:
            text (str): Message text
            title (str): Title name
            icon (QMessageBox.Icon, optional): Message icon. Defaults to QMessageBox.Icon.Information.
            buttons (QMessageBox.StandardButton | None, optional): Buttons to be added. Defaults to None.

        Returns:
            int: Message return code
        """
        msg = QMessageBox(parent=self._parent, text=text)
        msg.setWindowTitle(title)
        msg.setIcon(icon)
        msg.setStandardButtons(
            buttons if buttons else QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel
        )
        ret = msg.exec()

        return ret

    def invalid_account(self) -> int:
        """
        Message box if user or password does not exist

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="El usuario o la contraseña son incorrectas",
            title="No existe el usuario/contraseña",
            icon=QMessageBox.Icon.Critical
        )

    def already_logged(self) -> int:
        """
        Message indicating that the server is already connected

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="Ya está conectado el servidor",
            title="Conexión",
            icon=QMessageBox.Icon.Warning
        )

    def non_connection_msg(self) -> int:
        """
        Show non-initialization server-object-message-box
        """
        return self.__create_msg_box(
            text="El servidor no está conectado. Revise la conexión a internet o ingrese el usuario/contraseña",
            title="Sin conexión",
            icon=QMessageBox.Icon.Critical
        )

    def _successful_config(self) -> int:
        """
        Message indicating that configuration update was successful

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="Se han actualizado las credenciales de las APIs",
            title="Proceso finalizado"
        )

    def _successful_upload(self) -> int:
        """
        Message indicating that order upload was successful

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="Las órdenes fueron cargadas exitosamente en Google Sheets",
            title="Proceso finalizado"
        )

    def _successful_update(self) -> int:
        """
        Message indicating that list update was successful

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="Se ha actualizado correctamente la lista de remisiones",
            title="Proceso finalizado"
        )

    def _successful_update_files(self) -> int:
        """
        Message indicating that list update was successful

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="Se ha actualizado correctamente la lista de archivos .pdf",
            title="Proceso finalizado"
        )

    def _successful_update_users(self) -> int:
        """
        Message indicating that list update was successful

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="Se ha actualizado correctamente la lista de empacadores/mensajeros",
            title="Proceso finalizado"
        )

    def _successful_report(self) -> int:
        """
        Message indicating that list update was successful

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="Se ha generado correctamente el reporte",
            title="Proceso finalizado"
        )

    def _successful_deletion(self) -> int:
        """
        Message indicating that order deletion was successful

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="Se han eliminado las órdenes correctamente",
            title="Proceso finalizado"
        )

    def _successful_logout(self) -> int:
        """
        Message indicating that logout was successful

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="Desconexión exitosa de la base de datos",
            title="Proceso finalizado"
        )

    def _invalid_config(self) -> int:
        """
        Message indicating that configuration update was invalid

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="No se han podido actualizar las credenciales de las APIs",
            title="Error al actualizar la configuración",
            icon=QMessageBox.Icon.Critical
        )

    def _invalid_upload(self) -> int:
        """
        Message indicating that order upload was upload

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="No fue posible cargar la remisión en Google Sheets",
            title="Error al cargar la orden",
            icon=QMessageBox.Icon.Critical
        )

    def _invalid_update(self) -> int:
        """
        Message indicating that order list update was invalid

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="No fue posible actualizar la lista de remisiones",
            title="Error al actualizar la lista de remisiones",
            icon=QMessageBox.Icon.Critical
        )

    def _invalid_update_files(self) -> int:
        """
        Message indicating that order list update was invalid

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="No fue posible actualizar la lista de archivos .pdf",
            title="Error al actualizar la lista de archivos",
            icon=QMessageBox.Icon.Critical
        )

    def _invalid_update_users(self) -> int:
        """
        Message indicating that order list update was invalid

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="No fue posible actualizar la lista empacadores/mensajeros",
            title="Error al actualizar la lista de usuarios",
            icon=QMessageBox.Icon.Critical
        )

    def _invalid_report(self) -> int:
        """
        Message indicating that order list update was invalid

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="No fue posible generar el reporte",
            title="Error al generar reporte",
            icon=QMessageBox.Icon.Critical
        )

    def _invalid_deletion(self) -> int:
        """
        Message indicating that order deletion was invalid

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="No fue posible borrar la orden",
            title="Error al eliminar la orden",
            icon=QMessageBox.Icon.Critical
        )

    def _invalid_logout(self) -> int:
        """
        Message box if log was not successful

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="No se pudo desconectar de la base de datos",
            title="Proceso finalizado",
            icon=QMessageBox.Icon.Critical
        )

    def verify_deletion(self) -> int:
        """
        Message box to verify if user wants to delete order

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="¿Está seguro de que desea eliminar esta/s remision/es?",
            title="Eliminación de órdenes",
            icon=QMessageBox.Icon.Question
        )

    def verify_upload(self) -> int:
        """
        Message box to verify if user wants to update order list

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="¿Está seguro de que desea cargar una nueva remisión en Google Sheets?",
            title="Cargar despacho",
            icon=QMessageBox.Icon.Question
        )

    def verify_update_users(self) -> int:
        """
        Message box to verify if user wants to update order list

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="¿Está seguro de que desea actualizar la lista de empacadores/mensajeros?",
            title="Actualización de usuarios",
            icon=QMessageBox.Icon.Question
        )

    def verify_update_orders(self) -> int:
        """
        Message box to verify if user wants to update order list

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="¿Está seguro de que desea actualizar la lista de remisiones?",
            title="Actualización de lista de remisiones",
            icon=QMessageBox.Icon.Question
        )

    def unavailable_service(self) -> int:
        """
        Message box to verify if user wants to update order list

        Returns:
            int: Message return code
        """
        return self.__create_msg_box(
            text="El servicio de despachos no está disponible para esta cuenta",
            title="Servicio no disponible",
            icon=QMessageBox.Icon.Critical
        )

    def delete_response(self, res: Any) -> int:
        """
        Args:
            res (Any): Server function response

        Returns:
            int: Message return code
        """
        return self._successful_deletion() if res else self._invalid_deletion()

    def update_response(self, res: Any) -> int:
        """
        Args:
            res (Any): Server function response

        Returns:
            int: Message return code
        """
        return self._successful_update() if res else self._invalid_update()

    def update_files_response(self, res: Any) -> int:
        """
        Args:
            res (Any): Server function response

        Returns:
            int: Message return code
        """
        return self._successful_update_files() if res else self._invalid_update_files()

    def update_users_response(self, res: Any) -> int:
        """
        Args:
            res (Any): Server function response

        Returns:
            int: Message return code
        """
        return self._successful_update_users() if res else self._invalid_update_users()

    def report_response(self, res: Any) -> int:
        """
        Args:
            res (Any): Server function response

        Returns:
            int: Message return code
        """
        return self._successful_report() if res else self._invalid_report()

    def upload_response(self, res: Any) -> int:
        """
        Args:
            res (Any): Server function response

        Returns:
            int: Message return code
        """
        return self._successful_upload() if res else self._invalid_upload()

    def config_response(self, res: Any) -> int:
        """
        Args:
            res (Any): Server function response

        Returns:
            int: Message return code
        """
        return self._successful_config() if res else self._invalid_config()

    def logout_response(self, res: Any) -> int:
        """
        Args:
            res (Any): Server function response

        Returns:
            int: Message return code
        """
        return self._successful_logout() if res else self._invalid_logout()
