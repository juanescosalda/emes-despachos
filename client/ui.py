# -*- coding: utf-8 -*-
"""
Created on Wed Dec 14 17:31:01 2022

@author: juane
"""

# %% PyQT UI test
import os
import sys
sys.path.append('.')

import logging
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.uic import loadUi
from server.dispatches import EmesDispatch
from client.worker import Worker
import client.logo_emes as logo_emes


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        loadUi("./client/gui.ui", self)

        # Init threadpool
        self.threadpool = QThreadPool()

        # Main buttons
        self.connectButton.pressed.connect(
            self.click_connect
        )
        self.editVarsButton.clicked.connect(
            self.click_edit
        )
        self.reportPartialButton.clicked.connect(
            self.click_generate_treasury_report
        )
        self.distpachesReportButton.clicked.connect(
            self.click_generate_dispatches_report
        )
        self.reportFinalButton.clicked.connect(
            self.click_generate_main_report
        )

        # Initialize EmesDispatch object
        self.emes = None

        # Set line edits text
        self.__set_line_edits_text()

        # Set today as default date
        today = QDate.currentDate()

        self.dateEditTreasury.setDate(today)
        self.dateEditDispatches.setDate(today)
        self.dateEditReport.setDate(today)

    def __set_line_edits_text(self) -> None:
        """
        Set text of QLineEdit objects
        """
        self.editURLDB.setText(
            'https://emes-empaques-default-rtdb.firebaseio.com/'
        )

        self.editIDPlanillas.setText(
            '1S4mTxoIFIKNibVKRvofJCyLTGeOW0ye4'
        )

        self.editIDRemisiones.setText(
            '1A2UP-JKrQvJV0SCMSD0IDa3ts-uOUJVR'
        )

        self.editJsonPath.setText(
            fr'C:\Users\{os.getlogin()}\Desktop\Emes despachos\emes-empaques-718c8b73c894.json'
            # r'C:\Digital\emes-empaques-718c8b73c894.json'
        )

        self.editJsonPathGS.setText(
            fr'C:\Users\{os.getlogin()}\Desktop\Emes despachos\emes-empaques-gspread.json'
            # r'C:\Digital\emes-empaques-gspread.json'
        )

        # not modifications enabled
        self.__enable_read_only()

        # hide text
        self.__hide_text()

    def __enable_read_only(self) -> None:
        """
        Enable read only mode 
        """
        self.editURLDB.setReadOnly(True)
        self.editIDPlanillas.setReadOnly(True)
        self.editIDRemisiones.setReadOnly(True)
        self.editJsonPath.setReadOnly(True)
        self.editJsonPathGS.setReadOnly(True)

    def __disable_read_only(self) -> None:
        """
        Enable read only mode 
        """
        self.editURLDB.setReadOnly(False)
        self.editIDPlanillas.setReadOnly(False)
        self.editIDRemisiones.setReadOnly(False)
        self.editJsonPath.setReadOnly(False)
        self.editJsonPathGS.setReadOnly(False)

    def __hide_text(self) -> None:
        """
        Hide text of QLineEdit objects
        """
        self.editURLDB.setEchoMode(QLineEdit.EchoMode.NoEcho)
        self.editIDPlanillas.setEchoMode(QLineEdit.EchoMode.NoEcho)
        self.editIDRemisiones.setEchoMode(QLineEdit.EchoMode.NoEcho)
        self.editJsonPath.setEchoMode(QLineEdit.EchoMode.NoEcho)
        self.editJsonPathGS.setEchoMode(QLineEdit.EchoMode.NoEcho)

    def __show_text(self) -> None:
        """
        Show text of QLineEdit objects
        """
        self.editURLDB.setEchoMode(QLineEdit.EchoMode.Normal)
        self.editIDPlanillas.setEchoMode(QLineEdit.EchoMode.Normal)
        self.editIDRemisiones.setEchoMode(QLineEdit.EchoMode.Normal)
        self.editJsonPath.setEchoMode(QLineEdit.EchoMode.Normal)
        self.editJsonPathGS.setEchoMode(QLineEdit.EchoMode.Normal)

    def progress_fn(self, n):
        print("%d%% done" % n)

    def print_output(self, s):
        print(s)

    def thread_complete(self):
        print("Thread complete!")

    def click_connect(self):
        """
        Button click method to show suppliers in subwindow
        """
        self.__thread_connect()

        self.connectButton.setStyleSheet(
            '''
            QPushButton {
                background-color: rgb(141, 206, 233);
                color: rgb(255, 255, 255);
                border-color: rgb(0, 150, 214);
                border-style: solid;
                border-width: 0px 1px 2px 0px;
                border-radius: 10px;
            }
            '''
        )
        self.connectButton.setText('Conectado')

    def click_edit(self):
        """
        Button click method to create EmesReport object
        """
        self.__thread_edit()

    def click_disconnect(self):
        """
        Button click method to create summary
        """
        self.__thread_disconnect()

    def click_generate_treasury_report(self):
        """
        Button click to generate treasury report
        """
        self.__thread_generate_treasury_report()

    def click_generate_dispatches_report(self):
        """
        Button click to generate dispatches report
        """
        self.__thread_generate_dispatches_report()

    def click_generate_main_report(self):
        """
        Button click to generate main report
        """
        self.__thread_generate_main_report()

    def __thread_connect(self):
        """
        Add thread to create report Emes
        """
        # worker
        worker = Worker(self.__connect)
        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # start thread
        self.threadpool.start(worker)

    def __thread_edit(self):
        """
        Add thread to create report Emes
        """
        # worker
        worker = Worker(self.__edit)
        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # start thread
        self.threadpool.start(worker)

    def __thread_disconnect(self):
        """
        Add thread to create report Emes
        """
        # worker
        worker = Worker(self.__disconnect)
        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # start thread
        self.threadpool.start(worker)

    def __thread_generate_treasury_report(self):
        """
        Add thread to create report Emes
        """
        # worker
        worker = Worker(self.__generate_treasury_report)
        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # start thread
        self.threadpool.start(worker)

    def __thread_generate_dispatches_report(self):
        """
        Add thread to create report Emes
        """
        # worker
        worker = Worker(self.__generate_dispatches_report)
        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # start thread
        self.threadpool.start(worker)

    def __thread_generate_main_report(self):
        """
        Add thread to create report Emes
        """
        # worker
        worker = Worker(self.__generate_main_report)
        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # start thread
        self.threadpool.start(worker)

    def __connect(self) -> None:
        """
        Button click method to show suppliers in subwindow
        """
        try:
            if self.emes is None:
                db_url = self.editURLDB.text()
                folder_id_planillas = self.editIDPlanillas.text()
                folder_id_remisiones = self.editIDRemisiones.text()
                json_path_db = self.editJsonPath.text()
                json_path_gs = self.editJsonPathGS.text()

                self.emes = EmesDispatch(
                    db_url,
                    folder_id_planillas,
                    folder_id_remisiones,
                    json_path_db,
                    json_path_gs
                )
                self.emes.connect()

        except:
            logging.error(
                'Error creando el object EmesDispatch',
                exc_info=True
            )

    def __edit(self) -> None:
        """
        Button click method to show suppliers in subwindow
        """
        try:
            if self.editURLDB.echoMode() == QLineEdit.EchoMode.NoEcho:
                self.__show_text()
                self.__disable_read_only()
            else:
                self.__hide_text()
                self.__enable_read_only()
        except:
            logging.error(
                'Error editando variables del bot',
                exc_info=True
            )

    def __disconnect(self) -> None:
        """
        Disconnect bot
        """
        try:
            if self.emes is not None:
                self.emes.disconnect()
        except:
            logging.error(
                'Error deteniendo el bot',
                exc_info=True
            )

    def __generate_treasury_report(self) -> None:
        """
        Generate treasury report
        """
        try:
            file = \
                self.dateEditTreasury.date().toPyDate().strftime('%Y-%m-%d')

            if self.emes is not None:
                self.emes.generate_report(
                    type="treasury",
                    filename=file
                )
        except:
            logging.error(
                "Error al crear el reporte de Tesorería",
                exc_info=True
            )

    def __generate_dispatches_report(self) -> None:
        """
        Generate dispatches report
        """
        try:
            file = \
                self.dateEditDispatches.date().toPyDate().strftime('%Y-%m-%d')

            if self.emes is not None:
                self.emes.generate_report(
                    type="dispatches",
                    filename=file
                )
        except:
            logging.error(
                "Error al crear el reporte de Despachos",
                exc_info=True
            )

    def __generate_main_report(self) -> None:
        """
        Generate main report (Consolidado)
        """
        try:
            file = \
                self.dateEditReport.date().toPyDate().strftime('%Y-%m-%d')

            if self.emes is not None:
                self.emes.generate_report(
                    type="main",
                    filename=file
                )
        except:
            logging.error(
                "Error al crear el reporte principal",
                exc_info=True
            )

    def closeEvent(self, event):
        """
        Close event

        Args:
            event: PyQt6 event
        """
        close = QMessageBox.question(
            self,
            "Salida del programa",
            "¿Está seguro que desea salir del programa?",
            QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.No
        )

        if close == QMessageBox.StandardButton.Ok:
            self.click_disconnect()
            event.accept()
            sys.exit()
        else:
            event.ignore()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.setFixedWidth(920)
    window.setFixedHeight(720)
    window.show()
    sys.exit(app.exec())
