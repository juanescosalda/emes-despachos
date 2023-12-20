import sys
import traceback
from PyQt6.QtGui import *
from PyQt6.QtCore import *


class WorkerSignals(QObject):
    """
    Defines the signals available from a running worker thread.

    finished: No data
    error: tuple (exctype, value, traceback.format_exc() )
    result: object data returned from processing, anything

    """
    finished = pyqtSignal()
    error = pyqtSignal(tuple)
    result = pyqtSignal(object)
    progress = pyqtSignal(int)


class Worker(QRunnable):
    """
    Worker thread

    Inherits from QRunnable to handler worker thread setup, signals and wrap-up.

    type callback: function
    """

    def __init__(self, fn):
        super().__init__()

        # Store constructor arguments (re-used for processing)
        self.fn = fn
        self.signals = WorkerSignals()

    @pyqtSlot()
    def run(self):
        """
        Initialise the runner function.
        """
        try:
            result = self.fn()
        except:
            traceback.print_exc()
            exctype, value = sys.exc_info()[:2]
            self.signals.error.emit((exctype, value, traceback.format_exc()))
        else:
            self.signals.result.emit(result)
        finally:
            self.signals.finished.emit()
