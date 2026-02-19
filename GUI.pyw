import sys
import traceback
import pythoncom
from wakepy import keep
from PyQt6 import QtCore, QtGui, QtWidgets
import main as translator_main
from excel_app import cleanup_excel

class QtStream(QtCore.QObject):
    """Файлоподобный объект для перенаправления stdout/stderr в GUI."""

    text = QtCore.pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.encoding = "utf-8"

    def write(self, s):
        if s is None:
            return 0
        try:
            text = str(s)
        except Exception:
            text = "[stream] <unprintable>"
        if text:
            self.text.emit(text)
        return len(text)

    def flush(self):
        return

    def isatty(self):
        return False

class TranslateWorker(QtCore.QThread):
    log = QtCore.pyqtSignal(str)
    finished_ok = QtCore.pyqtSignal()
    finished_fail = QtCore.pyqtSignal(str)

    def __init__(self, input_file: str, parent=None):
        super().__init__(parent)
        self.input_file = input_file

    def run(self):
        pythoncom.CoInitialize()

        old_out, old_err = sys.stdout, sys.stderr

        out_stream = QtStream()
        err_stream = QtStream()
        out_stream.text.connect(self.log)
        err_stream.text.connect(self.log)

        try:
            sys.stdout = out_stream
            sys.stderr = err_stream

            with keep.running():
                translator_main.main(self.input_file)

            self.finished_ok.emit()

        except SystemExit as e:
            code = getattr(e, "code", 1)
            if code in (0, None):
                self.finished_ok.emit()
            else:
                self.finished_fail.emit(f"SystemExit: {code}")

        except Exception as e:
            try:
                cleanup_excel()
            except Exception:
                pass
            detail = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            self.finished_fail.emit(detail)

        finally:
            sys.stdout, sys.stderr = old_out, old_err
            try:
                cleanup_excel()
            except Exception:
                pass
            pythoncom.CoUninitialize()

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI Excel Translator — Логи")
        self.resize(900, 600)

        self.worker = None
        self.input_file = None

        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)
        layout = QtWidgets.QVBoxLayout(central)

        btn_row = QtWidgets.QHBoxLayout()
        self.choose_btn = QtWidgets.QPushButton("Выбрать файл")
        self.choose_btn.clicked.connect(self.on_choose_file)
        btn_row.addWidget(self.choose_btn)

        self.start_btn = QtWidgets.QPushButton("Начать перевод")
        self.start_btn.clicked.connect(self.on_start)
        btn_row.addWidget(self.start_btn)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)

        self.log_view = QtWidgets.QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setLineWrapMode(QtWidgets.QPlainTextEdit.LineWrapMode.NoWrap)
        layout.addWidget(self.log_view)

    @QtCore.pyqtSlot(str)
    def append_log(self, text: str) -> None:
        if not text:
            return
        cursor = self.log_view.textCursor()
        cursor.movePosition(QtGui.QTextCursor.MoveOperation.End)
        cursor.insertText(text)
        self.log_view.setTextCursor(cursor)
        self.log_view.ensureCursorVisible()

    def on_choose_file(self) -> None:
        input_file, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Выберите .xlsx файл",
            "",
            "Excel Files (*.xlsx)",
        )

        if input_file:
            self.input_file = input_file
            self.log_view.clear()
            self.append_log(f"Выбран файл: {input_file}\n\n")

    def on_start(self) -> None:
        if self.worker is not None and self.worker.isRunning():
            return

        if not self.input_file:
            self.append_log("Сначала выберите .xlsx файл.\n")
            return

        self.start_btn.setEnabled(False)

        self.worker = TranslateWorker(self.input_file, self)
        self.worker.log.connect(self.append_log)
        self.worker.finished_ok.connect(self.on_finished_ok)
        self.worker.finished_fail.connect(self.on_finished_fail)
        self.worker.start()

    def on_finished_ok(self) -> None:
        self.start_btn.setEnabled(True)

    def on_finished_fail(self, detail: str) -> None:
        self.append_log("\n\n❌ Ошибка:\n" + (detail or "") + "\n")
        self.start_btn.setEnabled(True)

    def closeEvent(self, event):
        try:
            cleanup_excel()
        except Exception:
            pass
        super().closeEvent(event)

def main() -> None:
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
