import sys
import traceback
from pathlib import Path
import threading
from concurrent.futures import CancelledError
import pythoncom
from wakepy import keep
from PyQt6 import QtCore, QtGui, QtWidgets
import main as translator_main
from excel_app import cleanup_excel
from api_key_service import get_openai_api_key


def _load_app_icon() -> QtGui.QIcon:
    """Загружает иконку приложения"""

    candidates: list[Path] = []

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass))

    candidates.append(Path(__file__).resolve().parent)

    for base_dir in list(candidates):
        candidates.append(base_dir / "_internal")

    for base_dir in candidates:
        icon_path = base_dir / "app_icon.ico"
        if icon_path.exists():
            return QtGui.QIcon(str(icon_path))
    return QtGui.QIcon()

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
    finished_cancelled = QtCore.pyqtSignal()

    def __init__(self, input_file: str, api_key: str, cancel_event: threading.Event, parent=None):
        super().__init__(parent)
        self.input_file = input_file
        self.api_key = api_key
        self.cancel_event = cancel_event

    def request_cancel(self) -> None:
        self.cancel_event.set()

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
                translator_main.main(self.input_file, self.api_key, cancel_event=self.cancel_event)

            self.finished_ok.emit()

        except SystemExit as e:
            code = getattr(e, "code", 1)
            if code in (0, None):
                self.finished_ok.emit()
            else:
                self.finished_fail.emit(f"SystemExit: {code}")

        except CancelledError:
            self.finished_cancelled.emit()

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

        self.setWindowTitle("AI Excel Translator")

        self.setWindowIcon(_load_app_icon())
        self.setMinimumSize(550, 500)
        self.resize(550, 500)

        self.worker = None
        self.input_file = None
        self.cancel_event: threading.Event | None = None

        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)
        layout = QtWidgets.QVBoxLayout(central)
        layout.setSpacing(15)

        self.info_label = QtWidgets.QLabel("Выберите таблицу для перевода (.xlsx)")
        self.info_label.setObjectName("InfoLabel")
        self.info_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.info_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(self.info_label)

        self.choose_btn = QtWidgets.QPushButton("📂 Выбрать файл")
        self.choose_btn.setObjectName("ChooseBtn")
        self.choose_btn.setMinimumHeight(45)
        self.choose_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.choose_btn.clicked.connect(self.on_choose_file)
        layout.addWidget(self.choose_btn)

        self.log_view = QtWidgets.QPlainTextEdit()
        self.log_view.setPlaceholderText("Лог процесса перевода появится здесь...")
        self.log_view.setReadOnly(True)
        self.log_view.setLineWrapMode(QtWidgets.QPlainTextEdit.LineWrapMode.NoWrap)
        layout.addWidget(self.log_view)

        self.action_stack = QtWidgets.QStackedWidget()
        layout.addWidget(self.action_stack)

        self.start_btn = QtWidgets.QPushButton("🚀 Начать перевод")
        self.start_btn.setObjectName("StartBtn")
        self.start_btn.setMinimumHeight(45)
        self.start_btn.setCursor(QtCore.Qt.CursorShape.ArrowCursor)
        self.start_btn.setStyleSheet(
            """
            QPushButton {
                background-color: #2ecc71;
                color: white;
                padding: 12px;
                font-size: 14px;
                font-weight: bold;
                border-radius: 6px;
            }
            QPushButton:disabled { background-color: #95a5a6; }
            QPushButton:hover { background-color: #27ae60; }
            """
        )
        self.start_btn.clicked.connect(self.on_start)
        self.start_btn.setEnabled(False)

        self.cancel_btn = QtWidgets.QPushButton("⛔ Отмена")
        self.cancel_btn.setObjectName("CancelBtn")
        self.cancel_btn.setMinimumHeight(45)
        self.cancel_btn.setEnabled(True)
        self.cancel_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.cancel_btn.clicked.connect(self.on_cancel)

        self.action_stack.addWidget(self.start_btn)
        self.action_stack.addWidget(self.cancel_btn)
        self.action_stack.setCurrentWidget(self.start_btn)

        self.action_stack.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed)
        self.action_stack.setMinimumHeight(45)

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
            "📂 Выберите .xlsx файл",
            "",
            "Excel Files (*.xlsx)",
        )

        if input_file:
            self.input_file = input_file
            self.log_view.clear()
            self.append_log(f"✅ Выбран файл: {input_file}\n\n")
            self.start_btn.setEnabled(True)
            self.start_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)

    def on_start(self) -> None:
        if self.worker is not None and self.worker.isRunning():
            return

        if not self.input_file:
            self.append_log("⚠️ Сначала выберите .xlsx файл.\n")
            return

        try:
            api_key = get_openai_api_key()
        except Exception as e:
            self.append_log(f"❌ Не удалось получить OpenAI API ключ: {e}\n")
            return

        if not api_key:
            self.append_log("❌ Перевод отменён: API ключ не настроен.\n")
            return

        self.choose_btn.setEnabled(False)

        self.cancel_btn.setEnabled(True)
        self.cancel_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)

        self.action_stack.setCurrentWidget(self.cancel_btn)

        self.cancel_event = threading.Event()

        self.worker = TranslateWorker(self.input_file, api_key, cancel_event=self.cancel_event, parent=self)
        self.worker.log.connect(self.append_log)
        self.worker.finished_ok.connect(self.on_finished_ok)
        self.worker.finished_fail.connect(self.on_finished_fail)
        self.worker.finished_cancelled.connect(self.on_finished_cancelled)
        self.worker.start()

    def on_cancel(self) -> None:
        if self.worker is None or not self.worker.isRunning():
            return

        self.append_log("\n⛔ Запрошена отмена...\n")
        try:
            self.worker.request_cancel()
        except Exception:
            if self.cancel_event is not None:
                self.cancel_event.set()

        self.cancel_btn.setEnabled(False)
        self.cancel_btn.setCursor(QtCore.Qt.CursorShape.ArrowCursor)

    def on_finished_ok(self) -> None:
        self.action_stack.setCurrentWidget(self.start_btn)
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(True)
        self.cancel_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.choose_btn.setEnabled(True)

        self.choose_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.start_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)

        if not self.input_file:
            self.start_btn.setEnabled(False)
            self.start_btn.setCursor(QtCore.Qt.CursorShape.ArrowCursor)

    def on_finished_fail(self, detail: str) -> None:
        self.append_log("\n\n❌ Ошибка:\n" + (detail or "") + "\n")
        self.action_stack.setCurrentWidget(self.start_btn)
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(True)
        self.cancel_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.choose_btn.setEnabled(True)

        self.choose_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.start_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)

        if not self.input_file:
            self.start_btn.setEnabled(False)
            self.start_btn.setCursor(QtCore.Qt.CursorShape.ArrowCursor)

    def on_finished_cancelled(self) -> None:
        self.append_log("⛔ Перевод отменён. Результат не сохранён.\n\n")

        self.action_stack.setCurrentWidget(self.start_btn)
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(True)
        self.cancel_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.choose_btn.setEnabled(True)

        self.choose_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.start_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)

        if not self.input_file:
            self.start_btn.setEnabled(False)
            self.start_btn.setCursor(QtCore.Qt.CursorShape.ArrowCursor)

    def closeEvent(self, event):
        try:
            cleanup_excel()
        except Exception:
            pass
        super().closeEvent(event)

def main() -> None:
    app = QtWidgets.QApplication(sys.argv)

    app.setWindowIcon(_load_app_icon())
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
