import socket
from typing import NamedTuple, Optional
from PyQt6.QtCore import QSettings, Qt
from PyQt6.QtWidgets import QApplication, QInputDialog, QLineEdit, QMessageBox
from openai import APITimeoutError, APIConnectionError, OpenAI

SETTINGS_ORG = "AI_Tools"
SETTINGS_APP = "PPT_Translator"
SETTINGS_KEY = "openai_api_key"


class ApiKeyValidationResult(NamedTuple):
    is_valid: bool
    error_msg: str
    is_network_error: bool = False


def can_reach_openai(timeout_s: float = 3.0) -> bool:
    """Быстрая проверка доступа к api.openai.com (TCP 443)."""

    try:
        with socket.create_connection(("api.openai.com", 443), timeout=timeout_s):
            return True
    except OSError:
        return False


def show_no_internet_message(parent, details: str = "") -> bool:
    """Показывает простое окно "Нет интернета" (Retry/Cancel).

    Returns:
        True: Retry
        False: Cancel
    """

    text = (
        "Нет подключения к интернету или сервис OpenAI недоступен.\n\n"
        "Проверьте интернет и нажмите Retry, либо Cancel для выхода."
    )
    if details:
        text = f"{text}\n\nДетали:\n{details}"

    btn = QMessageBox.warning(
        parent,
        "Нет интернета",
        text,
        QMessageBox.StandardButton.Retry | QMessageBox.StandardButton.Cancel,
        QMessageBox.StandardButton.Retry,
    )
    return btn == QMessageBox.StandardButton.Retry

def validate_api_key(api_key: str) -> ApiKeyValidationResult:
    """Проверка ключа без списания токенов за генерацию."""
    try:
        test_client = OpenAI(api_key=api_key)
        test_client.models.list()
        return ApiKeyValidationResult(True, "")
    except (APIConnectionError, APITimeoutError) as e:
        msg = str(e) or "Не удалось подключиться к OpenAI (ошибка сети)."
        return ApiKeyValidationResult(False, msg, True)
    except Exception as e:
        status_code = getattr(e, "status_code", None)
        message = None

        body = getattr(e, "body", None)
        if isinstance(body, dict):
            err = body.get("error")
            if isinstance(err, dict):
                message = err.get("message")
            else:
                message = body.get("message")

        if not message:
            message = getattr(e, "message", None)

        if not message:
            message = str(e) or None

        msg_label = message or "Unknown error"

        if status_code:
            return ApiKeyValidationResult(False, f"Код: {status_code}\n{msg_label}")

        return ApiKeyValidationResult(False, msg_label)

def get_openai_api_key() -> Optional[str]:
    """Возвращает валидный OpenAI API Key.

    Returns:
        str: валидный ключ
        None: если пользователь отменил ввод (Cancel)
    """
    settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
    api_key = (settings.value(SETTINGS_KEY, "") or "").strip()

    parent = QApplication.activeWindow()

    while True:
        if not can_reach_openai():
            if not show_no_internet_message(parent):
                return None
            continue

        if api_key:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            QApplication.processEvents()
            try:
                validation = validate_api_key(api_key)
            finally:
                QApplication.restoreOverrideCursor()

            if validation.is_valid:
                return api_key

            if validation.is_network_error:
                if not show_no_internet_message(parent, validation.error_msg):
                    return None
                continue

            QMessageBox.warning(
                parent,
                "Ошибка API ключа",
                f"Сохраненный ключ невалиден, введите новый ключ\n\nОшибка:\n{validation.error_msg}",
            )
            api_key = ""
            continue

        key, ok = QInputDialog.getText(
            parent,
            "Настройка API",
            "Введите ваш OpenAI API Key (ключ будет проверен и сохранен в реестре):",
            QLineEdit.EchoMode.Password,
            "",
        )
        key = (key or "").strip()

        if ok and not key:
            QMessageBox.warning(
                parent,
                "Пустой ключ",
                "Поле API ключа пустое. Пожалуйста, введите ключ или нажмите Cancel для выхода.",
            )
            api_key = ""
            continue

        if ok and key:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            QApplication.processEvents()
            try:
                validation = validate_api_key(key)
            finally:
                QApplication.restoreOverrideCursor()

            if validation.is_valid:
                settings.setValue(SETTINGS_KEY, key)
                QMessageBox.information(parent, "Успех", "API ключ успешно проверен и сохранен!")
                return key

            if validation.is_network_error:
                if not show_no_internet_message(parent, validation.error_msg):
                    return None
                api_key = ""
                continue

            QMessageBox.critical(
                parent,
                "Ошибка",
                f"Ключ не прошел проверку\n\nОшибка:\n{validation.error_msg}",
            )
            api_key = ""
            continue

        return None
