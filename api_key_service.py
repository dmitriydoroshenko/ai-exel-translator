from concurrent.futures import CancelledError
from typing import Optional, Tuple
from PyQt6.QtCore import QSettings, Qt
from PyQt6.QtWidgets import QApplication, QInputDialog, QLineEdit, QMessageBox
from openai import OpenAI

SETTINGS_ORG = "AI_Tools"
SETTINGS_APP = "PPT_Translator"
SETTINGS_KEY = "openai_api_key"

def validate_api_key(api_key: str) -> Tuple[bool, str]:
    """Проверка ключа без списания токенов за генерацию."""
    try:
        test_client = OpenAI(api_key=api_key)
        test_client.models.list()
        return True, ""
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
            return False, f"Код: {status_code}\n{msg_label}"

        return False, msg_label

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
        if api_key:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            QApplication.processEvents()
            try:
                is_valid, _error_msg = validate_api_key(api_key)
            finally:
                QApplication.restoreOverrideCursor()

            if is_valid:
                return api_key

            QMessageBox.warning(
                parent,
                "Ошибка API ключа",
                f"Сохраненный ключ невалиден, введите новый ключ\n\nОшибка:\n{_error_msg}",
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
                is_valid, _error_msg = validate_api_key(key)
            finally:
                QApplication.restoreOverrideCursor()

            if is_valid:
                settings.setValue(SETTINGS_KEY, key)
                QMessageBox.information(parent, "Успех", "API ключ успешно проверен и сохранен!")
                return key

            QMessageBox.critical(
                parent,
                "Ошибка",
                f"Ключ не прошел проверку\n\nОшибка:\n{_error_msg}",
            )
            api_key = ""
            continue

        return None

def get_openai_client() -> OpenAI:
    """Создаёт и возвращает OpenAI client.

    Raises:
        RuntimeError: если не удалось получить ключ из-за ошибки.
        CancelledError: если пользователь отменил ввод ключа.
    """
    try:
        api_key = get_openai_api_key()
    except Exception as e:
        raise RuntimeError(f"Не удалось получить OpenAI API ключ: {e}") from e

    if not api_key:
        raise CancelledError("❌ Перевод отменён: API ключ не настроен.")

    return OpenAI(api_key=api_key)
