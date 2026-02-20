from concurrent.futures import CancelledError
from typing import Tuple
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

def get_openai_client() -> OpenAI:
    """Создаёт и возвращает OpenAI client"""

    try:
        settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
        api_key = (settings.value(SETTINGS_KEY, "") or "").strip()
        parent = QApplication.activeWindow()

        while True:
            # 1) Если ключ есть в настройках — проверяем его и используем.
            if api_key:
                QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
                QApplication.processEvents()
                try:
                    is_valid, error_msg = validate_api_key(api_key)
                finally:
                    QApplication.restoreOverrideCursor()

                if is_valid:
                    return OpenAI(api_key=api_key)

                QMessageBox.warning(
                    parent,
                    "Ошибка API ключа",
                    "Сохраненный ключ невалиден, введите новый ключ"
                    f"\n\nОшибка:\n{error_msg}",
                )
                api_key = ""
                continue

            # 2) Запрашиваем ключ у пользователя.
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
                    is_valid, error_msg = validate_api_key(key)
                finally:
                    QApplication.restoreOverrideCursor()

                if is_valid:
                    settings.setValue(SETTINGS_KEY, key)
                    QMessageBox.information(parent, "Успех", "API ключ успешно проверен и сохранен!")
                    return OpenAI(api_key=key)

                QMessageBox.critical(
                    parent,
                    "Ошибка",
                    f"Ключ не прошел проверку\n\nОшибка:\n{error_msg}",
                )
                api_key = ""
                continue

            raise CancelledError("❌ Перевод отменён: API ключ не настроен.")

    except CancelledError:
        raise
    except Exception as e:
        raise RuntimeError(f"Не удалось получить OpenAI API ключ: {e}") from e
