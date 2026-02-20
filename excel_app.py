import win32com.client as win32

class ExcelApp:
    """Контекст-менеджер для Excel.Application.

    Задача: централизовать создание/настройку Excel, а также гарантировать Quit()
    при любом выходе из блока `with` (в т.ч. при исключениях).
    """

    def __init__(self, visible: bool = False, display_alerts: bool = False):
        self.visible = visible
        self.display_alerts = display_alerts
        self.excel = None

    def __enter__(self):
        self.excel = win32.DispatchEx("Excel.Application")
        self.excel.Visible = self.visible
        self.excel.DisplayAlerts = self.display_alerts
        return self

    def open_workbook(self, path: str):
        if self.excel is None:
            raise RuntimeError("ExcelApp не запущен")
        return WorkbookSession(self.excel, path)

    def __exit__(self, _exc_type, _exc, _tb):
        if self.excel is not None:
            try:
                self.excel.Quit()
            except Exception:
                pass
            self.excel = None


class WorkbookSession:
    """Контекст-менеджер для Workbooks.Open(...).

    Всегда закрывает книгу при выходе (без сохранения), потому что сохранение
    в вашем пайплайне делается через SaveAs(output_file).
    """

    def __init__(self, excel, path: str):
        self.excel = excel
        self.path = path
        self.wb = None

    def __enter__(self):
        self.wb = self.excel.Workbooks.Open(self.path)
        return self.wb

    def __exit__(self, _exc_type, _exc, _tb):
        if self.wb is not None:
            try:
                self.wb.Close(SaveChanges=False)
            except Exception:
                pass
            self.wb = None