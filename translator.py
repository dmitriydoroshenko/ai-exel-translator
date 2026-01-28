import win32com.client as win32
import os
import json
import sys
import io
from dotenv import load_dotenv
from openai import OpenAI

# Загружаем переменные из файла .env в систему
load_dotenv()

# Настройка вывода кириллицы для терминала
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    print("Ошибка: API ключ не найден в переменных окружения (.env)")
    sys.exit()

# Инициализация клиента
client = OpenAI(api_key=api_key)

def translate_batch(batch_dict):
    """Отправляет словарь на перевод и возвращает результат или None при ошибке"""
    if not batch_dict: 
        return {}
    
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a professional translator. Return JSON with same keys but Russian translations."},
                {"role": "user", "content": json.dumps(batch_dict, ensure_ascii=False)}
            ],
            response_format={"type": "json_object"},
            timeout=30 # Таймаут, чтобы скрипт не висел бесконечно
        )
        
        result = json.loads(response.choices[0].message.content)
        
        # Проверка: если ИИ вернул пустой объект вместо перевода
        if not result:
            return None
        return result

    except Exception as e:
        print(f"\n[КРИТИЧЕСКАЯ ОШИБКА API]: {e}")
        return None

def main():
    input_file = os.path.abspath("test.xlsx")
    output_file = os.path.abspath("translated_final_secure.xlsx")

    # Проверка существования файла
    if not os.path.exists(input_file):
        print(f"Файл {input_file} не найден.")
        return

    # Запуск Excel
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True # Можно видеть процесс
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(input_file)
    except Exception as e:
        print(f"Не удалось запустить Excel или открыть файл: {e}")
        return

    try:
        for sheet in wb.Sheets:
            print(f"\nОбработка листа: {sheet.Name}")
            used_range = sheet.UsedRange
            
            batch = {}
            cells_in_batch = []
            
            # Читаем ячейки
            for r in range(1, used_range.Rows.Count + 1):
                for c in range(1, used_range.Columns.Count + 1):
                    cell = used_range.Cells(r, c)
                    val = cell.Value
                    
                    if isinstance(val, str) and len(val.strip()) > 1:
                        if not str(cell.Formula).startswith('='):
                            batch[cell.GetAddress()] = val
                            cells_in_batch.append(cell)
                            
                            if len(batch) >= 30:
                                res = translate_batch(batch)
                                
                                # ЕСЛИ ОШИБКА — ПРЕРЫВАЕМ ВСЁ
                                if res is None:
                                    print("\n[СТОП] Программа завершена без сохранения из-за ошибки перевода.")
                                    wb.Close(False) # Закрыть без сохранения изменений
                                    excel.Quit()
                                    sys.exit() # Полный выход из Python
                                
                                # Если всё ок — записываем
                                for c_obj in cells_in_batch:
                                    addr = c_obj.GetAddress()
                                    if addr in res:
                                        c_obj.Value = res[addr]
                                
                                print(f"Пачка переведена успешно.")
                                batch, cells_in_batch = {}, []

            # Перевод остатка
            if batch:
                res = translate_batch(batch)
                if res is None:
                    print("\n[СТОП] Ошибка на финальном этапе. Файл не сохранен.")
                    wb.Close(False)
                    excel.Quit()
                    sys.exit()
                
                for c_obj in cells_in_batch:
                    addr = c_obj.GetAddress()
                    if addr in res:
                        c_obj.Value = res[addr]

        # СОХРАНЕНИЕ ПРОИСХОДИТ ТОЛЬКО ЗДЕСЬ
        wb.SaveAs(output_file)
        print(f"\n[УСПЕХ] Перевод завершен без ошибок. Файл сохранен: {output_file}")

    except Exception as e:
        print(f"\n[ОШИБКА В ПРОЦЕССЕ]: {e}")
        wb.Close(False)
    finally:
        excel.Quit()

if __name__ == "__main__":
    main()