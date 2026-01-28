import win32com.client as win32
import os
import json
import sys
import io
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    print("Ошибка: API ключ не найден в переменных окружения (.env)")
    sys.exit()

client = OpenAI(api_key=api_key)

def translate_batch(batch_dict):
    """Отправка пачки текста на перевод в OpenAI."""
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
            timeout=30
        )
        result = json.loads(response.choices[0].message.content)
        return result if result else None
    except Exception as e:
        print(f"\n[КРИТИЧЕСКАЯ ОШИБКА API]: {e}")
        return None

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = os.path.join(script_dir, "input")
    output_folder = os.path.join(script_dir, "output")

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    try:
        files = [f for f in os.listdir(input_folder) if f.endswith(".xlsx")]
        if not files:
            print(f"Файлы не найдены в папке: {input_folder}")
            return
        filename = files[0]
    except Exception as e:
        print(f"Ошибка при доступе к папке input: {e}")
        return

    input_file = os.path.join(input_folder, filename)
    
    name_part, extension = os.path.splitext(filename)
    output_file = os.path.join(output_folder, f"{name_part}_cn{extension}")

    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(input_file)
    except Exception as e:
        print(f"Не удалось запустить Excel: {e}")
        return

    try:
        total_sheets = wb.Sheets.Count

        for index, sheet in enumerate(wb.Sheets, 1):
            used_range = sheet.UsedRange
            
            cells_to_translate = []
            for r in range(1, used_range.Rows.Count + 1):
                for c in range(1, used_range.Columns.Count + 1):
                    cell = used_range.Cells(r, c)
                    val = cell.Value
                    if isinstance(val, str) and len(val.strip()) > 1:
                        if not str(cell.Formula).startswith('='):
                            cells_to_translate.append(cell)

            total_cells = len(cells_to_translate)
            
            if total_cells == 0:
                print(f"Лист [{index}/{total_sheets}]: {sheet.Name} Прогресс: 100% ✅ (Нет текста)")
                continue

            batch = {}
            processed_count = 0
            
            sys.stdout.write(f"Лист [{index}/{total_sheets}]: {sheet.Name} Прогресс: 0%")
            sys.stdout.flush()

            for i, cell in enumerate(cells_to_translate):
                batch[cell.GetAddress()] = cell.Value
                
                # Переводим, если набрали 30 ячеек или это последняя ячейка на листе
                if len(batch) >= 30 or i == total_cells - 1:
                    res = translate_batch(batch)
                    
                    if res is None:
                        print(f"\n[СТОП] Ошибка API на листе {sheet.Name}.")
                        wb.Close(False)
                        excel.Quit()
                        sys.exit()

                    # Запись перевода обратно в Excel
                    for addr, translated_text in res.items():
                        sheet.Range(addr).Value = translated_text
                    
                    processed_count += len(batch)
                    percent = min(100, int((processed_count / total_cells) * 100))
                    
                    # Обновление прогресса в той же строке
                    sys.stdout.write(f"\rЛист [{index}/{total_sheets}]: {sheet.Name} Прогресс: {percent}%")
                    sys.stdout.flush()
                    
                    batch = {}

            # Финальная галочка и переход на новую строку
            sys.stdout.write(" ✅\n")
            sys.stdout.flush()

        # Сохранение результата
        wb.SaveAs(output_file)
        print(f"\nГотово! Результат в: output/{os.path.basename(output_file)}")

    except Exception as e:
        print(f"\n[ОШИБКА]: {e}")
        if 'wb' in locals(): wb.Close(False)
    finally:
        if 'excel' in locals(): excel.Quit()

if __name__ == "__main__":
    main()
