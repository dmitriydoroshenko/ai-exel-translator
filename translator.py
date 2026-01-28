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
        excel.Visible = True
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(input_file)
    except Exception as e:
        print(f"Не удалось запустить Excel: {e}")
        return

    try:
        for sheet in wb.Sheets:
            print(f"\nОбработка листа: {sheet.Name}")
            used_range = sheet.UsedRange
            
            batch = {}
            cells_in_batch = []
            
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
                                if res is None:
                                    print("\n[СТОП] Ошибка API.")
                                    wb.Close(False)
                                    excel.Quit()
                                    sys.exit()
                                
                                for c_obj in cells_in_batch:
                                    addr = c_obj.GetAddress()
                                    if addr in res:
                                        c_obj.Value = res[addr]
                                
                                print(f"Пачка переведена...")
                                batch, cells_in_batch = {}, []

            if batch:
                res = translate_batch(batch)
                if res:
                    for c_obj in cells_in_batch:
                        addr = c_obj.GetAddress()
                        if addr in res:
                            c_obj.Value = res[addr]

        wb.SaveAs(output_file)
        print(f"\n[УСПЕХ] Готово! Файл в output: {os.path.basename(output_file)}")

    except Exception as e:
        print(f"\n[ОШИБКА]: {e}")
        wb.Close(False)
    finally:
        excel.Quit()

if __name__ == "__main__":
    main()
