import win32com.client as win32
import os
import json
import sys
import io
import time
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

# Цены за 1 млн токенов
PRICE_IN = 1.75
PRICE_OUT = 14.00

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    print("Ошибка: API ключ не найден в переменных окружения (.env)")
    sys.exit()

client = OpenAI(api_key=api_key)

total_in = 0
total_out = 0

def translate_batch(batch_dict):
    """Отправка пачки текста на перевод в OpenAI."""
    global total_in, total_out
    if not batch_dict: 
        return {}
    try:
        SYSTEM_ROLE = (
            "## Role\n"
            "You are an expert Game Localization (L10N) Specialist and professional mobile game localizer. "
            "Your goal is to translate English mobile gaming market reports and game text into Simplified Chinese, "
            "ensuring the output is natural and uses industry-standard jargon used by developers and publishers.\n\n"

            "## Terminology & Style Guidelines\n"
            "- Do Not Translate Game Titles: Keep all game names/titles in their original English form.\n"
            "- Avoid Literalism: Do not translate word-for-word. Focus on industry 'jargon.'\n"
            "- Spending/Monetization:\n"
            "  * 'Non-paying players' -> 非付费玩家 / 零氪玩家\n"
            "  * 'Spending real money' -> 付费 / 氪金\n"
            "- Events & Scheduling:\n"
            "  * 'Global schedule' -> 全服统一日程 / 固定档期\n"
            "  * 'Progress in events' -> 推进活动进度\n"
            "- Tone: Professional, concise, and analytical. Use 'Game-speak.'\n\n"

            "## STRICT RULES\n"
            "1. DO NOT translate game titles, bundle names, or offer names. Keep them in English."
            "2. Translate all other values (descriptions, analysis, labels) into Simplified Chinese using the guidelines above."
            "3. Keep JSON keys unchanged."
            "4. Return ONLY a valid JSON object without any markdown formatting or extra text outside the JSON."
        )

        response = client.chat.completions.create(
            model="gpt-5.2",
            messages=[
                {"role": "system", "content": SYSTEM_ROLE},
                {"role": "user", "content": json.dumps(batch_dict, ensure_ascii=False)}
            ],
            response_format={"type": "json_object"},
            timeout=30
        )

        if response.usage:
            total_in += response.usage.prompt_tokens
            total_out += response.usage.completion_tokens

        content = response.choices[0].message.content
        
        if content is None:
            print("\n[ОШИБКА]: API вернул пустой ответ (None).")
            return None
            
        result = json.loads(content)
        return result if result else {}

    except Exception as e:
        print(f"\n[КРИТИЧЕСКАЯ ОШИБКА API]: {e}")
        return None

def main():
    start_time = time.time()
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

    translations_cache = {} 

    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(input_file)
        
        print("Перевод названий листов...")
        sheet_batch = {f"sh_{i}": sheet.Name for i, sheet in enumerate(wb.Sheets)}
        translated_sheet_data = translate_batch(sheet_batch)
        
        if translated_sheet_data:
            for i, sheet in enumerate(wb.Sheets):
                sheet.Name = translated_sheet_data.get(f"sh_{i}", sheet.Name)

        total_sheets = wb.Sheets.Count

        for index, sheet in enumerate(wb.Sheets, 1):
            sys.stdout.write(f"Лист [{index}/{total_sheets}]: {sheet.Name} —> Сбор данных...")
            sys.stdout.flush()
            used_range = sheet.UsedRange
            cell_mapping = []  
            unique_texts_to_translate = set() 
            
            for r in range(1, used_range.Rows.Count + 1):
                for c in range(1, used_range.Columns.Count + 1):
                    cell = used_range.Cells(r, c)
                    val = cell.Value
                    if isinstance(val, str) and len(val.strip()) > 1:
                        if not str(cell.Formula).startswith('='):
                            text = val.strip()
                            cell_mapping.append((cell.GetAddress(), text))
                            if text not in translations_cache:
                                unique_texts_to_translate.add(text)

            for chart_obj in sheet.ChartObjects():
                if chart_obj.Chart.HasTitle:
                    title = chart_obj.Chart.ChartTitle.Text
                    if title and len(title.strip()) > 1:
                        text = title.strip()
                        cell_mapping.append((f"CHART:{chart_obj.Name}", text))
                        if text not in translations_cache:
                            unique_texts_to_translate.add(text)

            if unique_texts_to_translate:
                unique_list = list(unique_texts_to_translate)
                sys.stdout.write(f" -> Перевод {len(unique_list)} новых строк...")
                sys.stdout.flush()

                for i in range(0, len(unique_list), 30):
                    batch = {f"id_{j}": text for j, text in enumerate(unique_list[i:i+30])}
                    res = translate_batch(batch)
                    if res is None:
                        wb.Close(False); excel.Quit(); sys.exit()

                    for batch_id, trans_text in res.items():
                        orig_text = batch[batch_id]
                        translations_cache[orig_text] = trans_text

            sys.stdout.write(f" -> Применяю перевод...")
            sys.stdout.flush()

            for identifier, original_text in cell_mapping:
                translated_text = translations_cache.get(original_text, original_text)
                
                if identifier.startswith("CHART:"):
                    c_name = identifier.replace("CHART:", "")
                    chart = sheet.ChartObjects(c_name).Chart
                    chart.ChartTitle.Text = translated_text
                    try:
                        chart.ChartTitle.Font.Name = "Microsoft YaHei"
                        for axis in chart.Axes():
                            try: axis.TickLabels.Font.Name = "Microsoft YaHei"
                            except: pass
                    except: pass
                else:
                    cell_range = sheet.Range(identifier)
                    cell_range.Value = translated_text
                    try: cell_range.Font.Name = "Microsoft YaHei"
                    except: pass

            sys.stdout.write(" ✅\n")
            sys.stdout.flush()

        wb.SaveAs(output_file)

        end_time = time.time()
        duration = end_time - start_time
        
        total_cost = (total_in / 1_000_000 * PRICE_IN) + (total_out / 1_000_000 * PRICE_OUT)

        print(f"\nГотово! Результат в: output/{os.path.basename(output_file)}")
        print(f"Токены: {total_out + total_in} | Стоимость: ${total_cost:.4f}")
        print(f"Общее время: {int(duration // 60)} мин. {int(duration % 60)} сек.\n")
        

    except Exception as e:
        print(f"\n[ОШИБКА]: {e}")
        if 'wb' in locals(): wb.Close(False)
    finally:
        if 'excel' in locals(): excel.Quit()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n[STOP] Программа остановлена. Очистка ресурсов...")
        try:
            excel_app = win32.GetActiveObject("Excel.Application")
            excel_app.Quit()
            print("Процесс Excel успешно завершен.")
        except Exception:
            pass
        sys.exit(0)