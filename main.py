import win32com.client as win32
import os
import sys
import io
import time
from wakepy import keep
from translator import Translator

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    translator = Translator()
except Exception as e:
    print(str(e))
    sys.exit()

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

    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(input_file)
        
        print("Перевод названий листов...")
        sheet_batch = {f"sh_{i}": sheet.Name for i, sheet in enumerate(wb.Sheets)}
        translated_sheet_data = translator.translate_batch(sheet_batch)
        
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
                            unique_texts_to_translate.add(text)

            for chart_obj in sheet.ChartObjects():
                chart = chart_obj.Chart
                if chart.HasTitle:
                    text = chart.ChartTitle.Text.strip()
                    cell_mapping.append((f"CHART_TITLE:{chart_obj.Name}", text))
                    unique_texts_to_translate.add(text)
                
                for s_idx in range(1, chart.SeriesCollection().Count + 1):
                    series = chart.SeriesCollection(s_idx)
                    try:
                        text = series.Name.strip()
                        if text and not text.isdigit():
                            cell_mapping.append((f"CHART_SERIES:{chart_obj.Name}:{s_idx}", text))
                            unique_texts_to_translate.add(text)
                    except: pass

                for ax_type in [1, 2]:
                    try:
                        axis = chart.Axes(ax_type)
                        if axis.HasTitle:
                            text = axis.AxisTitle.Text.strip()
                            cell_mapping.append((f"CHART_AXIS:{chart_obj.Name}:{ax_type}", text))
                            unique_texts_to_translate.add(text)
                    except: pass

            if unique_texts_to_translate:
                unique_list = list(unique_texts_to_translate)
                sys.stdout.write(f" -> Перевод {len(unique_list)} новых строк...")
                sys.stdout.flush()

                translations_map = translator.translate_texts(unique_list)
            else:
                translations_map = {}

            sys.stdout.write(f" -> Применяю перевод...")
            sys.stdout.flush()

            for identifier, original_text in cell_mapping:
                translated_text = translations_map.get(original_text, original_text)
                
                if identifier.startswith("CHART_TITLE:"):
                    c_name = identifier.replace("CHART_TITLE:", "")
                    chart = sheet.ChartObjects(c_name).Chart
                    chart.ChartTitle.Text = translated_text
                    try: chart.ChartTitle.Font.Name = "Microsoft YaHei"
                    except: pass
                
                elif identifier.startswith("CHART_SERIES:"):
                    parts = identifier.split(":")
                    series = sheet.ChartObjects(parts[1]).Chart.SeriesCollection(int(parts[2]))
                    series.Name = translated_text
                
                elif identifier.startswith("CHART_AXIS:"):
                    parts = identifier.split(":")
                    axis = sheet.ChartObjects(parts[1]).Chart.Axes(int(parts[2]))
                    axis.AxisTitle.Text = translated_text
                    try: axis.AxisTitle.Font.Name = "Microsoft YaHei"
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

        total_cost = translator.total_cost_usd

        print(f"\nГотово! Результат в: output/{os.path.basename(output_file)}")
        print(f"Токены: {translator.usage.total_tokens} | Стоимость: ${total_cost:.4f}")
        print(f"Общее время: {int(duration // 60)} мин. {int(duration % 60)} сек.\n")
        

    except Exception as e:
        print(f"\n[ОШИБКА]: {e}")
        if 'wb' in locals(): wb.Close(False)
    finally:
        if 'excel' in locals(): excel.Quit()

if __name__ == "__main__":
    try:
        with keep.running():
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