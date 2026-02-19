import os
import sys
import time
from wakepy import keep
from translator import Translator
from excel_app import ExcelApp, cleanup_excel

def main(input_file=None):
    try:
        start_time = time.time()

        translator = Translator()

        if input_file:
            input_file = os.path.abspath(input_file)

            if not os.path.isfile(input_file):
                raise FileNotFoundError(f"Файл не найден: {input_file}")

            name_part, extension = os.path.splitext(os.path.basename(input_file))
            if extension.lower() != ".xlsx":
                raise ValueError("Поддерживаются только файлы .xlsx")
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            input_folder = os.path.join(script_dir, "input")

            try:
                files = [f for f in os.listdir(input_folder) if f.endswith(".xlsx")]
            except Exception as e:
                raise RuntimeError(f"Ошибка при доступе к папке input: {e}") from e

            if not files:
                raise FileNotFoundError(f"Файлы .xlsx не найдены в папке: {input_folder}")

            filename = files[0]

            input_file = os.path.join(input_folder, filename)
            name_part, extension = os.path.splitext(filename)

        output_dir = os.path.dirname(input_file)
        base_output_name = f"{name_part}_cn"
        output_file = os.path.join(output_dir, f"{base_output_name}{extension}")
        index = 1
        while os.path.exists(output_file):
            output_file = os.path.join(output_dir, f"{base_output_name}_{index}{extension}")
            index += 1

        with ExcelApp() as exel_app:
            with exel_app.open_workbook(input_file) as workbook:
        
                print("Перевод названий листов...")
                sheet_batch = {f"sh_{i}": sheet.Name for i, sheet in enumerate(workbook.Sheets)}
                translated_sheet_data = translator.translate_batch(sheet_batch)
                
                if translated_sheet_data:
                    for i, sheet in enumerate(workbook.Sheets):
                        sheet.Name = translated_sheet_data.get(f"sh_{i}", sheet.Name)

                total_sheets = workbook.Sheets.Count

                for index, sheet in enumerate(workbook.Sheets, 1):
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

                workbook.SaveAs(output_file)

                end_time = time.time()
                duration = end_time - start_time

                print(f"\nГотово! Результат в: {output_file}")
                print(f"Токены: {translator.usage.total_tokens} | Стоимость: ${translator.total_cost_usd:.4f}")
                print(f"Общее время: {int(duration // 60)} мин. {int(duration % 60)} сек.\n")

    except Exception as e:
        print(f"\n\033[31m❌ {e}\033[0m\n")
        cleanup_excel()
        sys.exit(1)

if __name__ == "__main__":
    try:
        with keep.running():
            main()
    except KeyboardInterrupt:
        print("\n\n[STOP] Программа остановлена. Очистка ресурсов...")
        cleanup_excel()
        print("Процесс Excel успешно завершен.")
        sys.exit(0)
