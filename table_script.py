import os
import sys

import openpyxl

PREFIXES = ['PEK1E', 'MOW1H']


def get_script_dir():
    """Папка, где лежит сам .exe или .py (а не временная папка распаковки)."""
    # PyInstaller (onefile/onedir), cx_Freeze, py2exe
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))

    # Nuitka --onefile: __file__ указывает во временную папку,
    # а реальный путь к exe лежит в sys.argv[0]
    if '__compiled__' in globals():
        return os.path.dirname(os.path.abspath(sys.argv[0]))

    # обычный запуск .py
    return os.path.dirname(os.path.abspath(__file__))


def process_excel(file_path):
    try:
        # data_only=True, чтобы читать значения формул, а не сами формулы
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        changed = 0

        for row in range(1, ws.max_row + 1):
            cell_b = ws.cell(row, 2)  # столбец B
            cell_c = ws.cell(row, 3)  # столбец C

            # B или C уже заполнены — пропускаем
            if cell_b.value is not None or cell_c.value is not None:
                continue

            cell_g = ws.cell(row, 7)  # столбец G
            if cell_g.value is None:
                continue

            for line in str(cell_g.value).splitlines():
                line = line.strip()
                if not line:
                    continue

                matched_prefix = next((p for p in PREFIXES if line.startswith(p)), None)
                if matched_prefix is None:
                    continue

                parts = line[len(matched_prefix):].strip().split('/')
                if len(parts) >= 3:  # нужно минимум 3 части
                    cell_c.value = parts[0].strip()  # первое значение
                    cell_b.value = parts[2].strip()  # третье значение
                    changed += 1
                    break

        # сохраняем один раз, а не на каждой строке
        if changed:
            wb.save(file_path)

        print(f"Обработан: {os.path.basename(file_path)} (заполнено строк: {changed})")

    except PermissionError:
        print(f"Ошибка: {os.path.basename(file_path)} открыт в Excel — закройте его и запустите снова")
    except Exception as e:
        print(f"Ошибка при обработке {os.path.basename(file_path)}: {e}")


def process_all_excel_files():
    script_dir = get_script_dir()
    print(f"Папка: {script_dir}")

    excel_files = [f for f in os.listdir(script_dir)
                   if f.lower().endswith('.xlsx') and not f.startswith('~')]

    if not excel_files:
        print("В папке со скриптом не найдено .xlsx файлов")
        return

    print(f"Найдено файлов: {len(excel_files)}")
    for file in excel_files:
        process_excel(os.path.join(script_dir, file))

    print("Все файлы обработаны!")


if __name__ == "__main__":
    try:
        process_all_excel_files()
    finally:
        # чтобы окно консоли не закрывалось сразу
        input("\nНажмите Enter для выхода...")
