#!/usr/bin/env python3
"""
Объединение прайс-листов от разных поставщиков в один сводный Excel-файл.

Каждый поставщик присылает файл со своей структурой: свой порядок столбцов,
свои названия колонок ("Наименование" / "Товар" / "Name"), свой формат цены
("1200", "1 200,00 ₽", "1200.00"). Скрипт приводит все файлы к единой схеме
по правилам из JSON-конфига и собирает их в один файл.

Примеры запуска:
    python merge_pricelists.py --input-dir sample_data --config mapping_config.json
    python merge_pricelists.py --files a.xlsx b.xlsx -c mapping_config.json -o out.xlsx
"""

import argparse
import fnmatch
import json
import os
import re
import sys

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# =====================================================================
# Единая схема сводного файла
# =====================================================================

# Ключ в конфиге -> название столбца в сводном файле.
# Порядок словаря = порядок столбцов в результате.
TARGET_FIELDS = {
    "name": "Наименование",
    "sku": "Артикул",
    "price": "Цена",
    "qty": "Количество",
}

SUPPLIER_COLUMN = "Поставщик"
SOURCE_COLUMN = "Источник (файл)"
DUPLICATE_COLUMN = "Дубль между поставщиками"

# Значения, которые пишутся в столбец-флаг дублей
DUPLICATE_YES = "да"
DUPLICATE_NO = ""


# =====================================================================
# Аргументы командной строки
# =====================================================================

def parse_args():
    """Разбирает аргументы командной строки."""
    parser = argparse.ArgumentParser(
        description="Объединяет прайсы от разных поставщиков в один сводный Excel-файл.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Примеры:\n"
            "  python merge_pricelists.py --input-dir sample_data\n"
            "  python merge_pricelists.py --files sample_data/alpha_price.xlsx "
            "sample_data/beta_price.xlsx -o output/svod.xlsx\n"
        ),
    )
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument(
        "--input-dir", "-i",
        help="папка с исходными .xlsx файлами (берутся все файлы из папки)",
    )
    source.add_argument(
        "--files", "-f", nargs="+",
        help="список путей к .xlsx файлам вместо папки",
    )
    parser.add_argument(
        "--config", "-c", default="mapping_config.json",
        help="JSON-конфиг с маппингом столбцов (по умолчанию mapping_config.json)",
    )
    parser.add_argument(
        "--output", "-o", default="output/svodny_price.xlsx",
        help="путь к сводному файлу (по умолчанию output/svodny_price.xlsx)",
    )
    return parser.parse_args()


# =====================================================================
# Конфиг маппинга
# =====================================================================

def load_config(path):
    """Читает JSON-конфиг и проверяет, что он правильно устроен.

    Ожидаемая структура:
        {"files": {"имя_файла.xlsx": {"supplier": "...",
                                      "columns": {"name": "...", "sku": "...",
                                                  "price": "...", "qty": "..."}}}}
    """
    if not os.path.exists(path):
        sys.exit(f"Ошибка: файл конфига не найден: {path}")

    try:
        with open(path, encoding="utf-8") as f:
            config = json.load(f)
    except json.JSONDecodeError as e:
        sys.exit(
            f"Ошибка: конфиг {path} — некорректный JSON "
            f"(строка {e.lineno}, символ {e.colno}): {e.msg}"
        )
    except OSError as e:
        sys.exit(f"Ошибка: не удалось прочитать конфиг {path}: {e}")

    files = config.get("files")
    if not isinstance(files, dict) or not files:
        sys.exit(
            f"Ошибка: в конфиге {path} нет блока \"files\" со списком файлов. "
            "Пример структуры смотрите в README."
        )

    for file_name, block in files.items():
        if not isinstance(block, dict):
            sys.exit(f"Ошибка: в конфиге блок для файла '{file_name}' должен быть объектом {{...}}")

        columns = block.get("columns")
        if not isinstance(columns, dict):
            sys.exit(f"Ошибка: в конфиге у файла '{file_name}' нет блока \"columns\" с маппингом столбцов")

        missing = [field for field in TARGET_FIELDS if field not in columns]
        if missing:
            sys.exit(
                f"Ошибка: в конфиге у файла '{file_name}' не заданы поля: {', '.join(missing)}. "
                f"Нужны все: {', '.join(TARGET_FIELDS)}"
            )

    return files


def get_mapping_for_file(mappings, path):
    """Находит в конфиге блок маппинга для конкретного файла.

    Сначала ищется точное совпадение имени файла, потом — ключи-шаблоны
    вида "beta_price*.xlsx" (без учёта регистра): поставщики часто
    присылают файлы с датой в имени.
    """
    file_name = os.path.basename(path)
    if file_name in mappings:
        return mappings[file_name]

    matches = [
        pattern
        for pattern in mappings
        if fnmatch.fnmatchcase(file_name.lower(), pattern.lower())
    ]
    if len(matches) > 1:
        sys.exit(
            f"Ошибка: файл '{file_name}' подходит сразу под несколько шаблонов "
            f"в конфиге: {', '.join(sorted(matches))}.\n"
            "       Уточните шаблоны, чтобы файлу соответствовал только один."
        )
    if not matches:
        sys.exit(
            f"Ошибка: в конфиге нет маппинга для файла '{file_name}'.\n"
            f"       Добавьте блок \"{file_name}\" в раздел \"files\" конфига\n"
            "       или ключ-шаблон вида \"поставщик_*.xlsx\" для файлов с датой в имени.\n"
            f"       Сейчас в конфиге описаны: {', '.join(sorted(mappings))}"
        )
    return mappings[matches[0]]


# =====================================================================
# Поиск исходных файлов
# =====================================================================

def collect_source_files(args):
    """Возвращает отсортированный список путей к исходным .xlsx файлам."""
    if args.input_dir:
        if not os.path.isdir(args.input_dir):
            sys.exit(f"Ошибка: папка с исходными файлами не найдена: {args.input_dir}")
        paths = [
            os.path.join(args.input_dir, name)
            for name in sorted(os.listdir(args.input_dir))
            # ~$ — временные файлы, которые Excel создаёт для открытых книг
            if name.lower().endswith(".xlsx") and not name.startswith("~$")
        ]
        if not paths:
            sys.exit(f"Ошибка: в папке {args.input_dir} нет ни одного .xlsx файла")
        return paths

    paths = []
    for path in args.files:
        if not os.path.exists(path):
            sys.exit(f"Ошибка: файл не найден: {path}")
        if not path.lower().endswith(".xlsx"):
            sys.exit(f"Ошибка: поддерживаются только .xlsx файлы, получен: {path}")
        paths.append(path)
    return paths


# =====================================================================
# Чтение одного файла и приведение его к единой схеме
# =====================================================================

def read_source_file(path, block):
    """Читает один Excel-файл как есть.

    dtype=object: текстовые ячейки остаются строками, а числовые — числами,
    поэтому ячейка со значением 3.125 не попадает под правило неоднозначности
    "1.200" в parse_number.

    Из блока конфига берутся необязательные параметры:
        "sheet"      — имя листа (по умолчанию первый лист);
        "skip_rows"  — сколько строк пропустить сверху до строки заголовков.
    """
    sheet = block.get("sheet", 0)
    skip_rows = block.get("skip_rows", 0)

    try:
        return pd.read_excel(path, sheet_name=sheet, skiprows=skip_rows, dtype=object)
    except Exception as e:
        # Ловим любую ошибку чтения и показываем понятный текст вместо трейсбека.
        # Самые частые причины — нет листа с таким именем или файл повреждён.
        if "Worksheet" in str(e):
            sys.exit(
                f"Ошибка: в файле {path} нет листа '{sheet}'.\n"
                "       Проверьте параметр \"sheet\" в конфиге."
            )
        sys.exit(
            f"Ошибка: файл {path} не читается как Excel — возможно, он повреждён "
            f"или сохранён в другом формате (.xls, .csv).\n"
            f"       Подробности: {e}"
        )


def apply_mapping(df, block, path):
    """Оставляет только нужные столбцы и переименовывает их по единой схеме."""
    columns = block["columns"]
    file_name = os.path.basename(path)

    # В заголовках бывают переносы строк ("Цена,\nруб."), двойные
    # и неразрывные пробелы — приводим их к виду, в котором пишут конфиг
    df = df.rename(columns=normalize_text)

    missing = [
        source_column
        for source_column in columns.values()
        if normalize_text(source_column) not in df.columns
    ]
    if missing:
        sys.exit(
            f"Ошибка: в файле '{file_name}' нет столбцов, указанных в конфиге: "
            f"{', '.join(missing)}.\n"
            f"       В файле есть столбцы: {', '.join(str(c) for c in df.columns)}"
        )

    result = pd.DataFrame()
    for field, target_column in TARGET_FIELDS.items():
        result[target_column] = df[normalize_text(columns[field])]

    result[SUPPLIER_COLUMN] = block.get("supplier", file_name)
    result[SOURCE_COLUMN] = file_name
    return result


# =====================================================================
# Нормализация значений
# =====================================================================

def normalize_text(value):
    """Обрезает пробелы по краям и схлопывает повторяющиеся пробелы внутри."""
    if pd.isna(value):
        return ""
    # \xa0 — неразрывный пробел, часто приезжает из выгрузок и из Word
    text = str(value).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def normalize_sku(value):
    """Приводит артикул к единому виду: без пробелов, в верхнем регистре.

    Регистр важен: 'mouse-w-101' и 'MOUSE-W-101' — один и тот же товар,
    без этого пересечения между поставщиками не найдутся.
    """
    return normalize_text(value).upper()


# Число внутри строки: начинается и заканчивается цифрой, внутри — цифры,
# пробелы, точки и запятые. Так точка из "руб." или "шт." в число не попадает.
NUMBER_PATTERN = re.compile(r"-?\d(?:[\d .,]*\d)?")


def parse_number(value):
    """Превращает цену/количество в число.

    Понимает: 1200, "1200.00", "1 200,00 ₽", "1 340,00 руб.", "1.340,00",
    "1,200.00", "1.200.000", "15 шт".
    Возвращает None, если распознать не удалось или запись неоднозначна:
    "1,200" может быть и 1200, и 1.2 — такое не угадываем, а отдаём в отчёт.
    """
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).replace("\xa0", " ")
    numbers = NUMBER_PATTERN.findall(text)
    # Ни одного числа ("под заказ") или несколько ("10-20 шт", "1.2e3") —
    # какое из них имелось в виду, неизвестно
    if len(numbers) != 1:
        return None
    return parse_number_token(numbers[0])


def parse_number_token(token):
    """Разбирает одно число вида "1 200 000,50" с учётом разделителей."""
    sign = -1 if token.startswith("-") else 1
    token = token.lstrip("-")

    decimal = find_decimal_separator(token)
    if decimal == "?":
        return None

    if decimal:
        integer_part, fraction = token.rsplit(decimal, 1)
    else:
        integer_part, fraction = token, ""
    if fraction and not fraction.isdigit():
        # "340,00 5" — после десятичного разделителя только цифры
        return None

    # Всё, что осталось в целой части, кроме цифр, — разделители тысяч:
    # пробелы и тот знак, который не десятичный. Группы должны быть по 3 цифры,
    # иначе это не разделители, а мусор ("1.2.3")
    groups = re.split(r"[ .,]+", integer_part)
    if len(groups) > 1 and (
        not 1 <= len(groups[0]) <= 3
        or any(len(group) != 3 for group in groups[1:])
    ):
        return None

    return sign * float("".join(groups) + "." + (fraction or "0"))


def find_decimal_separator(token):
    """Определяет десятичный разделитель в числе.

    Возвращает "." или ",", пустую строку, если дробной части нет,
    и "?", если запись неоднозначна или некорректна.
    """
    has_dot = "." in token
    has_comma = "," in token

    if has_dot and has_comma:
        # "1,200.00" и "1.200,00": десятичный — тот, что стоит последним
        decimal = "." if token.rfind(".") > token.rfind(",") else ","
        return decimal if token.count(decimal) == 1 else "?"

    if not has_dot and not has_comma:
        return ""

    separator = "." if has_dot else ","
    if token.count(separator) > 1:
        # "1.200.000" — один и тот же знак несколько раз: разделитель тысяч
        return ""

    integer_part, fraction = token.split(separator)
    if len(fraction) == 3 and integer_part.strip() != "0":
        # "1,200" или "1.200": то ли тысячи, то ли дробь — не угадываем.
        # Целая часть "0" ("0,125") разделителем тысяч быть не может
        return "?"
    return separator


def normalize_frame(df):
    """Нормализует все столбцы одного прайса.

    Возвращает df и счётчики: сколько цен и количеств не удалось распознать.
    """
    df = df.copy()
    df[TARGET_FIELDS["name"]] = df[TARGET_FIELDS["name"]].apply(normalize_text)
    df[TARGET_FIELDS["sku"]] = df[TARGET_FIELDS["sku"]].apply(normalize_sku)

    prices = df[TARGET_FIELDS["price"]].apply(parse_number)
    quantities = df[TARGET_FIELDS["qty"]].apply(parse_number)

    # Нераспознанной считаем только непустую ячейку: пробелы — это пустое значение
    has_price = df[TARGET_FIELDS["price"]].apply(normalize_text) != ""
    has_quantity = df[TARGET_FIELDS["qty"]].apply(normalize_text) != ""
    bad_prices = int((prices.isna() & has_price).sum())
    bad_quantities = int((quantities.isna() & has_quantity).sum())

    df[TARGET_FIELDS["price"]] = prices.round(2)
    df[TARGET_FIELDS["qty"]] = quantities

    return df, bad_prices, bad_quantities


def drop_empty_rows(df):
    """Убирает строки без названия и без артикула — это пустые строки прайса."""
    is_empty = (df[TARGET_FIELDS["name"]] == "") & (df[TARGET_FIELDS["sku"]] == "")
    return df[~is_empty].copy(), int(is_empty.sum())


# =====================================================================
# Объединение и поиск пересечений по артикулу
# =====================================================================

def merge_frames(frames):
    """Складывает прайсы всех поставщиков в одну таблицу."""
    merged = pd.concat(frames, ignore_index=True)
    return merged


def mark_cross_supplier_duplicates(df):
    """Помечает строки, у которых артикул встречается у нескольких поставщиков.

    Строки не удаляются — заказчик сам решает, что с ними делать.
    Возвращает df со столбцом-флагом и словарь {артикул: [поставщики]}.
    """
    df = df.copy()
    sku_column = TARGET_FIELDS["sku"]

    # Сколько РАЗНЫХ поставщиков у каждого артикула
    suppliers_per_sku = df.groupby(sku_column)[SUPPLIER_COLUMN].nunique()
    cross_skus = suppliers_per_sku[suppliers_per_sku > 1].index

    # Пустой артикул сравнивать не с чем
    is_duplicate = df[sku_column].isin(cross_skus) & (df[sku_column] != "")
    df[DUPLICATE_COLUMN] = is_duplicate.map({True: DUPLICATE_YES, False: DUPLICATE_NO})

    overlaps = {}
    for sku in cross_skus:
        if sku == "":
            continue
        suppliers = df.loc[df[sku_column] == sku, SUPPLIER_COLUMN].unique().tolist()
        overlaps[sku] = sorted(suppliers)

    return df, overlaps


def sort_result(df):
    """Сортирует сводную таблицу: сначала артикул, внутри — поставщик.

    Так строки-дубли по одному артикулу стоят рядом и их удобно сравнивать.
    """
    return df.sort_values(
        by=[TARGET_FIELDS["sku"], SUPPLIER_COLUMN], kind="stable"
    ).reset_index(drop=True)


# =====================================================================
# Сохранение результата
# =====================================================================

def save_result(df, path):
    """Сохраняет сводную таблицу в Excel и оформляет её для чтения глазами."""
    out_dir = os.path.dirname(path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Сводный прайс")
            format_sheet(writer.sheets["Сводный прайс"], df)
    except PermissionError:
        sys.exit(
            f"Ошибка: нет доступа для записи файла {path} "
            "(возможно, он открыт в Excel)."
        )
    except OSError as e:
        sys.exit(f"Ошибка: не удалось сохранить файл {path}: {e}")


def format_sheet(worksheet, df):
    """Жирный заголовок, закреплённая шапка, ширина столбцов, подсветка дублей."""
    highlight = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")
    worksheet.freeze_panes = "A2"

    for index, column in enumerate(df.columns, start=1):
        values = df[column].astype(str)
        width = max(len(str(column)), int(values.str.len().max() or 0)) + 2
        worksheet.column_dimensions[get_column_letter(index)].width = min(width, 45)

    # Цена — всегда два знака после запятой, чтобы столбец читался как прайс
    price_index = list(df.columns).index(TARGET_FIELDS["price"]) + 1
    for row_number in range(2, len(df) + 2):
        worksheet.cell(row=row_number, column=price_index).number_format = "#,##0.00"

    # Подсветка строк, помеченных флагом дубля (строка 1 — заголовок)
    for row_number, flag in enumerate(df[DUPLICATE_COLUMN], start=2):
        if flag == DUPLICATE_YES:
            for column_number in range(1, len(df.columns) + 1):
                worksheet.cell(row=row_number, column=column_number).fill = highlight


# =====================================================================
# Отчёт в консоль
# =====================================================================

def print_report(file_stats, total_rows, overlaps, duplicate_rows, output_path):
    """Печатает итоговый отчёт о работе скрипта."""
    print()
    print("=== Отчёт об объединении прайсов ===")
    print(f"Обработано файлов: {len(file_stats)}")
    for stat in file_stats:
        print(
            f"  - {stat['file']} ({stat['supplier']}): "
            f"строк взято {stat['rows']}"
        )
        if stat["empty_rows"]:
            print(f"      пропущено пустых строк: {stat['empty_rows']}")
        if stat["bad_prices"]:
            print(f"      цен не распознано: {stat['bad_prices']}")
        if stat["bad_quantities"]:
            print(f"      количеств не распознано: {stat['bad_quantities']}")

    print(f"Всего строк в сводном файле: {total_rows}")
    print(f"Артикулов, встречающихся у разных поставщиков: {len(overlaps)}")
    print(f"Строк с флагом «{DUPLICATE_COLUMN}»: {duplicate_rows}")

    for sku, suppliers in sorted(overlaps.items()):
        print(f"  - {sku}: {', '.join(suppliers)}")

    print(f"Результат сохранён: {output_path}")


# =====================================================================
# Главная функция
# =====================================================================

def main():
    args = parse_args()
    mappings = load_config(args.config)
    paths = collect_source_files(args)

    frames = []
    file_stats = []

    for path in paths:
        block = get_mapping_for_file(mappings, path)
        raw = read_source_file(path, block)
        mapped = apply_mapping(raw, block, path)
        normalized, bad_prices, bad_quantities = normalize_frame(mapped)
        cleaned, empty_rows = drop_empty_rows(normalized)

        frames.append(cleaned)
        file_stats.append({
            "file": os.path.basename(path),
            "supplier": block.get("supplier", os.path.basename(path)),
            "rows": len(cleaned),
            "empty_rows": empty_rows,
            "bad_prices": bad_prices,
            "bad_quantities": bad_quantities,
        })

    merged = merge_frames(frames)
    merged, overlaps = mark_cross_supplier_duplicates(merged)
    merged = sort_result(merged)

    save_result(merged, args.output)

    duplicate_rows = int((merged[DUPLICATE_COLUMN] == DUPLICATE_YES).sum())
    print_report(file_stats, len(merged), overlaps, duplicate_rows, args.output)


if __name__ == "__main__":
    main()
