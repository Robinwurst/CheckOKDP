# services.py
"""
Модуль, содержащий вспомогательные функции для работы с Excel‑файлами.
"""

from pathlib import Path

import pandas as pd
from typing import Dict, List, Tuple

# Текст в ячейке «Объединенный», если код в реестре не найден (суффикс к коду).
REGISTRY_MISSING_SUFFIX = " — код в реестре ОКПД2 отсутствует"


def norm_cell_text(value: object) -> str:
    """Одинаковое сравнение ячеек: схлопываем пробелы по краям и между словами."""
    return " ".join(str(value).split()).strip()


def _coerce_header(val: object) -> str:
    """Убираем неразрывные пробелы и лишние пробелы в названиях колонок."""
    return " ".join(str(val).replace("\u00a0", " ").split()).strip()


def _engines_to_try(path: str) -> list[str | None]:
    """
    Порядок движков чтения: сначала типичный для расширения, затем запасной.
    calamine (Rust) часто открывает старые .xls/.xlsx, с которыми xlrd/openpyxl падают.
    """
    ext = Path(path).suffix.lower()
    seq: list[str | None]
    if ext == ".xls":
        seq = ["xlrd", "calamine"]
    elif ext == ".xlsb":
        seq = ["calamine"]
    elif ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        seq = [None, "openpyxl", "calamine"]
    else:
        seq = [None, "openpyxl", "calamine", "xlrd"]

    seen: set[str] = set()
    out: list[str | None] = []
    for e in seq:
        key = e if e is not None else "__auto__"
        if key in seen:
            continue
        seen.add(key)
        out.append(e)
    return out


def open_excel_workbook(path: str) -> tuple[pd.ExcelFile, str | None]:
    """
    Открывает книгу, перебирая движки. Не зависит от имени листа — только от файла.
    """
    path_s = str(path)
    last_exc: Exception | None = None
    for eng in _engines_to_try(path_s):
        try:
            return pd.ExcelFile(path_s, engine=eng), eng
        except Exception as e:
            last_exc = e
    assert last_exc is not None
    raise ValueError(
        "Не удалось открыть Excel-файл ни одним доступным способом. "
        "Попробуйте открыть файл в Excel и «Сохранить как» .xlsx или .xls (97–2003). "
        f"Последняя ошибка: {last_exc}"
    ) from last_exc


def read_registry_file(path: str) -> pd.DataFrame:
    """
    Считывает реестр ОКПД2: данные с 7-й строки файла, колонки «Код» и «Название».
    Лист выбирается автоматически — первый лист, где после пропуска строк есть эти колонки;
    имя листа (в т.ч. на русском) не важно.
    """
    path_s = str(path)
    xl, engine = open_excel_workbook(path_s)
    last_sheet_error: str | None = None
    for sheet in xl.sheet_names:
        try:
            df = pd.read_excel(
                path_s,
                sheet_name=sheet,
                skiprows=6,
                engine=engine,
            )
        except Exception as e:
            last_sheet_error = f"{sheet!r}: {e}"
            continue
        df.columns = [_coerce_header(c) for c in df.columns]
        if "Код" not in df.columns or "Название" not in df.columns:
            continue
        return df[["Код", "Название"]].dropna(subset=["Код", "Название"])

    hint = f" ({last_sheet_error})" if last_sheet_error else ""
    raise ValueError(
        "В реестре не найден лист с колонками «Код» и «Название» "
        "(шапка таблицы ожидается с 7-й строки файла, как в выгрузке classifikators.ru)."
        f"{hint}"
    )


def read_code_name_file(path: str) -> pd.DataFrame:
    """
    Второй файл: одна колонка «код название».
    Берётся первый лист (в порядке книги), где в первой колонке есть непустые ячейки.
    Название листа не используется. Поддержка разных версий Excel через несколько движков.
    """
    path_s = str(path)
    xl, engine = open_excel_workbook(path_s)
    for sheet in xl.sheet_names:
        try:
            raw = pd.read_excel(
                path_s, sheet_name=sheet, header=None, engine=engine
            )
        except Exception:
            continue
        if raw.empty or raw.shape[1] < 1:
            continue
        col0 = raw.iloc[:, 0]
        for v in col0:
            if pd.notna(v) and str(v).strip() != "":
                return pd.DataFrame({"Value": raw.iloc[:, 0].copy()})

    raise ValueError(
        "Во втором файле не найден лист с непустыми значениями в первой колонке"
    )


def split_code_name(cell: str) -> Tuple[str, str]:
    """
    Делит ячейку на «код» и «название».
    Предполагается, что первый пробел отделяет их.
    """
    if not isinstance(cell, str):
        cell = str(cell)
    parts = cell.split(" ", 1)
    code = parts[0].strip()
    name = parts[1].strip() if len(parts) > 1 else ""
    return code, name


def find_longest_matches(prefix: str, registry: Dict[str, str]) -> List[Tuple[str, str]]:
    """
    Возвращает все записи реестра, начинающиеся с заданного префикса,
    и среди них только те, чья длина кода максимальна.
    """
    if not prefix:
        return []
    matches = [(c, registry[c]) for c in registry if c.startswith(prefix)]
    if not matches:
        return []
    max_len = max(len(c) for c, _ in matches)
    return [(c, n) for c, n in matches if len(c) == max_len]


def merge_code_with_registry(row_code: str, registry: Dict[str, str]) -> Tuple[str, bool]:
    """
    Возвращает (строка_результата, найдено_ли_в_реестре).
    Если совпадений нет: «<код> — код в реестре ОКПД2 отсутствует», второй элемент False.
    """
    longest = find_longest_matches(row_code, registry)
    if not longest:
        return f"{row_code}{REGISTRY_MISSING_SUFFIX}", False
    return "; ".join(f"{c} {n}" for c, n in longest), True


def analyze_row(cell: object, registry: Dict[str, str]) -> Tuple[str, bool, bool]:
    """
    Разбор одной строки второго файла.

    Возвращает:
        (текст для колонки «Объединенный», найден_ли_хотя_бы_один_код_в_реестре,
         подсветить_строку).

    Подсветка: код не в реестре; несколько вариантов с максимальной длиной кода;
    один вариант, но код или наименование в файле не совпадают с записью реестра.
    """
    if not isinstance(cell, str):
        cell = str(cell) if pd.notna(cell) else ""
    code, name = split_code_name(cell)
    longest = find_longest_matches(code, registry)
    if not longest:
        return f"{code}{REGISTRY_MISSING_SUFFIX}", False, True
    merged = "; ".join(f"{c} {n}" for c, n in longest)
    if len(longest) > 1:
        return merged, True, True
    c_reg, n_reg = longest[0]
    if norm_cell_text(code) != norm_cell_text(c_reg) or norm_cell_text(name) != norm_cell_text(
        n_reg
    ):
        return merged, True, True
    return merged, True, False


def build_new_column(row_code: str, registry: Dict[str, str]) -> str:
    """
    Для строки кода формирует строку вида:
        «код1 название1; код2 название2; …»
    Если подходящих записей нет — текст с пометкой об отсутствии в реестре ОКПД2.
    """
    text, _ = merge_code_with_registry(row_code, registry)
    return text
