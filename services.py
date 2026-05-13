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


def _excel_read_engine(path: str) -> str | None:
    """Для старых .xls pandas требует движок xlrd."""
    if Path(path).suffix.lower() == ".xls":
        return "xlrd"
    return None


def read_registry_file(path: str, sheet_index: int = 0) -> pd.DataFrame:
    """
    Считывает реестр OKDP2 из Excel‑файла, начиная с 7‑й строки.
    Ожидает, что в файле присутствуют колонки «Код» и «Название».
    """
    engine = _excel_read_engine(path)
    df = pd.read_excel(path, sheet_name=sheet_index, skiprows=6, engine=engine)
    df.columns = [c.strip() for c in df.columns]
    if "Код" not in df.columns or "Название" not in df.columns:
        raise ValueError("Файл реестра должен содержать колонки 'Код' и 'Название'")
    return df[["Код", "Название"]].dropna(subset=["Код", "Название"])


def read_code_name_file(path: str) -> pd.DataFrame:
    """
    Считывает второй файл, где в одной ячейке находится объединённый код+название.
    Создаётся DataFrame с одной колонкой «Value».
    Лист выбирается по содержимому (первый лист, где в 1-й колонке есть непустые
    ячейки), без привязки к названию листа. Поддерживаются .xlsx и .xls.
    """
    path_s = str(path)
    engine = _excel_read_engine(path_s)
    xl_file = pd.ExcelFile(path_s, engine=engine)
    df: pd.DataFrame | None = None
    for sheet_name in xl_file.sheet_names:
        raw = pd.read_excel(path_s, sheet_name=sheet_name, header=None, engine=engine)
        if raw.empty or raw.shape[1] < 1:
            continue
        col0 = raw.iloc[:, 0]
        for v in col0:
            if pd.notna(v) and str(v).strip() != "":
                df = raw
                break
        if df is not None:
            break
    if df is None:
        raise ValueError(
            "Во втором файле не найден лист с непустыми значениями в первой колонке"
        )
    out = pd.DataFrame({"Value": df.iloc[:, 0].copy()})
    return out


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
