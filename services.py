# services.py
"""
Модуль, содержащий вспомогательные функции для работы с Excel‑файлами.
"""

import pandas as pd
from typing import Dict, List, Tuple

# Текст в ячейке «Объединенный», если код в реестре не найден (суффикс к коду).
REGISTRY_MISSING_SUFFIX = " — код в реестре ОКПД2 отсутствует"


def read_registry_file(path: str, sheet_index: int = 0) -> pd.DataFrame:
    """
    Считывает реестр OKDP2 из Excel‑файла, начиная с 7‑й строки.
    Ожидает, что в файле присутствуют колонки «Код» и «Название».
    """
    df = pd.read_excel(path, sheet_name=sheet_index, skiprows=6)
    df.columns = [c.strip() for c in df.columns]
    if "Код" not in df.columns or "Название" not in df.columns:
        raise ValueError("Файл реестра должен содержать колонки 'Код' и 'Название'")
    return df[["Код", "Название"]].dropna(subset=["Код", "Название"])


def read_code_name_file(path: str, sheet_index: int = 0) -> pd.DataFrame:
    """
    Считывает второй файл, где в одной ячейке находится объединённый код+название.
    Создаётся DataFrame с одной колонкой «Value».
    """
    df = pd.read_excel(path, sheet_name=sheet_index, header=None)
    df.columns = ["Value"]
    return df


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


def build_new_column(row_code: str, registry: Dict[str, str]) -> str:
    """
    Для строки кода формирует строку вида:
        «код1 название1; код2 название2; …»
    Если подходящих записей нет — текст с пометкой об отсутствии в реестре ОКПД2.
    """
    text, _ = merge_code_with_registry(row_code, registry)
    return text
