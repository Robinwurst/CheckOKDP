# services.py
"""
Модуль, содержащий вспомогательные функции для работы с Excel‑файлами.
"""

import io
import os
import shutil
import tempfile
from pathlib import Path
from typing import Callable, Dict, List, Tuple

import pandas as pd

# Текст в ячейке «Объединенный», если код в реестре не найден (суффикс к коду).
REGISTRY_MISSING_SUFFIX = " — код в реестре ОКПД2 отсутствует"


def norm_cell_text(value: object) -> str:
    """Одинаковое сравнение ячеек: схлопываем пробелы по краям и между словами."""
    return " ".join(str(value).split()).strip()


def _coerce_header(val: object) -> str:
    """Убираем неразрывные пробелы и лишние пробелы в названиях колонок."""
    return " ".join(str(val).replace("\u00a0", " ").split()).strip()


OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _peek_file_head(path: str, nbytes: int = 16384) -> bytes:
    with open(path, "rb") as f:
        return f.read(nbytes)


def sniff_excel_container(path: str) -> str:
    """
    Определение контейнера по байтам: расширение часто не совпадает с реальным форматом
    (например «.xls», который на самом деле ZIP/xlsx, или HTML-выгрузка).
    """
    try:
        blob = _peek_file_head(path)
    except OSError:
        return "unknown"
    if not blob:
        return "empty"
    if blob[:4] == b"PK\x03\x04" or blob[:2] == b"PK":
        return "zip_ooxml"
    if blob[:8] == OLE_MAGIC:
        return "ole_biff"
    stripped = blob.lstrip()
    if stripped.startswith(b"<?xml") or b":Workbook" in blob[:8000] or b"<ss:Workbook" in blob[:8000]:
        return "xml_ss"
    low = blob[:8192].lower()
    if b"<html" in low or (b"<table" in low and b"<ss:" not in blob[:3000]):
        return "html_like"
    return "unknown"


def _dedupe_engines(seq: list[str | None]) -> list[str | None]:
    seen: set[str] = set()
    out: list[str | None] = []
    for e in seq:
        k = e if e is not None else "__auto__"
        if k in seen:
            continue
        seen.add(k)
        out.append(e)
    return out


def engines_for_file(path: str) -> list[str | None]:
    """Порядок движков с учётом реального типа файла и расширения."""
    ext = Path(path).suffix.lower()
    kind = sniff_excel_container(path)
    seq: list[str | None] = []

    if kind == "zip_ooxml":
        seq.extend([None, "openpyxl", "calamine"])
    elif kind == "ole_biff":
        seq.extend(["xlrd", "calamine"])
    elif kind == "xml_ss":
        seq.extend(["calamine", None, "openpyxl", "xlrd"])
    elif kind == "html_like":
        seq.extend(["calamine", "xlrd", None, "openpyxl"])
    else:
        seq.extend([None, "openpyxl", "calamine", "xlrd"])

    if ext == ".xlsb":
        seq = ["calamine", None, "openpyxl"] + seq
    elif ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        if kind == "ole_biff":
            seq = ["xlrd", "calamine", None, "openpyxl"] + seq
    elif ext == ".xls":
        if kind == "zip_ooxml":
            seq = [None, "openpyxl", "calamine", "xlrd"] + seq
        elif kind == "xml_ss":
            seq = ["calamine", None, "openpyxl", "xlrd"] + seq

    return _dedupe_engines(seq)


def _clone_to_tempfile(original: str) -> str:
    p = Path(original)
    suf = p.suffix if p.suffix else ".bin"
    fd, dest = tempfile.mkstemp(prefix="checkokdp_", suffix=suf)
    os.close(fd)
    shutil.copyfile(original, dest)
    return dest


def _excel_read_sources(path: str) -> list[tuple[str, bool]]:
    """Пары (путь, это_временная_копия). Сначала оригинал, затем байтовая копия."""
    p = str(Path(path).expanduser().resolve(strict=False))
    out: list[tuple[str, bool]] = [(p, False)]
    try:
        out.append((_clone_to_tempfile(p), True))
    except OSError:
        pass
    return out


def _cleanup_temp(path: str, is_temp: bool) -> None:
    if not is_temp:
        return
    try:
        Path(path).unlink(missing_ok=True)
    except OSError:
        pass


def _open_excel_file(path: str) -> tuple[pd.ExcelFile, str | None]:
    last_exc: Exception | None = None
    for eng in engines_for_file(path):
        try:
            return pd.ExcelFile(path, engine=eng), eng
        except Exception as e:
            last_exc = e
    assert last_exc is not None
    raise last_exc


def _read_excel_sheet(path: str, sheet_name: str | int, **kwargs) -> pd.DataFrame:
    """Читает лист, перебирая движки."""
    last_exc: Exception | None = None
    for eng in engines_for_file(path):
        try:
            return pd.read_excel(path, sheet_name=sheet_name, engine=eng, **kwargs)
        except Exception as e:
            last_exc = e
    assert last_exc is not None
    raise last_exc


def open_excel_workbook(path: str) -> tuple[pd.ExcelFile, str | None]:
    """
    Открывает книгу. Для чтения реестра и второго файла лучше вызывать
    read_registry_file / read_code_name_file — там копия файла и запасные пути.
    """
    return _open_excel_file(str(Path(path).expanduser().resolve(strict=False)))


def _read_html_tables(path: str) -> list[pd.DataFrame]:
    """HTML-таблицы из «Сохранить как веб-страницу» / поддельного .xls."""
    raw = Path(path).read_bytes()
    last_exc: Exception | None = None
    for enc in ("utf-8-sig", "utf-8", "cp1251", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            dfs = pd.read_html(io.StringIO(text))
            if dfs:
                return dfs
        except Exception as e:
            last_exc = e
    try:
        dfs = pd.read_html(path)
        if dfs:
            return dfs
    except Exception as e:
        last_exc = e
    assert last_exc is not None
    raise ValueError(f"Не удалось разобрать как HTML-таблицу: {last_exc}") from last_exc


def _registry_from_dataframe(df: pd.DataFrame) -> pd.DataFrame | None:
    df2 = df.copy()
    df2.columns = [_coerce_header(c) for c in df2.columns]
    if "Код" not in df2.columns or "Название" not in df2.columns:
        return None
    return df2[["Код", "Название"]].dropna(subset=["Код", "Название"])


def _read_registry_from_html(path: str) -> pd.DataFrame:
    for tbl in _read_html_tables(path):
        if tbl.empty or tbl.shape[1] < 2:
            continue
        hdr_row = min(20, len(tbl) - 1)
        for header_ix in range(hdr_row + 1):
            chunk = tbl.iloc[header_ix:].reset_index(drop=True)
            if chunk.shape[0] < 2:
                continue
            cols = [_coerce_header(x) for x in chunk.iloc[0]]
            body = chunk.iloc[1:].copy()
            body.columns = cols
            got = _registry_from_dataframe(body)
            if got is not None and not got.empty:
                return got
        try:
            df_try = tbl.copy()
            df_try.columns = [_coerce_header(c) for c in df_try.columns]
            got = _registry_from_dataframe(df_try)
            if got is not None and not got.empty:
                return got
        except Exception:
            continue
    raise ValueError(
        "В HTML-таблице не найдены колонки «Код» и «Название». "
        "Сохраните реестр из Excel как .xlsx или скопируйте лист в новую книгу."
    )


def _read_codename_from_html(path: str) -> pd.DataFrame:
    for tbl in _read_html_tables(path):
        if tbl.empty or tbl.shape[1] < 1:
            continue
        col0 = tbl.iloc[:, 0]
        for v in col0:
            if pd.notna(v) and str(v).strip() != "":
                return pd.DataFrame({"Value": tbl.iloc[:, 0].copy()})
    raise ValueError("В HTML не найдена первая колонка с данными «код название».")


def _read_registry_from_excel_path(read_path: str) -> pd.DataFrame:
    xl, preferred_eng = _open_excel_file(read_path)
    last_sheet_error: str | None = None
    try:
        for sheet in xl.sheet_names:
            try:
                df = _read_excel_sheet(
                    read_path, sheet_name=sheet, skiprows=6, header=0
                )
            except Exception as e:
                last_sheet_error = f"{sheet!r}: {e}"
                continue
            got = _registry_from_dataframe(df)
            if got is not None and not got.empty:
                return got
            try:
                df2 = pd.read_excel(
                    read_path,
                    sheet_name=sheet,
                    skiprows=6,
                    engine=preferred_eng,
                    header=0,
                )
            except Exception:
                continue
            got = _registry_from_dataframe(df2)
            if got is not None and not got.empty:
                return got
    finally:
        try:
            xl.close()
        except Exception:
            pass

    hint = f" ({last_sheet_error})" if last_sheet_error else ""
    raise ValueError(
        "В реестре не найден лист с колонками «Код» и «Название» "
        "(шапка таблицы ожидается с 7-й строки файла, как в выгрузке classifikators.ru)."
        f"{hint}"
    )


def _read_codename_from_excel_path(read_path: str) -> pd.DataFrame:
    xl, preferred_eng = _open_excel_file(read_path)
    try:
        for sheet in xl.sheet_names:
            try:
                raw = _read_excel_sheet(read_path, sheet_name=sheet, header=None)
            except Exception:
                try:
                    raw = pd.read_excel(
                        read_path,
                        sheet_name=sheet,
                        header=None,
                        engine=preferred_eng,
                    )
                except Exception:
                    continue
            if raw.empty or raw.shape[1] < 1:
                continue
            col0 = raw.iloc[:, 0]
            for v in col0:
                if pd.notna(v) and str(v).strip() != "":
                    return pd.DataFrame({"Value": raw.iloc[:, 0].copy()})
    finally:
        try:
            xl.close()
        except Exception:
            pass
    raise ValueError(
        "Во втором файле не найден лист с непустыми значениями в первой колонке"
    )


def _run_with_file_variants(path: str, reader: Callable[[str], pd.DataFrame]) -> pd.DataFrame:
    """Оригинал и временная копия (часто «лечит» битые OLE/блокировки, как после Save As в Excel)."""
    agg: list[str] = []
    for read_path, is_temp in _excel_read_sources(path):
        try:
            return reader(read_path)
        except Exception as e:
            agg.append(f"{Path(read_path).name}: {e}")
        finally:
            _cleanup_temp(read_path, is_temp)
    raise ValueError(
        "Не удалось прочитать Excel. Часто помогает открыть файл в Excel и "
        "«Сохранить как» новый .xlsx. Детали:\n"
        + "\n".join(agg[:8])
    )


def _file_looks_like_markup(path: str) -> bool:
    """Не пытаемся гонять read_html по чистому бинарнику (это долго и бессмысленно)."""
    kind = sniff_excel_container(path)
    if kind in ("html_like", "xml_ss"):
        return True
    blob = _peek_file_head(path, 32768).lower()
    return b"<table" in blob or b"<html" in blob or b"<?xml" in blob


def read_registry_file(path: str) -> pd.DataFrame:
    """
    Считывает реестр ОКПД2: данные с 7-й строки файла, колонки «Код» и «Название».
    Лист выбирается автоматически. Имя листа не важно.
    Устойчиво к неверному расширению, старым .xls и HTML-выгрузкам.
    """
    path_base = str(Path(path).expanduser().resolve(strict=False))

    try:
        return _run_with_file_variants(path_base, _read_registry_from_excel_path)
    except Exception as e_excel:
        html_err: str | None = None
        if _file_looks_like_markup(path_base):
            try:
                return _read_registry_from_html(path_base)
            except Exception as e_html:
                html_err = str(e_html)
        msg = str(e_excel)
        if html_err:
            msg += f"\n\nЗапасной разбор HTML/XML: {html_err}"
        raise ValueError(msg) from e_excel


def read_code_name_file(path: str) -> pd.DataFrame:
    """
    Второй файл: одна колонка «код название». Лист с данными ищется автоматически.
    """
    path_base = str(Path(path).expanduser().resolve(strict=False))

    try:
        return _run_with_file_variants(path_base, _read_codename_from_excel_path)
    except Exception as e_excel:
        html_err: str | None = None
        if _file_looks_like_markup(path_base):
            try:
                return _read_codename_from_html(path_base)
            except Exception as e_html:
                html_err = str(e_html)
        msg = str(e_excel)
        if html_err:
            msg += f"\n\nЗапасной разбор HTML: {html_err}"
        raise ValueError(msg) from e_excel


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
