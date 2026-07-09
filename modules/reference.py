# ============================================================
# modules/reference.py — чтение справочников
# ============================================================
import os
import pandas as pd
from config.settings import (
    REFERENCE_DIR, EXCLUDE_STATUS, EXCLUDE_ORDER,
    COL_NAME, COL_GROUP, COL_WEIGHT, COL_STATUS, COL_ORDER, COL_WAREHOUSE,
    NON_FOOD_WAREHOUSE
)


def _normalize(val) -> str:
    """Убирает все виды пробелов (включая неразрывные \xa0) по краям, переводит в нижний регистр."""
    if pd.isna(val):
        return ""
    # strip() убирает обычные пробелы, но не \xa0 (неразрывный пробел из Excel).
    # Явно заменяем все Unicode-пробелы через split/join, затем нижний регистр.
    import unicodedata
    s = str(val)
    # Заменяем все виды пробельных символов на обычный пробел, затем strip
    s = " ".join(s.split())  # убирает \t, \n, \r, \xa0, \u200b и т.д.
    return s.strip().lower()


def _find_col(df: pd.DataFrame, col_name: str):
    """
    Ищет колонку без учёта регистра и лишних пробелов.
    Возвращает реальное название колонки в df, или None.
    """
    target = _normalize(col_name)
    for c in df.columns:
        if _normalize(c) == target:
            return c
    return None


def load_all_references() -> dict:
    """Читает все Excel-файлы из REFERENCE_DIR. Возвращает dict: {network_code: DataFrame}"""
    refs = {}
    if not os.path.isdir(REFERENCE_DIR):
        raise FileNotFoundError(f"Папка справочников не найдена: {REFERENCE_DIR}")

    for fname in os.listdir(REFERENCE_DIR):
        if not fname.lower().endswith((".xlsx", ".xlsm")):
            continue
        fpath = os.path.join(REFERENCE_DIR, fname)
        try:
            with pd.ExcelFile(fpath) as xl:
                for sheet in xl.sheet_names:
                    df = xl.parse(sheet)
                    if df.empty:
                        continue

                    # Базовая очистка названий колонок от пробелов
                    df.columns = [str(c).strip() for c in df.columns]

                    # Проверка наличия обязательной колонки Наименование
                    if not _find_col(df, COL_NAME):
                        continue

                    # 1. Фильтрация по Статусу
                    real_status = _find_col(df, COL_STATUS)
                    if real_status and EXCLUDE_STATUS:
                        exclude_status_clean = [_normalize(x) for x in EXCLUDE_STATUS]
                        before = len(df)
                        df = df[~df[real_status].apply(lambda x: _normalize(x) in exclude_status_clean)]
                        diff = before - len(df)
                        if diff:
                            print(f"    [{sheet}] Исключено по статусу ({real_status!r}): {diff} шт.")

                    # 2. Фильтрация по Заказу
                    real_order = _find_col(df, COL_ORDER)
                    if real_order and EXCLUDE_ORDER:
                        exclude_order_clean = [_normalize(x) for x in EXCLUDE_ORDER]
                        before = len(df)
                        df = df[~df[real_order].apply(lambda x: _normalize(x) in exclude_order_clean)]
                        diff = before - len(df)
                        if diff:
                            print(f"    [{sheet}] Исключено по заказу ({real_order!r}): {diff} шт.")
                    elif not real_order:
                        print(f"    [{sheet}] ВНИМАНИЕ: колонка '{COL_ORDER}' не найдена!")
                        print(f"    [{sheet}] Доступные колонки: {list(df.columns)}")

                    refs[sheet] = df.reset_index(drop=True)
                    print(f"  [+] Загружен справочник: {fname} -> Лист: {sheet} ({len(df)} строк)")

        except Exception as e:
            print(f"  [!] Ошибка чтения {fname}: {e}")

    return refs


def get_network_items(refs: dict, network_code: str, category: str = None) -> pd.DataFrame:
    """Возвращает DataFrame товаров для сети."""
    df = refs.get(network_code)
    if df is None:
        raise ValueError(f"Справочник для сети '{network_code}' не найден.")

    real_warehouse = _find_col(df, COL_WAREHOUSE)
    if real_warehouse:
        if category == "nonfood":
            df = df[df[real_warehouse].apply(
                lambda x: _normalize(x) == _normalize(NON_FOOD_WAREHOUSE)
            )]
        elif category == "food":
            df = df[df[real_warehouse].apply(
                lambda x: _normalize(x) != _normalize(NON_FOOD_WAREHOUSE)
            )]

    # Собираем нужные колонки
    res_cols = {}
    for global_name, target in [("name", COL_NAME), ("group", COL_GROUP), ("weight", COL_WEIGHT)]:
        real_c = _find_col(df, target)
        if real_c:
            res_cols[global_name] = df[real_c]
        else:
            if global_name == "weight":
                res_cols["weight"] = 0.0
            elif global_name == "group":
                res_cols["group"] = "Общий"
            else:
                res_cols[global_name] = ""

    new_df = pd.DataFrame({
        COL_NAME: res_cols["name"],
        COL_GROUP: res_cols["group"],
        COL_WEIGHT: res_cols["weight"]
    })

    # Очистка данных
    new_df[COL_NAME] = new_df[COL_NAME].astype(str).str.strip()
    new_df[COL_GROUP] = new_df[COL_GROUP].fillna("Общий").astype(str).str.strip()
    new_df[COL_WEIGHT] = pd.to_numeric(new_df[COL_WEIGHT], errors='coerce').fillna(0.0)

    # Исключаем пустые наименования
    new_df = new_df[new_df[COL_NAME] != ""]
    return new_df.reset_index(drop=True)