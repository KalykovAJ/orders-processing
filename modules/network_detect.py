# ============================================================
# modules/network_detect.py — определение сетей по папкам
# ============================================================
import os


def detect_networks(root_folder: str, refs: dict) -> dict:
    """
    Сопоставляет подпапки root_folder с сетями из справочника.
    Возвращает {folder_name: network_code}
    """
    mapping = {}
    known_codes = set(refs.keys())

    for entry in os.scandir(root_folder):
        if not entry.is_dir():
            continue
        folder = entry.name.strip()
        if folder in known_codes:
            mapping[folder] = folder
        else:
            # Попытка нечёткого сопоставления (верхний/нижний регистр)
            matched = next(
                (code for code in known_codes if code.lower() == folder.lower()),
                None
            )
            if matched:
                mapping[folder] = matched

    return mapping  # {folder_name: network_code}


def get_azs_files(network_folder: str) -> list:
    """
    Возвращает список Excel-файлов заявочников в папке сети.
    """
    files = []
    if not os.path.isdir(network_folder):
        return files
    for f in os.listdir(network_folder):
        if f.lower().endswith((".xlsx", ".xlsm", ".xls")):
            files.append(os.path.join(network_folder, f))
    return files


def parse_azs_number(filename: str, is_mp: bool = False) -> tuple:
    """
    Парсит имя файла заявочника.
    Для МП: возвращает (azs_number, category) где category 'food'|'nonfood'
    Для остальных: (azs_number, None)
    """
    base = os.path.splitext(os.path.basename(filename))[0].strip()
    if is_mp:
        if base.lower().endswith("ф"):
            return base[:-1].strip(), "food"
        elif base.lower().endswith("н"):
            return base[:-1].strip(), "nonfood"
        else:
            return base, None
    return base, None