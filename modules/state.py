# ============================================================
# modules/state.py — отслеживание состояния (реестр частей 1С)
# ============================================================
import os
import json

STATE_FILE = ".azs_state.json"


def _state_path(root_folder: str) -> str:
    return os.path.join(root_folder, STATE_FILE)


def load_state(root_folder: str) -> dict:
    path = _state_path(root_folder)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            try:
                state = json.load(f)
            except Exception:
                state = {}
            if not isinstance(state, dict):
                state = {}
            if "parts" not in state:
                state["parts"] = {}
            if "1c_parts" not in state:
                state["1c_parts"] = 0
            if "known_files" not in state:
                state["known_files"] = {}
            return state
    return {"1c_parts": 0, "known_files": {}, "parts": {}}


def save_state(root_folder: str, state: dict):
    path = _state_path(root_folder)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def next_1c_part(root_folder: str) -> int:
    """Увеличивает счётчик частей 1С и возвращает новый номер."""
    state = load_state(root_folder)
    part = state.get("1c_parts", 0) + 1
    state["1c_parts"] = part
    save_state(root_folder, state)
    return part


def reset_1c_parts(root_folder: str):
    """Сбрасывает счётчик частей и снимок заявок (полный сброс для нового периода)."""
    state = load_state(root_folder)
    state["1c_parts"] = 0
    state["known_files"] = {}
    state["parts"] = {}
    save_state(root_folder, state)


# ── Отслеживание заявочников между запусками ─────────────────

def get_known_files(root_folder: str) -> dict:
    """Возвращает плоский снимок заявочников, сохранённый при прошлом запуске."""
    state = load_state(root_folder)
    return state.get("known_files", {})


def save_known_files(root_folder: str, known_files: dict):
    """Сохраняет текущий плоский снимок заявочников в state."""
    state = load_state(root_folder)
    state["known_files"] = known_files
    save_state(root_folder, state)


def sync_known_files(root_folder: str, current_files: dict) -> dict:
    """Синхронизирует known_files с текущим состоянием файловой системы."""
    known = get_known_files(root_folder)
    synced = {}
    for net_code, known_list in known.items():
        current_set = set(current_files.get(net_code, []))
        filtered = [f for f in known_list if f in current_set]
        if filtered:
            synced[net_code] = filtered
    return synced


def find_unsynced_orders(root_folder: str, networks: dict) -> dict:
    """
    Сверяет реестр зарегистрированных заявок (state["parts"]) с тем,
    что реально лежит на диске сейчас. Используется главным меню
    (main.py) для блокировки операций «Разделить по группам» и
    «Посчитать тоннаж», пока заявки не будут пересобраны заново
    (через п.1 или «Запустить все»).

    Возвращает {"deleted": [...], "new": [...]} — строки вида
    '[СЕТЬ] имя_файла.xlsx'.

    Если операция 1 ещё ни разу не запускалась (реестр пуст) —
    возвращает оба списка пустыми: проверять пока нечего, и эта
    ситуация уже отдельно обрабатывается проверкой наличия шаблонов.
    """
    state = load_state(root_folder)
    parts_registry = state.get("parts", {})
    if not parts_registry:
        return {"deleted": [], "new": []}

    registered: dict = {}
    for p_files in parts_registry.values():
        for net_code, f_list in p_files.items():
            registered.setdefault(net_code, set()).update(f_list)

    if not registered:
        return {"deleted": [], "new": []}

    current_files: dict = {}
    for folder_name, net_code in (networks or {}).items():
        net_folder = os.path.join(root_folder, folder_name)
        if not os.path.isdir(net_folder):
            continue
        files = current_files.setdefault(net_code, set())
        for f in os.listdir(net_folder):
            if f.endswith((".xlsx", ".xlsm", ".xls")) and not f.lower().startswith("шаблон_"):
                files.add(f)

    deleted = []
    for net_code, f_set in registered.items():
        current_set = current_files.get(net_code, set())
        for fname in sorted(f_set):
            if fname not in current_set:
                deleted.append(f"[{net_code}] {fname}")

    new = []
    for net_code, f_set in current_files.items():
        reg_set = registered.get(net_code, set())
        for fname in sorted(f_set):
            if fname not in reg_set:
                new.append(f"[{net_code}] {fname}")

    return {"deleted": deleted, "new": new}


def prune_deleted_files(root_folder: str, networks: dict):
    """Удаляет из реестра частей (state['parts']) файлы, которых больше нет
    на диске. Вызывается операцией 1 («Обработать заявки»), чтобы снятие
    рассинхронизации (удалённые заявки) не блокировало операции 2/3 вечно —
    раньше реестр обновлялся только внутри split_by_groups (операция 2),
    а она сама блокировалась той же рассинхронизацией — замкнутый круг.
    """
    state = load_state(root_folder)
    parts_registry = state.get("parts", {})
    if not parts_registry:
        return

    current_files: dict = {}
    for folder_name, net_code in (networks or {}).items():
        net_folder = os.path.join(root_folder, folder_name)
        if not os.path.isdir(net_folder):
            continue
        files = current_files.setdefault(net_code, set())
        for f in os.listdir(net_folder):
            if f.endswith((".xlsx", ".xlsm", ".xls")) and not f.lower().startswith("шаблон_"):
                files.add(f)

    changed = False
    for p_num in list(parts_registry.keys()):
        p_files = parts_registry[p_num]
        for net_code in list(p_files.keys()):
            cur_set = current_files.get(net_code, set())
            filtered = [f for f in p_files[net_code] if f in cur_set]
            if len(filtered) != len(p_files[net_code]):
                changed = True
            if filtered:
                p_files[net_code] = filtered
            else:
                del p_files[net_code]
        if not p_files:
            del parts_registry[p_num]
            changed = True

    if not changed:
        return

    state["parts"] = parts_registry
    known_flat: dict = {}
    for p_files in parts_registry.values():
        for net_code, f_list in p_files.items():
            known_flat.setdefault(net_code, []).extend(f_list)
    for net_code in known_flat:
        known_flat[net_code] = sorted(set(known_flat[net_code]))
    state["known_files"] = known_flat
    save_state(root_folder, state)


# ── Хэши содержимого заявочников (фикс №3: замена файла с тем ─
# же именем, но новыми цифрами внутри — иначе это никак не ловится
# diff'ом заголовков колонок в template_builder) ──────────────

def get_azs_hashes(root_folder: str) -> dict:
    """Возвращает {"СЕТЬ|имя_файла.xlsx": md5}, сохранённые при прошлом
    запуске операции 1. Используется build_template, чтобы отличить
    «файл не менялся» от «файл заменили новым с тем же именем»."""
    state = load_state(root_folder)
    return state.get("azs_file_hashes", {})


def save_azs_hashes(root_folder: str, hashes: dict):
    """Сохраняет актуальные хэши заявочников в state."""
    state = load_state(root_folder)
    state["azs_file_hashes"] = hashes
    save_state(root_folder, state)


def detect_new_files(root_folder: str, current_files: dict) -> dict:
    """Сравнивает current_files со снимком из прошлого запуска."""
    known = get_known_files(root_folder)
    if not known:
        return {}
    synced_known = sync_known_files(root_folder, current_files)
    new_files = {}
    for net_code, files in current_files.items():
        known_set = set(synced_known.get(net_code, []))
        added = [f for f in files if f not in known_set]
        if added:
            new_files[net_code] = added
    return new_files