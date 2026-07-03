# ============================================================
# main.py — главный модуль запуска
# ============================================================
import os
import sys
import traceback

import zipfile as _zipfile

def _safe_zipfile_del(self):
    try:
        self.close()
    except Exception:
        pass

_zipfile.ZipFile.__del__ = _safe_zipfile_del

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modules.reference import load_all_references
from modules.network_detect import detect_networks
from modules.template_builder import build_template
from modules.tonnage import calculate_tonnage
from modules.group_split import split_by_groups
from modules.aux_files import create_or_update_nedovoz, create_or_update_spisok
from modules.state import find_unsynced_orders, prune_deleted_files


def choose_folder() -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        folder = filedialog.askdirectory(title="Выберите папку с заявками")
        root.destroy()
        if folder:
            return folder
    except Exception:
        pass
    print("\nДиалог недоступен. Введите путь к папке:")
    return input("Путь: ").strip().strip('"')


def menu(title: str, options: list) -> str:
    print(f"\n{'='*50}")
    print(f"  {title}")
    print(f"{'='*50}")
    for i, opt in enumerate(options, 1):
        print(f"  {i}. {opt}")
    print(f"{'='*50}")
    while True:
        try:
            choice = int(input("Выбор: ").strip())
            if 1 <= choice <= len(options):
                return options[choice - 1]
        except (ValueError, KeyboardInterrupt):
            pass
        print("Неверный ввод.")


def _get_template_path(network_folder: str, net_code: str, category: str = None) -> str:
    suffix = ""
    if category == "food":
        suffix = "_food"
    elif category == "nonfood":
        suffix = "_nonfood"
    return os.path.join(network_folder, f"Шаблон_{net_code}{suffix}.xlsx")


def _get_all_nets(networks: dict) -> list:
    """Возвращает список всех уникальных кодов сетей."""
    return sorted(set(networks.values()))


def process_orders(root_folder: str, refs: dict, networks: dict):
    """Обрабатывает заявки для всех найденных сетей.
    Недовоз/Список синхронизируются здесь же, сразу после сборки шаблонов —
    это единственное место, где они обновляются (используется и при
    одиночном запуске п.1, и при «Запустить все»), чтобы вывод не оказывался
    в другом месте программы (например, после разбивки по группам).
    """
    # Снимаем с реестра частей файлы, которых больше нет на диске —
    # иначе их «удалённость» будет блокировать операции 2/3 бесконечно,
    # ведь реестр обновляется только внутри split_by_groups (операция 2),
    # которая сама же блокируется этой рассинхронизацией.
    prune_deleted_files(root_folder, networks)

    selected_nets = _get_all_nets(networks)
    total_processed = set()

    for folder_name, net_code in networks.items():
        net_folder = os.path.join(root_folder, folder_name)

        # Проверяем наличие заявок в папке сети. Даже если заявок нет,
        # ниже всё равно вызываем build_template — он сам удалит
        # устаревший шаблон, если тот остался с прошлого раза (иначе
        # шаблон без единой заявки продолжал бы висеть на диске вечно).
        has_orders = any(
            f.endswith((".xlsx", ".xlsm", ".xls")) and not f.lower().startswith("шаблон_")
            for f in os.listdir(net_folder)
            if os.path.isfile(os.path.join(net_folder, f))
        )
        if has_orders:
            print(f"\n[{net_code}] Обновление шаблона...")
        else:
            print(f"\n[{net_code}] Заявок нет — проверяем, не остался ли старый шаблон...")

        if net_code == "МП":
            for cat in ("food", "nonfood"):
                tpl_path = _get_template_path(net_folder, net_code, cat)
                pf = build_template(
                    refs, net_code, net_folder, tpl_path,
                    category=cat, processed_files=set(),
                    root_folder=root_folder,
                )
                total_processed |= pf
        else:
            tpl_path = _get_template_path(net_folder, net_code)
            pf = build_template(
                refs, net_code, net_folder, tpl_path,
                processed_files=set(),
                root_folder=root_folder,
            )
            total_processed |= pf

    print()  # отделяем синхронизацию Недовоз/Список от вывода по последней сети
    create_or_update_nedovoz(root_folder, networks, selected_nets=selected_nets)
    create_or_update_spisok(root_folder, networks, selected_nets=selected_nets)
    print(f"\n[✓] Готово. Обработано заявок: {len(total_processed)}")


def process_tonnage(root_folder: str, networks: dict):
    """Считает тоннаж для всех найденных сетей."""
    template_paths = {}
    for folder_name, net_code in networks.items():
        net_folder = os.path.join(root_folder, folder_name)
        paths = []
        if net_code == "МП":
            for cat in ("food", "nonfood"):
                p = _get_template_path(net_folder, net_code, cat)
                if os.path.exists(p):
                    paths.append(p)
        else:
            p = _get_template_path(net_folder, net_code)
            if os.path.exists(p):
                paths.append(p)
        if paths:
            template_paths[net_code] = paths

    if not template_paths:
        print("\n[!] Шаблоны не найдены. Сначала выполните п.1 — обработку заявок.")
        return

    calculate_tonnage(root_folder, template_paths)


def process_group_split(root_folder: str, networks: dict, skip_check: bool = False):
    """Разбивает шаблоны по группам для всех найденных сетей.
    Недовоз/Список здесь больше не синхронизируются — это делает
    process_orders (операция 1), сразу после сборки шаблонов.
    """
    template_paths = {}
    for folder_name, net_code in networks.items():
        net_folder = os.path.join(root_folder, folder_name)
        paths = []
        if net_code == "МП":
            for cat in ("food", "nonfood"):
                p = _get_template_path(net_folder, net_code, cat)
                if os.path.exists(p):
                    paths.append(p)
        else:
            p = _get_template_path(net_folder, net_code)
            if os.path.exists(p):
                paths.append(p)
        if paths:
            template_paths[net_code] = paths

    if not template_paths:
        print("\n[!] Шаблоны не найдены. Сначала выполните п.1 — обработку заявок.")
        return

    result = split_by_groups(root_folder, template_paths, networks=networks,
                    skip_check=skip_check)
    if result is None:
        return  # Прервано из-за необработанных заявок — уже выведено предупреждение


def _block_if_unsynced_orders(root_folder: str, networks: dict) -> bool:
    """Проверяет рассинхронизацию заявок с диском. Блокирует операцию
    только если есть УДАЛЁННЫЕ заявки — для них на диске остаются
    «осиротевшие» части/тоннаж, и это нужно явно пересобрать через п.1.

    Новые заявки сюда не блокируются: их корректно отличает «опер.1 ещё
    не запускалась» от «опер.1 уже обработала файл» сама split_by_groups
    (по наличию столбца АЗС в шаблоне) — дублирующая грубая проверка здесь
    раньше блокировала и уже обработанные операцией 1 новые заявки
    бесконечно, т.к. в реестр частей они попадают только внутри самой
    операции 2, которая этим же блоком и не пускалась.

    Если одновременно есть и удалённые, и новые заявки, обе группы
    показываются в сообщении — раньше печатались только удалённые, и было
    непонятно, что есть ещё и необработанные новые."""
    result = find_unsynced_orders(root_folder, networks)
    deleted, new = result["deleted"], result["new"]
    if not deleted:
        return False

    print("\n[!] Заявки удалены с диска, но ещё числятся в обработанных частях:")
    for m in deleted:
        print(f"      • {m}")
    if new:
        print("\n[i] Также есть новые заявки, ещё не обработанные операцией 1:")
        for m in new:
            print(f"      • {m}")
    print("\n[!] Запустите «Запустить все операции», либо выполните операции")
    print("    по очереди начиная с п.1 «Обработать заявки и собрать шаблоны».")
    return True


def main():
    print("\n" + "="*50)
    print("  СИСТЕМА ОБРАБОТКИ ЗАЯВОК АЗС")
    print("="*50)

    root_folder = choose_folder()
    if not root_folder or not os.path.isdir(root_folder):
        print("[!] Папка не выбрана. Выход.")
        return

    print(f"\n[>] Загрузка справочников...")
    try:
        refs = load_all_references()
        print(f"    Сети: {', '.join(refs.keys())}")
    except FileNotFoundError as e:
        print(f"[!] {e}")
        return

    networks = detect_networks(root_folder, refs)
    if not networks:
        print("[!] Папки сетей не найдены в выбранной директории.")
        return
    print(f"    Найдены сети: {', '.join(networks.values())}")

    while True:
        action = menu(
            "Главное меню",
            [
                "Обработать заявки и собрать шаблоны",
                "Разделить шаблоны по группам (1С)",
                "Посчитать тоннаж",
                "Запустить все операции",
                "Выход",
            ]
        )

        if action.startswith("Выход"):
            print("\nДо свидания!")
            break

        # Меню выбора сети удалено — всегда работаем со всеми сетями
        if action.startswith("Обработать"):
            process_orders(root_folder, refs, networks)
        elif action.startswith("Разделить"):
            if _block_if_unsynced_orders(root_folder, networks):
                continue
            process_group_split(root_folder, networks)
        elif action.startswith("Посчитать"):
            if _block_if_unsynced_orders(root_folder, networks):
                continue
            process_tonnage(root_folder, networks)
        elif action.startswith("Запустить"):
            print("\n[>>] Запуск всех операций для всех сетей...")
            print("\n--- 1/3 Обработка заявок ---")
            # Недовоз/Список синхронизируются здесь же (внутри process_orders)
            process_orders(root_folder, refs, networks)
            print("\n--- 2/3 Тоннаж ---")
            process_tonnage(root_folder, networks)
            print("\n--- 3/3 Разбивка по группам (1С) ---")
            process_group_split(root_folder, networks, skip_check=True)
            print("\n[✓] Все операции завершены.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[ОШИБКА] {e}")
        traceback.print_exc()
        input("\nНажмите Enter для выхода...")