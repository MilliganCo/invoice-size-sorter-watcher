# -*- coding: utf-8 -*-
"""
size_sorter_watcher.py · v2.11
------------------------------
Автоконвертация .xls → .xlsx и нормализация/сортировка столбцов размеров
в файлах инвойсов. Учитывает переносы строк, пробелы и прочие артефакты
в названиях ячеек.

Алгоритм:
1) Наблюдение за папкой (по умолчанию UNC C:\example\incoming_invoices).
2) При появлении *.xls/*.xlsx/*.xlsm:
   – конвертация .xls → .xlsx через COM (с удалением .xls через DEL_DELAY);
   – загрузка книги openpyxl;
   – выбор листа “С фото” (игнорируем пробелы/переносы) или первого,
     где detect_block найдёт ≥5 корректных заголовков размеров;
   – snapshot старого блока (значения + стили + ширины + скрытые);
   – удаление исходного блока;
   – вставка новых столбцов строго по ORDERED_SIZES с сохранением данных и стилей;
   – сохранение с суффиксом __normalized;
   – занесение в JSON-реестр, чтобы не обрабатывать повторно.
"""

import argparse, json, os, re, shutil, tempfile, time, uuid
from copy import copy
from pathlib import Path
from threading import Timer
from typing import Dict, Set, Tuple

import pythoncom, win32com.client                        # type: ignore
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from watchdog.observers import Observer, polling
from watchdog.events import FileSystemEventHandler

# ──────────────────── Константы ─────────────────────────────────────────────

ORDERED_SIZES = [
    "30*30","40*60","45*70","90*140","164-84-92","164-96-100","S-XXL",
    "XXS","XS","S","M","L","XL","XXL","XXXL","4XL","5XL","6XL","7XL",
    "40","42","44","46","48","50","52","54","56","58","60",
    "XXS-40","XS-42","S-44","M-46","L-48","XL-50","XXL-52","3XL-54",
    "4XL-56","5XL-58","6XL-60","7XL-62",
]
ALLOWED_EXT = (".xlsx", ".xlsm", ".xls")
SIZE_ALIASES = {
    "2XL": "XXL", "3XL": "XXXL",
    "XXXXL": "4XL", "XXXXXL": "5XL",
    "XXXXXXL": "6XL", "XXXXXXXL": "7XL",
}
DB_FILE    = ".size_sorter_done.json"
DEL_DELAY  = 5.0     # секунда перед удалением .xls
POLL_INT   = 1.0     # интервал для PollingObserver

# ──────────────────── Хелперы ───────────────────────────────────────────────

log = lambda m: print(f"[{time.strftime('%H:%M:%S')}] {m}", flush=True)
unc = lambda p: p if p.startswith("\\\\") else ("\\\\"+p.lstrip("\\")) if p.startswith("\\") else p

def canon(v) -> str:
    """
    Убираем все пробельные символы (включая \r,\n,\t), приводим к upper и
    заменяем алиасы. Если не в SIZE_ALIASES, возвращаем нормализованный s.
    """
    if v is None:
        return ""
    s = str(v).upper()
    # удаляем пробелы, переводы строк, табы и подчёркивания
    s = re.sub(r"[\s_]+", "", s)
    return SIZE_ALIASES.get(s, s)

def copy_style(src, dst):
    dst._style        = copy(src._style)
    dst.number_format = src.number_format

# ──────────────────── JSON “обработано” ─────────────────────────────────────

def load_done() -> Set[Tuple[str,float]]:
    try:
        return {(p, m) for p, m in json.load(open(DB_FILE, "r", encoding="utf-8"))}
    except:
        return set()

def save_done(s: Set[Tuple[str,float]]):
    try:
        json.dump(list(s), open(DB_FILE, "w", encoding="utf-8"))
    except:
        pass

DONE = load_done()

# ──────────────────── Конвертация .xls → .xlsx ───────────────────────────────

unlink_later = lambda p: Timer(DEL_DELAY, lambda: p.exists() and p.unlink()).start()

def convert_xls(src: Path) -> Path:
    if src.suffix.lower() in (".xlsx", ".xlsm"):
        return src
    dst = src.with_suffix(".xlsx")
    if dst.exists() and dst.stat().st_mtime >= src.stat().st_mtime:
        unlink_later(src)
        return dst

    log(f"Конвертация {src.name} → {dst.name}")
    tmp = Path(tempfile.gettempdir()) / f"xls2xlsx_{uuid.uuid4().hex}_{src.name}"
    shutil.copy2(src, tmp)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(
            str(tmp),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True
        )
        wb.SaveAs(str(dst), FileFormat=51)
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()
        tmp.unlink(missing_ok=True)

    unlink_later(src)
    return dst

# ──────────────────── Детекция строки-шапки размеров ────────────────────────

def detect_block(ws):
    """
    Ищем первую строку, где встретилось ≥5 распознанных размеров.
    Возвращаем (row_idx, first_size_col, last_size_col) или (None,...).
    """
    max_r = min(ws.max_row, 600)
    max_c = min(ws.max_column, 200)

    for r in range(1, max_r + 1):
        cols = [c for c in range(1, max_c + 1)
                if canon(ws.cell(r, c).value) in ORDERED_SIZES]
        if len(cols) >= 5:
            return r, min(cols), max(cols)
    return None, None, None

# ──────────────────── Перестройка блока размеров ──────────────────────────

def rebuild_sizes(ws, hr: int, first: int, last: int):
    # Снимок старого блока
    snap: Dict[str, Dict[str, object]] = {}
    for col in range(first, last + 1):
        key = canon(ws.cell(hr, col).value)
        if not key:
            continue
        values = [ws.cell(r, col).value for r in range(hr, ws.max_row + 1)]
        styles = [copy(ws.cell(r, col)._style) for r in range(hr, ws.max_row + 1)]
        nfmt   = [ws.cell(r, col).number_format     for r in range(hr, ws.max_row + 1)]
        wd     = ws.column_dimensions[get_column_letter(col)].width
        hid    = ws.column_dimensions[get_column_letter(col)].hidden
        snap[key] = {"vals": values, "styles": styles,
                     "nf": nfmt, "width": wd, "hidden": hid}

    if not snap:
        return

    # Удаляем старый диапазон
    ws.delete_cols(first, last - first + 1)

    # Вставляем заново по ORDERED_SIZES
    pos = first
    template = next(iter(snap))
    for sz in ORDERED_SIZES:
        ws.insert_cols(pos)
        letter = get_column_letter(pos)
        info = snap.get(sz, snap[template])
        ws.column_dimensions[letter].width  = info["width"]
        ws.column_dimensions[letter].hidden = info["hidden"]
        for i, st in enumerate(info["styles"]):
            r = hr + i
            cell = ws.cell(r, pos)
            cell._style = st
            cell.number_format = info["nf"][i]
            if sz in snap:
                cell.value = info["vals"][i]
        ws.cell(hr, pos).value = sz
        pos += 1

# ──────────────────── Обработка одного файла ──────────────────────────────

def process(path: Path):
    if "__normalized" in path.stem:
        return
    key = (str(path), path.stat().st_mtime)
    if key in DONE:
        return

    log(f"Обработка {path.name}")
    try:
        wb = load_workbook(path, data_only=False, keep_vba=path.suffix.lower() == ".xlsm")
    except Exception as e:
        log(f"  ! Ошибка открытия: {e}")
        return

    # Выбираем лист «С фото»
    target = None
    for sh in wb.worksheets:
        if re.sub(r"\s+", "", sh.title.lower()) == "сфото":
            target = sh
            break
    if target is None:
        # fallback — первый лист с detect_block
        for sh in wb.worksheets:
            if detect_block(sh)[0]:
                target = sh
                break

    if target is None:
        log("  ! Блок размеров не найден")
        return

    hr, first, last = detect_block(target)
    if hr is None:
        log("  ! Блок размеров не найден")
        return

    rebuild_sizes(target, hr, first, last)

    out = path.with_name(path.stem + "__normalized" + path.suffix)
    wb.save(out)
    log(f"  ✓ Сохранено {out.name}")

    DONE.add(key)
    save_done(DONE)

# ──────────────────── Watchdog Handler ─────────────────────────────────────

class Handler(FileSystemEventHandler):
    def __init__(self, dry=False):
        self.dry = dry

    def _run(self, src_path: str):
        pa = Path(unc(src_path))
        if pa.suffix.lower() not in ALLOWED_EXT or pa.name.startswith("~$"):
            return
        time.sleep(2)
        xlsx = convert_xls(pa)
        if not self.dry:
            process(xlsx)

    def on_created(self, event):
        if not event.is_directory:
            self._run(event.src_path)

    def on_moved(self, event):
        if not event.is_directory:
            self._run(event.dest_path)

def start_obs(root: Path, handler: Handler, poll_interval: float):
    try:
        obs = Observer()
        obs.schedule(handler, str(root), recursive=False)
        obs.start()
        return obs
    except FileNotFoundError:
        log("  ! PollingObserver fallback")
        obs = polling.PollingObserver(timeout=poll_interval)
        obs.schedule(handler, str(root), recursive=False)
        obs.start()
        return obs

# ──────────────────── CLI ──────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser("size-sorter v2.11")
    parser.add_argument("--watch-folder", default=r"C:\example\incoming_invoices")
    parser.add_argument("--once",       action="store_true", help="Одноразовая обработка")
    parser.add_argument("--dry-run",    action="store_true", help="Только лог, без записи")
    parser.add_argument("--poll-interval", type=float, default=POLL_INT,
                        help="Интервал для PollingObserver")
    args = parser.parse_args()

    root = Path(unc(args.watch_folder))
    if not (os.path.isdir(root) or root.exists()):
        parser.error(f"Путь недоступен: {root}")

    handler = Handler(dry=args.dry_run)

    if args.once:
        for f in root.glob("**/*"):
            if f.is_file() and f.suffix.lower() in ALLOWED_EXT and not f.name.startswith("~$"):
                handler._run(str(f))
        log("✓ Одноразовая обработка завершена")
        return

    obs = start_obs(root, handler, args.poll_interval)
    log(f"watching {root} … Ctrl+C")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        obs.stop()
        obs.join()
        log("exit")

if __name__ == "__main__":
    main()
