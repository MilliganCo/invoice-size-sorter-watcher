# size\_sorter\_watcher

> Public experimental source: see [PUBLICATION.md](PUBLICATION.md) for actual checks, known defects, restored modules and operational limits. Earlier audit decisions are historical.

Автоматическая нормализация и сортировка столбцов размеров в файлах инвойсов .xls/.xlsx/.xlsm.

## Описание

Скрипт наблюдает за указанной папкой (по умолчанию:
`C:\example\incoming_invoices`) и при появлении нового
файла инвойса:

1. Конвертирует `.xls` в `.xlsx` через COM (с удалением исходника).
2. Находит лист «С фото» (или первый с подходящим блоком размеров).
3. Детектит строку‑шапку, где встречается несколько заголовков размеров.
4. Снимает «снимок» оригинальных столбцов размеров (значения + стили).
5. Удаляет старый блок и вставляет новые столбцы **строго** по
   `ORDERED_SIZES`, сохраняя стили и данные.
6. Сохраняет результат в файл с суффиксом `__normalized`.

## Требования

См. `requirements.txt`:

* Python 3.8+
* openpyxl
* watchdog
* pywin32 (Windows-only)

## Использование

```bash
git clone <repo>
cd invoice_size_sorter_watcher
pip install -r requirements.txt
```

Запуск:

* **Watch mode** (по умолчанию):

  ```bash
  python size_sorter_watcher.py
  ```

  Наблюдает папку `--watch-folder`, обрабатывает новые файлы.

* **One-shot**:

  ```bash
  python size_sorter_watcher.py --once
  ```

  Одноразовая обработка всех существующих инвойсов.

* **Dry-run** (только лог, без записи):

  ```bash
  python size_sorter_watcher.py --dry-run
  ```

Дополнительные опции:

* `--watch-folder <path>` — путь к папке UNC или локальной.
* `--poll-interval <secs>` — интервал для PollingObserver fallback.

## Конфигурация

Можно изменить `ORDERED_SIZES`, `SIZE_ALIASES`, задержки `DEL_DELAY`,
`POLL_INT` прямо в коде скрипта.


## Dry-run correction — 2026-10-29

`--dry-run` now returns before the conversion step. The old implementation could create an XLSX
and schedule deletion of its source XLS even in dry-run mode. The apply path retains the existing
conversion behavior. Use copies/backups before enabling actual processing.

`python -m unittest discover -s tests -v` checks the real Handler AST with I/O boundaries mocked:
dry runs never enter conversion, workbook processing or the delay; normal routing and ignored
temporary/unsupported files are also covered. Excel COM, watcher delivery and real workbook
conversion were not exercised. Tests require only the Python standard library.
