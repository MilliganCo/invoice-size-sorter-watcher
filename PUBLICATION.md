# Public source preview — 2026-10-29

Legacy Windows Excel size-normalization watcher. Complete reachable history reviewed; no identified
credential, personal-data or actual invoice-record corpus blocker to source publication.

Fixed: --dry-run returns before conversion, delay and processing. Previously it could create an XLSX
and schedule deletion of the original XLS. The existing apply/conversion path is retained.
PASS — 3 standard-library regression tests execute the exact Handler AST with I/O mocked.
They cover dry-run exclusion, normal converted-path routing and temporary/unsupported files.
NOT RUN — Excel COM conversion, real invoice processing and watchdog observer delivery.
Actual apply mode may delete converted XLS inputs; use trusted copies/backups and read the code first.
