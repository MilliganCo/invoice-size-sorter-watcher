"""Execute the actual legacy Handler definition with I/O boundaries mocked.

The full module imports Excel COM/watchdog and reads DONE at import. Extracting
the exact production AST avoids those side effects without duplicating _run.
"""
import ast
import os
from pathlib import Path
import re
from types import SimpleNamespace
import unittest
from unittest.mock import Mock


def load_handler():
    source = Path(__file__).parents[1] / 'size_sorter_watcher.py'
    tree = ast.parse(source.read_text(encoding='utf-8-sig'), filename=str(source))
    nodes = [node for node in tree.body if
             (isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id in {'ALLOWED_EXT', 'unc'} for t in node.targets))
             or (isinstance(node, ast.ClassDef) and node.name == 'Handler')]
    namespace = {'Path': Path, 'os': os, 're': re, 'FileSystemEventHandler': object,
                 'log': Mock(), 'time': SimpleNamespace(sleep=Mock()),
                 'convert_xls': Mock(return_value=Path('converted.xlsx')), 'process': Mock()}
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(source), 'exec'), namespace)
    return namespace


class DryRunTests(unittest.TestCase):
    def test_dry_run_never_enters_conversion_or_processing(self):
        for suffix in ('.xls', '.xlsx', '.xlsm'):
            with self.subTest(suffix=suffix):
                ns = load_handler()
                ns['Handler'](dry=True)._run('synthetic-invoice' + suffix)
                ns['convert_xls'].assert_not_called()
                ns['process'].assert_not_called()
                ns['time'].sleep.assert_not_called()

    def test_apply_path_routes_the_converted_path_to_processing(self):
        ns = load_handler()
        ns['Handler'](dry=False)._run('synthetic-invoice.xls')
        ns['convert_xls'].assert_called_once()
        ns['process'].assert_called_once_with(Path('converted.xlsx'))

    def test_unsupported_and_office_temporary_files_are_ignored(self):
        for filename in ('readme.txt', '~$synthetic.xls'):
            ns = load_handler()
            ns['Handler'](dry=False)._run(filename)
            ns['convert_xls'].assert_not_called()
            ns['process'].assert_not_called()


if __name__ == '__main__':
    unittest.main()
