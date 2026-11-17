"""Тимчасове виправлення flet-cli 1.1.2 для Python 3.12/Windows.

У win_utils.py бінарний EXE відкривається з encoding='utf-8', що спричиняє
ValueError. Скрипт безпечно прибирає лише цей несумісний аргумент.
"""

from pathlib import Path

import flet_cli.__pyinstaller.win_utils as win_utils


path = Path(win_utils.__file__)
source = path.read_text(encoding="utf-8")
broken = 'open(exe_path, "ab", encoding="utf-8")'
fixed = 'open(exe_path, "ab")'

if broken in source:
    path.write_text(source.replace(broken, fixed), encoding="utf-8")
    print(f"Patched: {path}")
elif fixed in source:
    print(f"Already patched: {path}")
else:
    raise RuntimeError(f"Expected flet-cli pattern not found: {path}")
