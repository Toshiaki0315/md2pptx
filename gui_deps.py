"""GUIの起動前に、必要なライブラリが揃っているかを確認する

素の ModuleNotFoundError では「何を、どう入れればよいか」が分かりにくいため、
起動時にまとめて確認し、導入手順を示して終了する。

third-party を import する前に実行する必要があるため、
このモジュールは**標準ライブラリだけ**に依存する。
"""

from __future__ import annotations

import importlib.util
import os
import sys
from typing import IO

#: GUIから変換を実行するために必要なモジュールと、pipで指定する名前
REQUIRED_MODULES = (
    ('yaml', 'PyYAML'),
    ('pptx', 'python-pptx'),
    ('markdown', 'Markdown'),
    ('bs4', 'beautifulsoup4'),
    ('requests', 'requests'),
    ('PIL', 'Pillow'),
    ('pygments', 'Pygments'),
)


def missing_dependencies() -> list[str]:
    """未導入のライブラリ（pipでの名前）を返す"""
    return [
        package
        for module, package in REQUIRED_MODULES
        if importlib.util.find_spec(module) is None
    ]


def dependency_message(missing: list[str], app_dir: str, executable: str) -> str:
    """導入手順の案内文を組み立てる

    実行中のPythonをそのまま示す。GUIをダブルクリックで起動した場合、
    利用者が普段使うPythonとは別のものが動いていることがあるため。
    """
    return (
        'Error: 必要なライブラリが入っていません: ' + ', '.join(missing) + '\n'
        '       次のコマンドで導入してください。\n\n'
        f'         {executable} -m pip install -r {os.path.join(app_dir, "requirements.txt")}\n\n'
        '       システムのPythonを変えたくない場合は、仮想環境を作ってください。\n\n'
        f'         cd {app_dir}\n'
        f'         {executable} -m venv .venv\n'
        '         .venv/bin/pip install -r requirements.txt\n'
        '         .venv/bin/python gui.py\n'
    )


def exit_if_missing(app_dir: str, executable: str | None = None,
                    stream: IO[str] | None = None) -> None:
    """依存ライブラリが無い場合、導入方法を示して終了する"""
    missing = missing_dependencies()
    if not missing:
        return

    print(
        dependency_message(missing, app_dir, executable or sys.executable),
        file=stream or sys.stderr,
    )
    raise SystemExit(1)
