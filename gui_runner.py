"""GUIから変換CLIを呼び出す（画面に依存しない部分）

tkinter に依存しないため、単体テストから直接実行できる。
画面への反映は呼び出し側のコールバックに任せる。
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from collections.abc import Callable
from typing import Any

import yaml

from gui_config import RGB, GuiSettings, build_config

#: このスクリプトが置かれているディレクトリ（CLIと既定の設定ファイルの場所）
APP_DIR = os.path.dirname(os.path.abspath(__file__))
CLI_PATH = os.path.join(APP_DIR, 'md2pptx.py')
DEFAULT_CONFIG_PATH = os.path.join(APP_DIR, 'config.yaml')


def rgb_to_hex(color: RGB) -> str:
    """(0, 112, 192) → '#0070c0'"""
    return '#%02x%02x%02x' % color


def hex_to_rgb(value: str) -> RGB:
    """'#0070c0' → (0, 112, 192)"""
    text = value.lstrip('#')
    return (int(text[0:2], 16), int(text[2:4], 16), int(text[4:6], 16))


def open_in_file_manager(path: str) -> None:
    """OSのファイラーで出力先を開く"""
    if sys.platform == 'darwin':
        subprocess.run(['open', '-R', path], check=False)
    elif os.name == 'nt':
        subprocess.run(['explorer', '/select,', os.path.normpath(path)], check=False)
    else:
        subprocess.run(['xdg-open', os.path.dirname(path) or '.'], check=False)


def settings_to_config(settings: GuiSettings) -> dict[str, Any]:
    """画面の設定をCLIの設定辞書にする（テンプレートのパスは絶対パスにする）

    CLIはMarkdownの場所を作業フォルダにして動かすため、相対パスのままだと
    テンプレートを見失う。
    """
    config = build_config(settings)
    if settings.use_template:
        config['slides']['template_path'] = os.path.abspath(settings.template_path)
    return config


def write_temp_config(config: dict[str, Any]) -> str:
    """設定を一時ファイルへ書き出し、そのパスを返す（呼び出し側で削除する）"""
    with tempfile.NamedTemporaryFile(
        'w', suffix='.yaml', delete=False, encoding='utf-8'
    ) as f:
        yaml.safe_dump(config, f, allow_unicode=True, sort_keys=False)
        return f.name


def build_command(input_path: str, output_path: str, config_path: str) -> list[str]:
    """CLIを呼び出すコマンドを組み立てる

    GUIを動かしているPython（sys.executable）をそのまま使う。
    仮想環境から起動された場合に、その環境のライブラリを使うため。
    """
    return [sys.executable, CLI_PATH, input_path, '-o', output_path, '-c', config_path]


def run_conversion(
    settings: GuiSettings,
    on_line: Callable[[str], None],
) -> tuple[int, str]:
    """CLIを呼び出して変換し、(終了コード, 出力ファイルのパス) を返す

    CLIの出力は1行ずつ on_line に渡す。変換の前後で例外を投げないため、
    呼び出し側は戻り値だけを見ればよい。
    """
    config_path = ''
    try:
        input_path = os.path.abspath(settings.input_path)
        output_path = os.path.abspath(settings.output_path)
        config_path = write_temp_config(settings_to_config(settings))

        # Windowsで黒いコンソール画面が一瞬開くのを防ぐ
        creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0) if os.name == 'nt' else 0
        # Markdown内の画像は相対パスで書かれているため、Markdownの場所を作業フォルダにする
        process = subprocess.Popen(
            build_command(input_path, output_path, config_path),
            cwd=os.path.dirname(input_path) or None,
            creationflags=creationflags,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='replace',
            bufsize=1,
        )
        assert process.stdout is not None
        for line in process.stdout:
            on_line(line.rstrip('\n'))
        return process.wait(), output_path
    except Exception as e:  # 予期せぬ失敗でも画面を操作不能にしない
        on_line(f'Error: 変換を実行できませんでした: {e}')
        return 1, settings.output_path
    finally:
        if config_path:
            try:
                os.unlink(config_path)
            except OSError:
                pass
