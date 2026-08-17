"""MarkdownをPowerPoint（.pptx）に変換するCLIエントリポイント。"""

from __future__ import annotations

import os
import sys
import argparse
from typing import Any

import yaml
import traceback

from config_schema import validate_config
from generator import PPTXGenerator, TemplateError

# theme.accent_color / theme.text_color を反映するフォント設定のキー
ACCENT_FONT_KEYS = ['title_h1', 'title_h2', 'title_h3', 'table_header']
TEXT_FONT_KEYS = ['body', 'bullet_level_1', 'table_body']


def apply_theme(config: dict[str, Any]) -> dict[str, Any]:
    """theme設定を各フォント設定の color_rgb に展開する"""
    theme = config.get('theme') or {}
    if not theme:
        return config

    fonts = config.setdefault('fonts', {}) or {}
    config['fonts'] = fonts

    for color, keys in ((theme.get('accent_color'), ACCENT_FONT_KEYS),
                        (theme.get('text_color'), TEXT_FONT_KEYS)):
        if not color:
            continue
        for key in keys:
            fonts.setdefault(key, {})['color_rgb'] = color

    return config


def load_config(path: str) -> dict[str, Any]:
    """YAML設定ファイルを読み込む（空ファイルの場合は空の設定を返す）"""
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def read_text_file(path: str) -> str:
    """UTF-8のテキストファイルを読み込む"""
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """コマンドライン引数を解析する"""
    parser = argparse.ArgumentParser(description="MarkdownファイルをPowerPointに変換します。")
    parser.add_argument("input", help="変換するMarkdownファイルのパス")
    parser.add_argument("-o", "--output", help="出力ファイル名", default="output.pptx")
    parser.add_argument("-c", "--config", help="YAML設定ファイルのパス", default="config.yaml")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """CLIのエントリポイント（0=成功, 1=失敗 の終了コードを返す）"""
    print("INFO: 変換処理を開始します...")
    args = parse_args(argv)

    if not os.path.exists(args.input):
        print(f"Error: 入力ファイル '{args.input}' が見つかりません。")
        return 1

    if not os.path.exists(args.config):
        print(f"Error: 設定ファイル '{args.config}' が見つかりません。")
        return 1

    try:
        config = load_config(args.config)

        # 変換を始める前に設定を検査し、原因の分かりにくいエラーを避ける
        result = validate_config(config)
        for warning in result.warnings:
            print(f"Warning: {warning}")
        if not result.is_valid:
            print(f"Error: 設定ファイル '{args.config}' に問題があります。")
            for message in result.errors:
                print(f"  - {message}")
            return 1

        config = apply_theme(config)
        content = read_text_file(args.input)

        generator = PPTXGenerator(config)
        generator.generate(content, args.output)

        print(f"Success: '{args.output}' の生成が完了しました！")
        return 0

    except TemplateError as e:
        print(f"Error: テンプレート '{args.config}' の設定を確認してください。\n  - {e}")
        return 1
    except PermissionError:
        print(f"Error: '{args.output}' に書き込めません！PowerPointでファイルを開いたままにしていませんか？閉じてから再実行してください。")
        return 1
    except Exception as e:
        print(f"Error: 予期せぬエラーが発生しました: {e}")
        traceback.print_exc()
        return 1


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
