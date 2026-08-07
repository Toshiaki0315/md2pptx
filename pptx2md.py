"""PowerPoint（.pptx）を Markdown に逆変換するCLIエントリポイント。"""

from __future__ import annotations

import argparse
import os
import sys
import traceback

from pptx import Presentation

from extractor import extract, write_images

#: 画像の書き出し先の既定（出力Markdownと同じ場所からの相対パス）
DEFAULT_IMAGE_DIR = 'images'


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """コマンドライン引数を解析する"""
    parser = argparse.ArgumentParser(
        description="PowerPointファイルをMarkdownに逆変換します。"
    )
    parser.add_argument("input", help="変換するPPTXファイルのパス")
    parser.add_argument("-o", "--output", help="出力ファイル名", default="output.md")
    parser.add_argument(
        "--image-dir",
        help=f"画像の書き出し先ディレクトリ（既定: 出力先の隣の {DEFAULT_IMAGE_DIR}/）",
        default=None,
    )
    return parser.parse_args(argv)


def image_directory(args: argparse.Namespace) -> tuple[str, str]:
    """画像の書き出し先（実パス）と、Markdownに書くパスを返す"""
    if args.image_dir:
        return args.image_dir, args.image_dir

    # 既定では出力Markdownと同じ場所に images/ を作り、相対パスで参照する
    return os.path.join(os.path.dirname(args.output) or '.', DEFAULT_IMAGE_DIR), DEFAULT_IMAGE_DIR


def main(argv: list[str] | None = None) -> int:
    """CLIのエントリポイント（0=成功, 1=失敗 の終了コードを返す）"""
    print("INFO: 逆変換処理を開始します...")
    args = parse_args(argv)

    if not os.path.exists(args.input):
        print(f"Error: 入力ファイル '{args.input}' が見つかりません。")
        return 1

    try:
        output_dir, reference_dir = image_directory(args)
        result = extract(Presentation(args.input), reference_dir)

        with open(args.output, "w", encoding="utf-8") as f:
            f.write(result.markdown)
        write_images(result.images, output_dir)

        if result.images:
            print(f"INFO: 画像 {len(result.images)}枚を '{output_dir}' に書き出しました。")
        print(f"Success: '{args.output}' の生成が完了しました！")
        return 0

    except PermissionError:
        print(f"Error: '{args.output}' に書き込めません！ファイルを開いたままにしていませんか？")
        return 1
    except Exception as e:
        print(f"Error: 予期せぬエラーが発生しました: {e}")
        traceback.print_exc()
        return 1


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
