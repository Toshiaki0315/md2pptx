"""完成した資料から、書式だけのテンプレートを作るCLIエントリポイント。

    python make_template.py 資料.pptx -o テンプレート.pptx

スライドマスター・レイアウト・テーマ（配色・書体・ロゴ）はそのまま残し、
スライドだけを取り除く。md2pptx の slides.template_path に指定して使う。

完成した資料をそのままテンプレートに指定すると、その中身が出力に残ってしまう。
手作業で全スライドを消すこともできるが、消し忘れやすいためツールにしてある。
"""

from __future__ import annotations

import argparse
import os
import sys

from pptx import Presentation

#: スライドとの関連付けを表す属性（r:id）
RELATIONSHIP_ID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'


def remove_all_slides(prs) -> int:
    """プレゼンテーションからスライドを取り除き、削除した枚数を返す

    python-pptx にスライドを消すAPIが無いため、スライド一覧（sldIdLst）と
    そこからの関連付けを直接外す。マスター・レイアウト・テーマは触らない。
    """
    slide_ids = prs.slides._sldIdLst
    removed = 0
    for slide_id in list(slide_ids):
        prs.part.drop_rel(slide_id.get(RELATIONSHIP_ID))
        slide_ids.remove(slide_id)
        removed += 1
    return removed


def create_template(input_path: str, output_path: str) -> int:
    """資料からテンプレートを作り、取り除いたスライドの枚数を返す"""
    if os.path.abspath(input_path) == os.path.abspath(output_path):
        raise ValueError('元の資料を上書きしないよう、別のファイル名を指定してください。')

    prs = Presentation(input_path)
    removed = remove_all_slides(prs)
    prs.save(output_path)
    return removed


def default_output_path(input_path: str) -> str:
    """出力先の既定値（資料.pptx → 資料-template.pptx）"""
    base, ext = os.path.splitext(input_path)
    return f"{base}-template{ext or '.pptx'}"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """コマンドライン引数を解析する"""
    parser = argparse.ArgumentParser(
        description='PowerPoint資料から、書式だけのテンプレートを作ります。'
    )
    parser.add_argument('input', help='元になるPPTXファイルのパス')
    parser.add_argument('-o', '--output', help='出力するテンプレートのパス')
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """CLIのエントリポイント（0=成功, 1=失敗 の終了コードを返す）"""
    args = parse_args(argv)

    if not os.path.exists(args.input):
        print(f"Error: 入力ファイル '{args.input}' が見つかりません。")
        return 1

    output = args.output or default_output_path(args.input)
    try:
        removed = create_template(args.input, output)
    except ValueError as e:
        print(f"Error: {e}")
        return 1
    except PermissionError:
        print(f"Error: '{output}' に書き込めません！PowerPointで開いたままにしていませんか？")
        return 1
    except Exception as e:
        print(f"Error: テンプレートを作成できませんでした: {e}")
        return 1

    print(f"Success: '{output}' を作成しました（スライド{removed}枚を取り除きました）。")
    return 0


if __name__ == '__main__':  # pragma: no cover
    sys.exit(main())
