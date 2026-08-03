import pytest
import os
from unittest.mock import MagicMock, patch
from pptx.util import Inches
from io import BytesIO
from generator import PPTXGenerator
from utils import insert_image_fit

# --- フィクスチャ（テスト用の共通設定） ---
@pytest.fixture
def base_config():
    """テスト用の基本的なYAML設定のモック"""
    return {
        "slides": {
            "layout": "16:9"
        },
        "fonts": {
            "title_h1": {"name": "Meiryo", "size_pt": 44, "bold": True},
            "body": {"name": "Meiryo", "size_pt": 20},
            "inline_code": {"name": "Consolas", "color_rgb": [220, 20, 60]}
        },
        "images": {
            "default_height_inches": 3.5
        }
    }

# --- テストケース ---

def test_get_slide_size(base_config):
    """スライドサイズの計算ロジックのテスト"""
    gen = PPTXGenerator(base_config)
    w, h = gen._get_slide_size("4:3")
    assert w == Inches(10)
    assert h == Inches(7.5)
    
    # 未知のフォーマットの場合はデフォルト(16:9)が返るか
    w_def, h_def = gen._get_slide_size("unknown")
    assert w_def == Inches(10)
    assert h_def == Inches(5.625)

def test_insert_image_fit():
    """画像のリサイズとセンタリング計算のテスト"""
    mock_slide = MagicMock()
    mock_pic = MagicMock()
    # 仮の元画像サイズ (横100, 縦200)
    mock_pic.width = 100
    mock_pic.height = 200
    mock_slide.shapes.add_picture.return_value = mock_pic

    # 枠を (横500, 縦500) に設定してリサイズ実行
    pic = insert_image_fit(mock_slide, b"dummy_data", 0, 0, 500, 500)

    # ロジック内で最大1.5倍に制限(キャップ)されているため、1.5倍になる。
    assert pic.width == 150
    assert pic.height == 300

@patch('processors.insert_image_fit')
@patch('requests.get')
def test_markdown_integration(mock_get, mock_insert_image, base_config, tmp_path):
    """Markdownのパースからスライド生成までの一連の結合テスト"""
    
    # 外部API通信（requests.get）をモック化して通信をバイパス
    mock_response = MagicMock()
    mock_response.content = b"dummy_image_data"
    mock_response.status_code = 200
    mock_response.raise_for_status.return_value = None
    mock_get.return_value = mock_response

    # テスト用の網羅的なMarkdown文字列
    md_content = """
# タイトルスライド
ここはタイトルです。

## テキストと装飾のテスト
* 箇条書き1
* **太字** と `インラインコード`

> これはスピーカーノートです。

## 画像のテスト
![テスト画像](http://example.com/test.png)

## 表のテスト
| 列A | 列B |
|---|---|
| 値1 | 値2 |

## Mermaidのテスト
```mermaid
graph TD; A-->B;
```
    """

    gen = PPTXGenerator(base_config)
    output_path = tmp_path / "test_output.pptx"
    
    # 生成処理の実行
    gen.generate(md_content, str(output_path))
    
    # --- アサーション（結果の検証） ---
    
    # 1. ファイルが生成されたか
    assert os.path.exists(output_path)
    
    # 2. スライドの枚数が正しいか（h1とh2の数 = 5枚）
    assert len(gen.prs.slides) == 5
    
    # 3. ノートが正しく追加されたか（2枚目のスライド）
    notes_text = gen.prs.slides[1].notes_slide.notes_text_frame.text
    assert "これはスピーカーノートです。" in notes_text
    
    # 4. API通信と画像挿入が呼ばれたか（画像1回 + Mermaid1回 = 計2回）
    assert mock_get.call_count == 2
    assert mock_insert_image.call_count == 2

def test_main_argparse(mocker):
    """CLI引数パーサー(main関数)のテスト"""
    from md2pptx import main
    import sys
    
    # sys.argvをモックして、コマンドライン実行をシミュレート
    test_args = ["md2pptx.py", "dummy.md", "-o", "out.pptx", "-c", "config.yaml"]
    mocker.patch.object(sys, 'argv', test_args)
    
    # ファイル存在チェックをモック
    mocker.patch('os.path.exists', return_value=True)

    # === 重要：builtins.open ではなく md2pptx.open をパッチする ===
    mocker.patch('md2pptx.open', mocker.mock_open(read_data="slides:\n  layout: 16:9"))
    
    mocker.patch('yaml.safe_load', return_value={"slides": {"layout": "16:9"}, "fonts": {}, "images": {}})
    
    # 実ファイルへの書き込みを避けるためにgenerateをモック
    mock_gen_instance = MagicMock()
    mocker.patch('md2pptx.PPTXGenerator', return_value=mock_gen_instance)
    
    main()
    
    # 処理が走り、generateが呼ばれたことを確認
    mock_gen_instance.generate.assert_called_once()