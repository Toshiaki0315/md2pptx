"""md2pptx のユニットテスト・結合テスト

モジュール構成に合わせて以下の単位でテストを分けている。
  - utils.py      : 描画・レイアウトのヘルパー関数
  - processors.py : Markdownの各タグに対応する処理
  - generator.py  : プレゼンテーション全体の組み立て
  - md2pptx.py    : CLI
"""

import base64
import os
import subprocess
from io import BytesIO
from pathlib import Path
from unittest.mock import MagicMock, patch

import markdown
import pytest
from PIL import Image
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu, Inches, Pt

import md2pptx
import mermaid_renderer
import processors
import utils
import text_metrics
from generator import PPTXGenerator
from layout import SlideLayout
from text_metrics import (
    ParagraphMetrics,
    char_width_ratio,
    estimate_height_pt,
    estimate_line_count,
    estimate_text_width_pt,
    fit_scale,
)
from mermaid_renderer import MermaidRenderError, mermaid_conf, render_mermaid
from md2pptx import apply_theme, load_config, main, parse_args, read_text_file
from processors import (
    process_blockquote,
    process_code_or_mermaid,
    process_h3,
    process_heading,
    process_hr,
    process_image,
    process_table,
    process_text,
)
from utils import (
    DEFAULT_IMAGE_DPI,
    add_runs_from_tag,
    append_text_block,
    create_code_textbox,
    apply_font_style,
    auto_shrink_text,
    downscale_image,
    hex_to_rgb,
    insert_image_fit,
    shrink_body_shape,
)

# 1x1ピクセルの最小PNG（画像挿入テスト用のダミーデータ）
TINY_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAE"
    "hQGAhKmMIQAAAABJRU5ErkJggg=="
)


# --- フィクスチャ（テスト用の共通設定） ---


@pytest.fixture
def base_config():
    """テスト用の基本的なYAML設定のモック"""
    return {
        "slides": {"layout": "16:9", "show_slide_number": False},
        "fonts": {
            "title_h1": {"name": "Meiryo", "size_pt": 44, "bold": True},
            "body": {"name": "Meiryo", "size_pt": 20},
            "inline_code": {"name": "Consolas", "color_rgb": [220, 20, 60]},
        },
        "images": {"default_height_inches": 3.5},
    }


@pytest.fixture
def gen(base_config):
    """スライド未作成のジェネレーター"""
    return PPTXGenerator(base_config)


@pytest.fixture
def gen_with_slide(gen):
    """コンテンツスライドを1枚作成済みのジェネレーター"""
    process_heading(gen, parse_md("## 見出し").find("h2"))
    return gen


@pytest.fixture
def png_file(tmp_path):
    """ローカル画像ファイルのパス"""
    path = tmp_path / "tiny.png"
    path.write_bytes(TINY_PNG)
    return str(path)


@pytest.fixture
def mock_response():
    """requests.get の戻り値のモック"""
    response = MagicMock()
    response.content = TINY_PNG
    response.status_code = 200
    response.raise_for_status.return_value = None
    return response


def parse_md(md_text):
    """Markdown文字列をBeautifulSoupに変換するヘルパー"""
    html = markdown.markdown(md_text, extensions=["extra", "fenced_code", "sane_lists"])
    return BeautifulSoup(html, "html.parser")


def parse_html(html):
    """HTML文字列をBeautifulSoupに変換するヘルパー"""
    return BeautifulSoup(html, "html.parser")


def new_paragraph():
    """スタイル検証用の実物のparagraphオブジェクトを作る"""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    return box.text_frame.paragraphs[0]


def textboxes_of(slide):
    """スライド上のテキストボックス（プレースホルダー以外）を返す"""
    return [s for s in slide.shapes if s.has_text_frame and not s.is_placeholder]


# =====================================================================
# utils.py
# =====================================================================


class TestHexToRgb:
    @pytest.mark.parametrize("value", ["#ff8000", "ff8000"])
    def test_parses_with_and_without_hash(self, value):
        """先頭の # の有無にかかわらず解釈できる"""
        assert hex_to_rgb(value) == RGBColor(255, 128, 0)

    @pytest.mark.parametrize("value", [None, ""])
    def test_empty_returns_none(self, value):
        """空の入力は None を返す"""
        assert hex_to_rgb(value) is None


class TestApplyFontStyle:
    def test_all_properties_applied(self):
        """name / size_pt / bold / color_rgb がすべて反映される"""
        run = new_paragraph().add_run()
        apply_font_style(
            run, {"name": "Meiryo", "size_pt": 24, "bold": True, "color_rgb": [1, 2, 3]}
        )
        assert run.font.name == "Meiryo"
        assert run.font.size == Pt(24)
        assert run.font.bold is True
        assert run.font.color.rgb == RGBColor(1, 2, 3)

    def test_partial_config_leaves_others_untouched(self):
        """指定されていないプロパティは変更されない"""
        run = new_paragraph().add_run()
        apply_font_style(run, {"name": "Meiryo"})
        assert run.font.name == "Meiryo"
        assert run.font.size is None
        assert run.font.bold is None

    @pytest.mark.parametrize("config", [None, {}])
    def test_empty_config_is_noop(self, config):
        """設定が空の場合は何も適用しない"""
        run = new_paragraph().add_run()
        apply_font_style(run, config)
        assert run.font.name is None


class TestDownscaleImage:
    """埋め込み画像の縮小（ファイルサイズ削減）"""

    def _png(self, tmp_path, width, height, name="big.png"):
        """指定サイズのグラデーション画像（圧縮しにくい＝縮小効果が出る）を作る"""
        img = Image.new("RGB", (width, height))
        img.putdata(
            [((x * 7) % 256, (y * 13) % 256, (x + y) % 256)
             for y in range(height) for x in range(width)]
        )
        path = tmp_path / name
        img.save(path)
        return str(path)

    def test_large_image_is_resampled(self, tmp_path):
        """表示サイズに対して過大な画像は縮小される"""
        source = self._png(tmp_path, 1600, 800)
        result = downscale_image(source, Inches(4.0), Inches(3.0), dpi=100)

        with Image.open(result) as img:
            assert img.width == 400  # 4.0インチ × 100dpi
            assert img.height == 200  # 縦横比を維持

    def test_small_image_is_untouched(self, tmp_path):
        """必要解像度以下の画像は再エンコードせず、そのまま返す"""
        source = self._png(tmp_path, 100, 50)
        assert downscale_image(source, Inches(8.0), Inches(3.8), dpi=150) is source

    def test_never_upscales(self, tmp_path):
        """元より大きな表示枠でも拡大はしない"""
        source = self._png(tmp_path, 200, 100)
        result = downscale_image(source, Inches(10.0), Inches(10.0), dpi=300)
        assert result is source

    def test_file_size_is_reduced(self, tmp_path):
        """縮小によって実際にバイト数が減る"""
        source = self._png(tmp_path, 2000, 1000)
        result = downscale_image(source, Inches(4.0), Inches(3.0), dpi=100)

        assert isinstance(result, BytesIO)
        assert result.getbuffer().nbytes < os.path.getsize(source)

    def test_dpi_controls_resolution(self, tmp_path):
        """dpiの指定が出力解像度に反映される"""
        source = self._png(tmp_path, 2000, 1000)

        with Image.open(downscale_image(source, Inches(4.0), Inches(3.0), dpi=50)) as low:
            with Image.open(downscale_image(source, Inches(4.0), Inches(3.0), dpi=200)) as high:
                assert low.width == 200
                assert high.width == 800

    def test_original_is_kept_when_resize_grows(self, tmp_path, mocker):
        """再エンコードで逆に大きくなる画像（図版など）は元データを使う"""
        source = self._png(tmp_path, 2000, 1000)
        mocker.patch("utils._source_size", return_value=1)

        assert downscale_image(source, Inches(4.0), Inches(3.0), dpi=100) is source

    def test_jpeg_format_is_preserved(self, tmp_path):
        """JPEGはJPEGのまま再エンコードされる"""
        img = Image.new("RGB", (1600, 800), (10, 120, 200))
        path = tmp_path / "big.jpg"
        img.save(path)

        with Image.open(downscale_image(str(path), Inches(4.0), Inches(3.0), dpi=100)) as out:
            assert out.format == "JPEG"

    def test_bytesio_source_is_rewound(self, tmp_path):
        """メモリ上の画像を渡した場合も、読み取り位置を先頭に戻して返す"""
        source = self._png(tmp_path, 100, 50)
        buffer = BytesIO(open(source, "rb").read())
        buffer.read()  # 読み取り位置を末尾へ

        result = downscale_image(buffer, Inches(8.0), Inches(3.8), dpi=150)
        assert result.tell() == 0

    def test_broken_image_falls_back(self, capsys):
        """画像として読めないデータは警告のみ出し、元データを返す"""
        broken = BytesIO(b"not an image")
        result = downscale_image(broken, Inches(4.0), Inches(3.0), dpi=100)

        assert result is broken
        assert "画像の縮小に失敗" in capsys.readouterr().out

    def test_bytesio_large_image_is_downscaled(self, tmp_path):
        """メモリ上の画像（ダウンロードした画像・Mermaid図）も縮小対象になる"""
        source = self._png(tmp_path, 2000, 1000)
        buffer = BytesIO(open(source, "rb").read())

        result = downscale_image(buffer, Inches(4.0), Inches(3.0), dpi=100)

        assert result is not buffer
        with Image.open(result) as img:
            assert img.width == 400

    def test_size_of_unmeasurable_source_is_none(self):
        """バイト数を取得できないデータでも例外にしない"""
        broken = MagicMock()
        broken.tell.side_effect = OSError("測定不能")
        assert utils._source_size(broken) is None


class TestImageDpiConfig:
    """images.downscale / images.dpi の設定"""

    def test_default_dpi(self, gen):
        """既定では縮小が有効"""
        assert processors.image_dpi(gen) == DEFAULT_IMAGE_DPI

    def test_dpi_from_config(self, base_config):
        """dpi を設定で変更できる"""
        base_config["images"]["dpi"] = 96
        assert processors.image_dpi(PPTXGenerator(base_config)) == 96

    def test_downscale_can_be_disabled(self, base_config):
        """downscale: false で縮小を無効化できる"""
        base_config["images"]["downscale"] = False
        assert processors.image_dpi(PPTXGenerator(base_config)) is None

    @patch("processors.insert_image_fit")
    @patch("processors.downscale_image")
    def test_place_image_downscales(self, mock_downscale, _mock_fit, gen_with_slide, png_file):
        """縮小が有効な場合は表示枠のサイズで縮小してから配置する"""
        processors.place_image(
            gen_with_slide, png_file, Inches(1.0), Inches(1.5), Inches(8.0), Inches(3.8)
        )
        mock_downscale.assert_called_once_with(
            png_file, Inches(8.0), Inches(3.8), DEFAULT_IMAGE_DPI
        )

    @patch("processors.insert_image_fit")
    @patch("processors.downscale_image")
    def test_place_image_skips_when_disabled(
        self, mock_downscale, mock_fit, base_config, png_file
    ):
        """縮小が無効な場合は元画像のまま配置する"""
        base_config["images"]["downscale"] = False
        gen = PPTXGenerator(base_config)
        process_heading(gen, parse_md("## 見出し").find("h2"))

        processors.place_image(
            gen, png_file, Inches(1.0), Inches(1.5), Inches(8.0), Inches(3.8)
        )
        mock_downscale.assert_not_called()
        assert mock_fit.call_args[0][1] is png_file


class TestInsertImageFit:
    def _mock_slide(self, width, height):
        mock_slide = MagicMock()
        mock_pic = MagicMock()
        mock_pic.width, mock_pic.height = width, height
        mock_slide.shapes.add_picture.return_value = mock_pic
        return mock_slide

    def test_small_image_is_scaled_up_to_cap(self):
        """小さい画像は最大1.5倍までしか拡大されない"""
        pic = insert_image_fit(self._mock_slide(100, 200), b"dummy", 0, 0, 500, 500)
        assert (pic.width, pic.height) == (150, 300)

    def test_large_image_is_shrunk_to_fit(self):
        """大きい画像はアスペクト比を保って枠内に縮小される"""
        pic = insert_image_fit(self._mock_slide(1000, 500), b"dummy", 0, 0, 500, 500)
        assert (pic.width, pic.height) == (500, 250)

    def test_image_is_centered_in_frame(self):
        """縮小後の画像は枠の中央に配置される"""
        pic = insert_image_fit(self._mock_slide(1000, 500), b"dummy", 100, 200, 500, 500)
        assert pic.left == 100
        assert pic.top == 200 + (500 - 250) / 2


class TestAddRunsFromTag:
    def _runs_for(self, gen, html):
        p = new_paragraph()
        add_runs_from_tag(
            parse_html(html).p, p, {"name": "Meiryo"}, gen.fonts_conf.get("inline_code")
        )
        return p.runs

    def test_bold_and_italic(self, gen):
        """strong/em がフォントスタイルに変換される"""
        runs = self._runs_for(gen, "<p><strong>太字</strong><em>斜体</em></p>")
        assert [r.text for r in runs] == ["太字", "斜体"]
        assert runs[0].font.bold is True
        assert runs[1].font.italic is True

    def test_inline_code_uses_code_font(self, gen):
        """codeタグにはinline_codeの設定が適用される"""
        (run,) = self._runs_for(gen, "<p><code>x = 1</code></p>")
        assert run.font.name == "Consolas"
        assert run.font.color.rgb == RGBColor(220, 20, 60)

    def test_inline_code_falls_back_to_defaults(self, base_config):
        """inline_code が未設定でも既定の等幅フォント・色になる"""
        base_config["fonts"].pop("inline_code")
        gen = PPTXGenerator(base_config)
        (run,) = self._runs_for(gen, "<p><code>x</code></p>")
        assert run.font.name == "Consolas"
        assert run.font.color.rgb == RGBColor(220, 20, 60)

    def test_plain_text_uses_default_font(self, gen):
        """装飾の無いテキストには既定フォントが適用される"""
        (run,) = self._runs_for(gen, "<p>ふつうの文字</p>")
        assert run.font.name == "Meiryo"

    def test_nested_container_is_flattened(self, gen):
        """spanなどのコンテナ要素は再帰的に展開される"""
        runs = self._runs_for(gen, "<p><span>外<strong>中</strong></span></p>")
        assert [r.text for r in runs] == ["外", "中"]
        assert runs[1].font.bold is True

    def test_block_tags_are_skipped(self, gen):
        """リストや表などのブロック要素はインライン処理から除外される"""
        runs = self._runs_for(gen, "<p>本文<ul><li>項目</li></ul><table></table></p>")
        assert [r.text for r in runs] == ["本文"]

    def test_newlines_are_replaced(self, gen):
        """改行は半角スペースに置換される"""
        runs = self._runs_for(gen, "<p>1行目\n2行目</p>")
        assert runs[0].text == "1行目 2行目"


class TestUtilsArePure:
    """utils はジェネレーターに依存しない（引数だけで完結する）"""

    def test_no_generator_import(self):
        """utils モジュールが generator を参照していない"""
        source = (Path(utils.__file__)).read_text(encoding="utf-8")
        assert "generator" not in source

    def test_add_runs_works_without_generator(self):
        """ジェネレーターを作らずにインライン装飾を描画できる"""
        p = new_paragraph()
        add_runs_from_tag(
            parse_html("<p>本文<code>x</code></p>").p,
            p,
            {"name": "Meiryo"},
            {"name": "Courier", "color_rgb": [1, 2, 3]},
        )

        assert [r.text for r in p.runs] == ["本文", "x"]
        assert p.runs[1].font.name == "Courier"
        assert p.runs[1].font.color.rgb == RGBColor(1, 2, 3)

    def test_inline_code_defaults_without_config(self):
        """インラインコードの設定を渡さなくても既定値が使われる"""
        p = new_paragraph()
        add_runs_from_tag(parse_html("<p><code>x</code></p>").p, p, None)

        assert p.runs[0].font.name == "Consolas"
        assert p.runs[0].font.color.rgb == RGBColor(220, 20, 60)

    def test_append_text_block_takes_a_text_frame(self):
        """本文枠を直接渡して段落を追加できる"""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        text_frame = slide.shapes.add_textbox(
            Inches(1), Inches(1), Inches(4), Inches(2)
        ).text_frame

        append_text_block(text_frame, parse_md("本文").find("p"), reuse_first_paragraph=True)
        assert text_frame.paragraphs[0].runs[0].text == "本文"

    def test_create_code_textbox_takes_a_slide(self):
        """スライドと寸法を渡すだけでコード枠を作れる"""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        create_code_textbox(
            slide, Inches(1), Inches(1), Inches(4), Inches(2),
            content="print(1)", language="python",
            font_conf={"name": "Consolas", "size_pt": 12}, background_rgb=[1, 2, 3],
        )

        box = textboxes_of(slide)[0]
        assert box.fill.fore_color.rgb == RGBColor(1, 2, 3)
        assert "print" in box.text_frame.text


class TestShrinkBodyShape:
    def test_width_is_applied(self, gen_with_slide):
        """指定した幅が本文枠に反映される"""
        shrink_body_shape(gen_with_slide.current_slide, Inches(4.8))
        assert gen_with_slide.current_slide.placeholders[1].width == Inches(4.8)

    def test_inherited_geometry_is_kept(self, gen_with_slide):
        """幅だけを変えてもレイアウトから継承した位置・高さが失われない

        プレースホルダーの寸法は継承値のため、一部だけ書き換えると残りが0になる。
        shrink_body_shape 内の left/top 自己代入が効いていることの回帰テスト。
        """
        shape = gen_with_slide.current_slide.placeholders[1]
        top, height = shape.top, shape.height

        shrink_body_shape(gen_with_slide.current_slide, Inches(4.8))

        assert shape.left > 0
        assert (shape.top, shape.height) == (top, height)

    def test_max_height_is_applied(self, gen_with_slide):
        """max_height を渡すと高さも縮む"""
        shrink_body_shape(gen_with_slide.current_slide, Inches(8.0), max_height=Inches(2.0))
        assert gen_with_slide.current_slide.placeholders[1].height == Inches(2.0)

    def test_without_slide_is_noop(self, gen):
        """スライド未作成でも例外にならない"""
        shrink_body_shape(gen.current_slide, Inches(4.8))


class TestAppendTextBlock:
    def test_first_paragraph_is_reused_then_appended(self, gen_with_slide):
        """最初の段落は空の既存段落を再利用し、以降は追加される"""
        append_text_block(
            gen_with_slide.current_body, parse_md("1つ目").find("p"),
            reuse_first_paragraph=True,
        )
        assert len(gen_with_slide.current_body.paragraphs) == 1

        gen_with_slide.slide_has_text = True
        append_text_block(
            gen_with_slide.current_body, parse_md("2つ目").find("p"),
            reuse_first_paragraph=False,
        )
        assert len(gen_with_slide.current_body.paragraphs) == 2

    def test_level_and_spacing(self, gen_with_slide):
        """レベルと行間・段落後余白が設定される"""
        append_text_block(
            gen_with_slide.current_body, parse_md("* 項目").find("li"),
            reuse_first_paragraph=True, level=2,
        )
        p = gen_with_slide.current_body.paragraphs[0]
        assert p.level == 2
        assert p.space_after == Pt(12)
        assert p.line_spacing == 1.2


class TestAppendCodeTextbox:
    def test_creates_textbox_with_background(self, gen_with_slide):
        """コードは背景色付きの独立したテキストボックスに描画される"""
        processors.append_code_textbox(gen_with_slide, "print(1)", language="python")

        boxes = textboxes_of(gen_with_slide.current_slide)
        assert len(boxes) == 1
        assert boxes[0].fill.fore_color.rgb == RGBColor(40, 44, 52)
        assert gen_with_slide.slide_has_text is True

    def test_background_color_from_theme(self, base_config):
        """theme.code_bg_color で背景色を変更できる"""
        base_config["theme"] = {"code_bg_color": [1, 2, 3]}
        gen = PPTXGenerator(base_config)
        process_heading(gen, parse_md("## 見出し").find("h2"))

        processors.append_code_textbox(gen, "print(1)", language="python")
        assert textboxes_of(gen.current_slide)[0].fill.fore_color.rgb == RGBColor(1, 2, 3)

    def test_syntax_highlight_splits_runs(self, gen_with_slide):
        """シンタックスハイライトによりトークンごとにrunが分割される"""
        processors.append_code_textbox(gen_with_slide, "def f():\n    return 1\n", language="python")

        runs = textboxes_of(gen_with_slide.current_slide)[0].text_frame.paragraphs[0].runs
        assert len(runs) > 1
        assert "".join(r.text for r in runs) == "def f():\n    return 1\n"
        assert all(r.font.color.rgb is not None for r in runs)

    def test_unknown_language_falls_back_to_plain(self, gen_with_slide):
        """未知の言語指定でも例外にせずプレーンテキストとして描画する"""
        processors.append_code_textbox(gen_with_slide, "hello", language="no_such_language")
        runs = textboxes_of(gen_with_slide.current_slide)[0].text_frame.paragraphs[0].runs
        # pygmentsは末尾に改行を補うため、内容の一致のみを確認する
        assert "".join(r.text for r in runs).strip() == "hello"

    def test_language_is_guessed_when_omitted(self, gen_with_slide):
        """言語未指定でも推定してハイライトする"""
        processors.append_code_textbox(gen_with_slide, "def f():\n    return 1\n", language=None)
        runs = textboxes_of(gen_with_slide.current_slide)[0].text_frame.paragraphs[0].runs
        assert "".join(r.text for r in runs) == "def f():\n    return 1\n"

    def test_two_column_layout_shrinks_body(self, gen_with_slide):
        """テキストがある場合は本文枠を縮めて右側に配置する"""
        gen_with_slide.slide_has_text = True
        processors.append_code_textbox(gen_with_slide, "x", language="python")

        assert gen_with_slide.current_slide.placeholders[1].width == Inches(4.5)
        assert textboxes_of(gen_with_slide.current_slide)[0].left == Inches(5.0)

    def test_center_layout(self, gen_with_slide):
        """forced_layout=center では中央寄せの枠になる"""
        gen_with_slide.forced_layout = "center"
        processors.append_code_textbox(gen_with_slide, "x", language="python")
        assert textboxes_of(gen_with_slide.current_slide)[0].left == Inches(1.5)

    def test_default_layout(self, gen_with_slide):
        """テキストが無い場合はスライド幅いっぱいに配置する"""
        processors.append_code_textbox(gen_with_slide, "x", language="python")
        box = textboxes_of(gen_with_slide.current_slide)[0]
        assert (box.left, box.width) == (Inches(1.0), Inches(8.0))


class TestAutoShrinkText:
    def _fill_body(self, gen, line_count, text=None):
        for i in range(line_count):
            gen.slide_has_text = i > 0
            body = parse_md(text or f"行{i}").find("p")
            append_text_block(
                gen.current_body, body,
                reuse_first_paragraph=not gen.slide_has_text, font_conf={"size_pt": 20},
            )

    def _sizes(self, gen):
        return [
            run.font.size.pt for p in gen.current_body.paragraphs for run in p.runs
            if run.font.size
        ]

    def test_long_single_paragraph_is_shrunk(self, gen_with_slide):
        """折り返しの多い長い1段落も縮小される

        段落数だけを数えていた頃は「1段落」と判定され、
        実際には10行以上に折り返してはみ出していても縮小されなかった。
        """
        self._fill_body(gen_with_slide, 1, text="これは非常に長い一段落です。" * 30)
        auto_shrink_text(gen_with_slide.current_slide)

        assert all(size < 20 for size in self._sizes(gen_with_slide))

    def test_short_single_paragraph_is_untouched(self, gen_with_slide):
        """短い1段落は縮小しない"""
        self._fill_body(gen_with_slide, 1, text="短い本文です。")
        auto_shrink_text(gen_with_slide.current_slide)

        assert all(size == 20 for size in self._sizes(gen_with_slide))

    def test_fullwidth_text_shrinks_more_than_halfwidth(self, gen_with_slide, gen):
        """同じ文字数なら、全角のほうが幅を取るため強く縮小される"""
        self._fill_body(gen_with_slide, 1, text="あ" * 300)
        auto_shrink_text(gen_with_slide.current_slide)
        fullwidth = min(self._sizes(gen_with_slide))

        process_heading(gen, parse_md("## 見出し").find("h2"))
        self._fill_body(gen, 1, text="a" * 300)
        auto_shrink_text(gen.current_slide)
        halfwidth = min(self._sizes(gen))

        assert fullwidth < halfwidth

    def test_shrinks_when_too_many_lines(self, gen_with_slide):
        """行数が枠に収まらない場合はフォントサイズが縮小される"""
        self._fill_body(gen_with_slide, 12)
        auto_shrink_text(gen_with_slide.current_slide)

        sizes = [
            run.font.size for p in gen_with_slide.current_body.paragraphs for run in p.runs
        ]
        assert sizes and all(size < Pt(20) for size in sizes)

    def test_keeps_size_when_few_lines(self, gen_with_slide):
        """行数が少ない場合は縮小しない"""
        self._fill_body(gen_with_slide, 3)
        auto_shrink_text(gen_with_slide.current_slide)

        sizes = [
            run.font.size for p in gen_with_slide.current_body.paragraphs for run in p.runs
        ]
        assert sizes and all(size == Pt(20) for size in sizes)

    def test_heading_and_bullets_fit_in_the_frame(self, base_config, tmp_path):
        """h3見出しを含む本文がスライドからはみ出さない

        sample.md の「現状の分析と課題」で実際にはみ出していたケースの回帰テスト。
        h3 の space_before と、フォントの行高を考慮していなかったことが原因だった。
        """
        base_config["fonts"]["bullet_level_1"] = {"size_pt": 18}
        base_config["fonts"]["title_h3"] = {"size_pt": 20, "bold": True}
        gen = PPTXGenerator(base_config)

        md = (
            "## 現状の分析と課題\n\n現場では多くの課題が山積しています。\n\n"
            "### システムの老朽化\n* 既存システムの動作が遅い\n* メンテナンス担当者が不在\n\n"
            "### コミュニケーションの課題\n* 部署間での情報共有にタイムラグがある\n"
            "* 手作業によるデータ入力の負荷が高い\n"
        )
        gen.generate(md, str(tmp_path / "out.pptx"))

        body = gen.prs.slides[0].placeholders[1]
        tf = body.text_frame
        needed = estimate_height_pt(
            [utils._paragraph_metrics(p) for p in tf.paragraphs],
            Emu(int(body.width - tf.margin_left - tf.margin_right)).pt,
        )
        assert needed <= Emu(int(body.height - tf.margin_top - tf.margin_bottom)).pt

    def test_paragraph_level_font_size_is_also_shrunk(self, gen_with_slide):
        """段落レベルのフォントサイズ（h3が設定する基準値）も縮小対象にする"""
        self._fill_body(gen_with_slide, 1, text="あ" * 400)
        target = gen_with_slide.current_body.paragraphs[0]
        target.font.size = Pt(20)

        auto_shrink_text(gen_with_slide.current_slide)
        assert target.font.size.pt < 20

    def test_space_before_is_shrunk(self, gen_with_slide):
        """段落前の余白も縮小される"""
        self._fill_body(gen_with_slide, 1, text="あ" * 400)
        target = gen_with_slide.current_body.paragraphs[0]
        target.space_before = Pt(10)

        auto_shrink_text(gen_with_slide.current_slide)
        assert target.space_before.pt < 10

    def test_none_slide_is_noop(self):
        """スライドが None でも例外にならない"""
        auto_shrink_text(None)

    def test_unexpected_slide_structure_is_ignored(self):
        """想定外の構造のスライドを渡しても例外を投げない（縮小は best-effort）"""
        broken = MagicMock()
        broken.placeholders.__len__.side_effect = ValueError("broken")
        auto_shrink_text(broken)

    def test_slide_without_body_is_noop(self, gen):
        """本文プレースホルダーが無いレイアウトでも例外にならない"""
        blank = gen.prs.slides.add_slide(gen.prs.slide_layouts[6])
        auto_shrink_text(blank)


# =====================================================================
# processors.py
# =====================================================================


class TestProcessHeading:
    def test_h1_creates_title_slide(self, gen):
        """h1はタイトルスライド（レイアウト0）を作る"""
        process_heading(gen, parse_md("# タイトル").find("h1"))
        assert len(gen.prs.slides) == 1
        assert gen.current_slide.shapes.title.text == "タイトル"
        assert gen.current_slide.slide_layout == gen.prs.slide_layouts[0]

    def test_h2_creates_content_slide(self, gen):
        """h2はコンテンツスライド（レイアウト1）を作る"""
        process_heading(gen, parse_md("## 中身").find("h2"))
        assert gen.current_slide.slide_layout == gen.prs.slide_layouts[1]

    def test_title_font_is_applied(self, gen):
        """title_h1 の設定がタイトルに適用される"""
        process_heading(gen, parse_md("# タイトル").find("h1"))
        run = gen.current_slide.shapes.title.text_frame.paragraphs[0].runs[0]
        assert run.font.size == Pt(44)
        assert run.font.bold is True

    def test_generic_title_config_is_fallback(self, base_config):
        """title_h2 が無い場合は title の設定にフォールバックする"""
        base_config["fonts"]["title"] = {"size_pt": 18}
        gen = PPTXGenerator(base_config)
        process_heading(gen, parse_md("## 中身").find("h2"))
        run = gen.current_slide.shapes.title.text_frame.paragraphs[0].runs[0]
        assert run.font.size == Pt(18)

    def test_state_is_reset_for_new_slide(self, gen):
        """スライド作成時に本文とテキスト有無フラグが初期化される"""
        gen.slide_has_text = True
        process_heading(gen, parse_md("## 中身").find("h2"))
        assert gen.slide_has_text is False
        assert gen.current_body.text == ""

    def test_body_height_is_clamped_to_slide(self, gen):
        """デフォルトテンプレートでは本文枠がスライド内に収まるよう補正される"""
        process_heading(gen, parse_md("## 中身").find("h2"))
        shape = gen.current_slide.placeholders[1]
        assert shape.top + shape.height == gen.prs.slide_height - Inches(0.5)
        assert shape.width > 0  # 継承値のリセット防止

    def test_previous_slide_is_shrunk(self, gen):
        """新しいスライドを作る前に直前のスライドの自動縮小が走る"""
        process_heading(gen, parse_md("## 1枚目").find("h2"))
        with patch("processors.auto_shrink_text") as mock_shrink:
            process_heading(gen, parse_md("## 2枚目").find("h2"))
        mock_shrink.assert_called_once()

    def test_body_correction_failure_is_ignored(self, gen, mocker):
        """本文枠の補正に失敗してもスライド生成は継続する"""
        mocker.patch.object(SlideLayout, "body_height_for", side_effect=ValueError("broken"))

        process_heading(gen, parse_md("## 中身").find("h2"))
        assert gen.current_slide.shapes.title.text == "中身"


class TestProcessH3:
    BU_NONE = "{http://schemas.openxmlformats.org/drawingml/2006/main}buNone"

    def test_adds_subheading_paragraph(self, gen_with_slide):
        """h3は本文内の小見出し段落として追加される"""
        process_h3(gen_with_slide, parse_md("### 小見出し").find("h3"))

        assert gen_with_slide.current_body.paragraphs[-1].runs[0].text == "小見出し"
        assert gen_with_slide.slide_has_text is True

    def test_bullet_is_disabled(self, gen_with_slide):
        """箇条書き記号が無効化される（buNone要素が入る）"""
        process_h3(gen_with_slide, parse_md("### 小見出し").find("h3"))

        p_pr = gen_with_slide.current_body.paragraphs[-1]._element.get_or_add_pPr()
        assert p_pr.find(self.BU_NONE) is not None

    def test_font_config_is_applied(self, base_config):
        """title_h3 の設定が適用される"""
        base_config["fonts"]["title_h3"] = {"name": "Meiryo", "size_pt": 26, "bold": True}
        gen = PPTXGenerator(base_config)
        process_heading(gen, parse_md("## 見出し").find("h2"))

        process_h3(gen, parse_md("### 小見出し").find("h3"))
        p = gen.current_body.paragraphs[-1]
        assert p.runs[0].font.size == Pt(26)
        assert p.font.size == Pt(26)

    def test_empty_heading_is_skipped(self, gen_with_slide):
        """空のh3は段落を追加しない"""
        before = len(gen_with_slide.current_body.paragraphs)
        process_h3(gen_with_slide, parse_html("<h3>  </h3>").h3)
        assert len(gen_with_slide.current_body.paragraphs) == before


class TestProcessHr:
    def test_creates_untitled_slide(self, gen_with_slide):
        """水平線はタイトルの無い新しいスライドを作る"""
        process_hr(gen_with_slide, parse_md("---").find("hr"))

        assert len(gen_with_slide.prs.slides) == 2
        assert gen_with_slide.current_slide.shapes.title is None
        assert gen_with_slide.slide_has_text is False

    def test_body_uses_full_height(self, gen_with_slide):
        """タイトルが無い分、本文枠を上部から広く使う"""
        process_hr(gen_with_slide, parse_md("---").find("hr"))

        shape = gen_with_slide.current_slide.placeholders[1]
        assert shape.top == Inches(0.5)
        assert shape.width > 0

    def test_body_correction_failure_is_ignored(self, gen_with_slide, mocker):
        """本文枠の補正に失敗してもスライド生成は継続する"""
        mocker.patch.object(SlideLayout, "body_height_for", side_effect=ValueError("broken"))

        process_hr(gen_with_slide, parse_md("---").find("hr"))
        assert len(gen_with_slide.prs.slides) == 2


class TestProcessBlockquote:
    def test_note_is_written(self, gen_with_slide):
        """引用ブロックはスピーカーノートに書き込まれる"""
        process_blockquote(gen_with_slide, parse_md("> メモ").find("blockquote"))
        assert gen_with_slide.current_slide.notes_slide.notes_text_frame.text == "メモ"

    def test_multiple_notes_are_appended(self, gen_with_slide):
        """複数の引用ブロックは空行を挟んで追記される"""
        process_blockquote(gen_with_slide, parse_md("> 1つ目").find("blockquote"))
        process_blockquote(gen_with_slide, parse_md("> 2つ目").find("blockquote"))

        notes = gen_with_slide.current_slide.notes_slide.notes_text_frame.text
        assert notes == "1つ目\n\n2つ目"


class TestProcessImage:
    def test_local_path_is_inserted(self, gen_with_slide, png_file):
        """ローカル画像はダウンロードせずにそのまま挿入される"""
        process_image(gen_with_slide, parse_html(f'<img src="{png_file}">').img)
        assert len(gen_with_slide.current_slide.shapes) == 3  # title, body, picture

    @patch("requests.get")
    def test_remote_image_is_downloaded_with_timeout(
        self, mock_get, gen_with_slide, mock_response
    ):
        """URL画像はタイムアウト付きで取得される"""
        mock_get.return_value = mock_response
        process_image(gen_with_slide, parse_html('<img src="http://example.com/a.png">').img)

        mock_get.assert_called_once_with(
            "http://example.com/a.png", timeout=processors.HTTP_TIMEOUT_SEC
        )
        mock_response.raise_for_status.assert_called_once()

    @patch("processors.insert_image_fit")
    def test_fixed_position_from_config(self, mock_fit, base_config, png_file):
        """position_inches が指定されている場合はオートレイアウトを使わない"""
        base_config["images"]["position_inches"] = [5.2, 1.8]
        gen = PPTXGenerator(base_config)
        process_heading(gen, parse_md("## 見出し").find("h2"))
        gen.current_slide = MagicMock()

        process_image(gen, parse_html(f'<img src="{png_file}">').img)

        mock_fit.assert_not_called()
        gen.current_slide.shapes.add_picture.assert_called_once_with(
            png_file, Inches(5.2), Inches(1.8), height=Inches(3.5)
        )

    @patch("processors.insert_image_fit")
    def test_two_column_when_slide_has_text(self, mock_fit, gen_with_slide, png_file):
        """テキストがあるスライドでは本文枠を縮めて右側に配置する"""
        gen_with_slide.slide_has_text = True
        process_image(gen_with_slide, parse_html(f'<img src="{png_file}">').img)

        args = mock_fit.call_args[0]
        assert (args[2], args[4]) == (Inches(5.2), Inches(4.5))
        assert gen_with_slide.current_slide.placeholders[1].width == Inches(4.8)

    @patch("processors.insert_image_fit")
    def test_forced_two_column_layout(self, mock_fit, gen_with_slide, png_file):
        """forced_layout=2-column ではテキストが無くても右側に配置する"""
        gen_with_slide.forced_layout = "2-column"
        process_image(gen_with_slide, parse_html(f'<img src="{png_file}">').img)
        assert mock_fit.call_args[0][2] == Inches(5.2)

    @patch("processors.insert_image_fit")
    def test_forced_center_layout(self, mock_fit, gen_with_slide, png_file):
        """forced_layout=center ではテキストの有無によらず中央に配置する"""
        gen_with_slide.slide_has_text = True
        gen_with_slide.forced_layout = "center"
        process_image(gen_with_slide, parse_html(f'<img src="{png_file}">').img)

        assert (mock_fit.call_args[0][2], mock_fit.call_args[0][4]) == (
            Inches(1.0),
            Inches(8.0),
        )

    def test_missing_src_is_skipped(self, gen_with_slide, capsys):
        """src属性が無い画像は警告を出してスキップする"""
        process_image(gen_with_slide, parse_html("<img>").img)

        assert "Warning" in capsys.readouterr().out
        assert len(gen_with_slide.current_slide.shapes) == 2

    @patch("requests.get", side_effect=OSError("network down"))
    def test_download_failure_is_reported(self, _mock_get, gen_with_slide, capsys):
        """取得に失敗しても処理は継続し、警告のみ表示する"""
        process_image(gen_with_slide, parse_html('<img src="https://example.com/a.png">').img)
        assert "画像の挿入に失敗しました" in capsys.readouterr().out


class TestProcessTable:
    MD_TABLE = "| 列A | 列B |\n|---|---|\n| 値1 | 値2 |"

    def _table_of(self, gen):
        return [s for s in gen.current_slide.shapes if s.has_table][0].table

    def test_rows_and_cells_are_converted(self, gen_with_slide):
        """行数・列数とセルのテキストが再現される"""
        process_table(gen_with_slide, parse_md(self.MD_TABLE).find("table"))
        table = self._table_of(gen_with_slide)

        assert (len(table.rows), len(table.columns)) == (2, 2)
        assert table.cell(0, 0).text == "列A"
        assert table.cell(1, 1).text == "値2"
        assert gen_with_slide.slide_has_text is True

    def test_header_style(self, gen_with_slide):
        """ヘッダー行は太字・専用サイズ・背景色が付く"""
        process_table(gen_with_slide, parse_md(self.MD_TABLE).find("table"))
        table = self._table_of(gen_with_slide)

        header = table.cell(0, 0)
        assert header.text_frame.paragraphs[0].runs[0].font.bold is True
        assert header.text_frame.paragraphs[0].runs[0].font.size == Pt(14)
        assert header.fill.fore_color.rgb == RGBColor(31, 73, 125)
        assert table.cell(1, 0).text_frame.paragraphs[0].runs[0].font.size == Pt(12)

    def test_inline_decoration_in_cell(self, gen_with_slide):
        """セル内の太字・インラインコードも反映される"""
        process_table(gen_with_slide, parse_md("| 見出し |\n|---|\n| **強調** |").find("table"))
        run = self._table_of(gen_with_slide).cell(1, 0).text_frame.paragraphs[0].runs[0]
        assert run.font.bold is True

    def test_column_count_uses_widest_row(self, gen_with_slide):
        """列数が揃っていない場合は最大列数に合わせる"""
        html = "<table><tr><td>a</td></tr><tr><td>b</td><td>c</td></tr></table>"
        process_table(gen_with_slide, parse_html(html).table)
        assert len(self._table_of(gen_with_slide).columns) == 2

    def test_empty_table_is_skipped(self, gen_with_slide):
        """行が無い表は何も生成しない"""
        process_table(gen_with_slide, parse_html("<table></table>").table)
        assert len(gen_with_slide.current_slide.shapes) == 2

    def test_layout_is_split_when_slide_has_text(self, gen_with_slide):
        """テキストがある場合は本文枠を縮めて表を下半分に置く"""
        gen_with_slide.slide_has_text = True
        process_table(gen_with_slide, parse_md(self.MD_TABLE).find("table"))

        shape = [s for s in gen_with_slide.current_slide.shapes if s.has_table][0]
        assert shape.top == Inches(2.8)
        assert gen_with_slide.current_slide.placeholders[1].height == Inches(2.0)

    def test_layout_is_full_when_slide_is_empty(self, gen_with_slide):
        """テキストが無い場合は表を上部から配置する"""
        process_table(gen_with_slide, parse_md(self.MD_TABLE).find("table"))
        shape = [s for s in gen_with_slide.current_slide.shapes if s.has_table][0]
        assert shape.top == Inches(1.5)


class TestProcessCodeOrMermaid:
    def test_code_block_without_language_does_not_crash(self, gen_with_slide):
        """言語指定の無いコードブロックでも落ちない（class属性がNoneになるケース）

        `code_tag.get('class')` は class 属性が無いと None を返すため、
        そのまま `in` 演算子に渡すと TypeError になる回帰テスト。
        """
        pre = parse_md("```\nplain text\n```").find("pre")
        assert pre.find("code").get("class") is None  # 前提条件の確認

        process_code_or_mermaid(gen_with_slide, pre)

        runs = textboxes_of(gen_with_slide.current_slide)[0].text_frame.paragraphs[0].runs
        assert "plain text" in "".join(r.text for r in runs)

    def test_language_is_extracted(self, gen_with_slide):
        """language-xxx クラスから言語名が取り出される"""
        with patch("processors.append_code_textbox") as mock_append:
            process_code_or_mermaid(gen_with_slide, parse_md("```python\nx=1\n```").find("pre"))
        assert mock_append.call_args[1]["language"] == "python"

    def test_code_block_is_rendered(self, gen_with_slide):
        """通常のコードブロックは背景色付きテキストボックスになる"""
        process_code_or_mermaid(gen_with_slide, parse_md("```python\nprint(1)\n```").find("pre"))

        assert len(textboxes_of(gen_with_slide.current_slide)) == 1
        assert gen_with_slide.slide_has_text is True

    @patch("processors.insert_image_fit")
    @patch("requests.get")
    def test_mermaid_is_rendered_as_image(
        self, mock_get, mock_fit, gen_with_slide, mock_response
    ):
        """Mermaid記法は画像として挿入される"""
        mock_get.return_value = mock_response
        md = "```mermaid\ngraph TD; A-->B;\n```"
        process_code_or_mermaid(gen_with_slide, parse_md(md).find("pre"))

        assert mock_get.call_args[0][0].startswith(mermaid_renderer.DEFAULT_KROKI_ENDPOINT)
        assert mock_get.call_args[1]["timeout"] == mermaid_renderer.HTTP_TIMEOUT_SEC
        mock_fit.assert_called_once()

    @patch("processors.insert_image_fit")
    @patch("requests.get")
    def test_mermaid_two_column_layout(self, mock_get, mock_fit, gen_with_slide, mock_response):
        """テキストがあるスライドではMermaid図を右側に配置する"""
        mock_get.return_value = mock_response
        gen_with_slide.slide_has_text = True
        process_code_or_mermaid(
            gen_with_slide, parse_md("```mermaid\ngraph TD; A-->B;\n```").find("pre")
        )
        assert mock_fit.call_args[0][2] == Inches(5.2)

    @patch("processors.insert_image_fit")
    @patch("requests.get")
    def test_falls_back_to_mermaid_ink(
        self, mock_get, mock_fit, gen_with_slide, mock_response, capsys
    ):
        """Krokiが失敗した場合はmermaid.inkにフォールバックする"""
        mock_get.side_effect = [OSError("kroki down"), mock_response]
        process_code_or_mermaid(
            gen_with_slide, parse_md("```mermaid\ngraph TD; A-->B;\n```").find("pre")
        )

        assert mock_get.call_count == 2
        assert mock_get.call_args[0][0].startswith(mermaid_renderer.MERMAID_INK_URL)
        assert "代替API" in capsys.readouterr().out
        mock_fit.assert_called_once()

    @patch("processors.place_image")
    def test_renderer_off_places_nothing(self, mock_place, base_config, capsys):
        """mermaid.renderer が off の場合、図を配置せずスライドはそのまま進む"""
        base_config["mermaid"] = {"renderer": "off"}
        gen = PPTXGenerator(base_config)
        process_heading(gen, parse_md("## 見出し").find("h2"))

        process_code_or_mermaid(gen, parse_md("```mermaid\ngraph TD; A-->B;\n```").find("pre"))

        mock_place.assert_not_called()
        assert "スキップ" in capsys.readouterr().out

    @patch("requests.get", side_effect=OSError("timeout"))
    def test_both_apis_failing_is_reported(self, _mock_get, gen_with_slide, capsys):
        """両方のAPIが失敗しても警告のみ表示して処理を継続する"""
        process_code_or_mermaid(
            gen_with_slide, parse_md("```mermaid\ngraph TD; A-->B;\n```").find("pre")
        )
        assert "Mermaid図形の生成に失敗しました" in capsys.readouterr().out


class TestProcessText:
    def test_paragraph_uses_body_font(self, gen_with_slide):
        """段落にはbodyのフォント設定が適用される"""
        process_text(gen_with_slide, parse_md("本文です").find("p"))

        run = gen_with_slide.current_body.paragraphs[0].runs[0]
        assert run.text == "本文です"
        assert run.font.size == Pt(20)
        assert gen_with_slide.slide_has_text is True

    def test_empty_text_is_skipped(self, gen_with_slide):
        """空の要素は段落を追加しない"""
        process_text(gen_with_slide, parse_html("<p>   </p>").p)

        assert gen_with_slide.slide_has_text is False
        assert len(gen_with_slide.current_body.paragraphs) == 1

    def test_bullet_level_follows_nesting(self, gen_with_slide):
        """ネストの深さがリストのレベルになる"""
        for item in parse_md("* 親\n    * 子\n        * 孫").find_all("li"):
            process_text(gen_with_slide, item)

        assert [p.level for p in gen_with_slide.current_body.paragraphs] == [0, 1, 2]

    def test_bullet_level_is_capped(self, gen_with_slide):
        """レベルはPowerPointの上限(8)でクランプされる"""
        html = "<ul><li>深い</li></ul>"
        for _ in range(12):
            html = f"<ul><li>{html}</li></ul>"

        process_text(gen_with_slide, parse_html(html).find_all("li")[-1])
        assert gen_with_slide.current_body.paragraphs[0].level == 8

    def test_bullet_level_font_config(self, base_config):
        """bullet_level_N の設定がレベルごとに適用される"""
        base_config["fonts"]["bullet_level_1"] = {"size_pt": 18}
        gen = PPTXGenerator(base_config)
        process_heading(gen, parse_md("## 見出し").find("h2"))

        process_text(gen, parse_md("* 項目").find("li"))
        assert gen.current_body.paragraphs[0].runs[0].font.size == Pt(18)


# =====================================================================
# text_metrics.py
# =====================================================================


class TestCharWidth:
    @pytest.mark.parametrize("ch", ["あ", "漢", "全", "ー", "、"])
    def test_fullwidth_characters(self, ch):
        """日本語は全角幅として扱う"""
        assert char_width_ratio(ch) == 1.0

    @pytest.mark.parametrize("ch", ["a", "Z", "1", " ", "-"])
    def test_halfwidth_characters(self, ch):
        """英数字・記号は半角幅として扱う"""
        assert char_width_ratio(ch) == 0.5

    def test_width_of_mixed_text(self):
        """全角と半角が混在してもそれぞれの幅で合算される"""
        # 全角2文字(=2.0) + 半角4文字(=2.0) → 合計4.0文字ぶん
        assert estimate_text_width_pt("あいabcd", 10) == pytest.approx(40.0)


class TestEstimateLineCount:
    def test_text_wraps_by_available_width(self):
        """枠幅を超えるテキストは折り返して複数行になる"""
        # 全角20文字 × 10pt = 200pt を 100pt 幅に入れる → 2行
        assert estimate_line_count("あ" * 20, 10, 100) == 2

    def test_short_text_is_single_line(self):
        assert estimate_line_count("短い", 10, 500) == 1

    @pytest.mark.parametrize("text, width", [("", 100), ("あ", 0), ("あ", -10)])
    def test_degenerate_cases_are_one_line(self, text, width):
        """空文字や幅が取れない場合も1行として扱う（0除算を避ける）"""
        assert estimate_line_count(text, 10, width) == 1


class TestEstimateHeight:
    def _para(self, text, size=10, level=0, spacing=1.0, space_after=0.0, space_before=0.0):
        return ParagraphMetrics(text, size, level, spacing, space_after, space_before)

    def test_single_line_height(self):
        """1行の高さは フォントサイズ × フォントの行高 × 行送り"""
        # 10pt × 1.3(行高) × 1.2(行送り) = 15.6pt
        assert estimate_height_pt([self._para("あ", spacing=1.2)], 500) == pytest.approx(15.6)

    def test_line_height_exceeds_font_size(self):
        """行の高さはフォントサイズそのものではなく、フォントの行高を考慮する

        ここを 1.0 とみなしていた頃は必要な高さを約3割過小評価し、
        本文がスライドからはみ出していた。
        """
        height = estimate_height_pt([self._para("あ", size=10, spacing=1.0)], 500)
        assert height > 10
        assert height == pytest.approx(10 * text_metrics.LINE_HEIGHT_RATIO)

    def test_wrapped_lines_are_counted(self):
        """折り返した行数ぶんの高さになる"""
        # 全角30文字 × 10pt = 300pt を 100pt 幅 → 3行
        assert estimate_height_pt([self._para("あ" * 30)], 100) == pytest.approx(39.0)

    def test_space_after_is_included(self):
        assert estimate_height_pt([self._para("あ", space_after=5.0)], 500) == pytest.approx(18.0)

    def test_space_before_is_included(self):
        """段落前の余白（h3見出しなどが使う）も高さに含める"""
        assert estimate_height_pt([self._para("あ", space_before=10.0)], 500) == pytest.approx(23.0)

    def test_indent_reduces_available_width(self):
        """箇条書きのレベルが深いほど幅が狭まり、行数が増える"""
        flat = estimate_height_pt([self._para("あ" * 30, level=0)], 300)
        nested = estimate_height_pt([self._para("あ" * 30, level=3)], 300)
        assert nested > flat

    def test_scale_reduces_height(self):
        """縮小率を掛けると必要な高さが下がる"""
        paragraphs = [self._para("あ" * 30)]
        assert estimate_height_pt(paragraphs, 100, scale=0.5) < estimate_height_pt(
            paragraphs, 100
        )


class TestFitScale:
    def _paras(self, count, text="あ" * 20, size=20):
        return [ParagraphMetrics(text, size) for _ in range(count)]

    def test_no_shrink_when_it_fits(self):
        """収まっている場合は縮小しない"""
        assert fit_scale(self._paras(1), 500, 1000) == 1.0

    def test_shrinks_when_overflowing(self):
        """収まらない場合は1.0未満の縮小率を返す"""
        scale = fit_scale(self._paras(20), 500, 200)
        assert scale < 1.0

    def test_result_actually_fits(self):
        """返した縮小率で実際に枠へ収まる"""
        paragraphs = self._paras(10)
        scale = fit_scale(paragraphs, 500, 300)
        assert estimate_height_pt(paragraphs, 500, scale) <= 300

    def test_scale_has_a_floor(self):
        """極端に多い場合でも下限より小さくはしない（読めなくなるため）"""
        assert fit_scale(self._paras(500), 500, 50) == text_metrics.MIN_SHRINK_SCALE

    @pytest.mark.parametrize("paragraphs, height", [([], 100), ([ParagraphMetrics("あ", 10)], 0)])
    def test_degenerate_cases_do_not_shrink(self, paragraphs, height):
        """段落が無い・高さが取れない場合は縮小しない"""
        assert fit_scale(paragraphs, 500, height) == 1.0


# =====================================================================
# layout.py
# =====================================================================


class TestSlideLayout:
    """スライドサイズからの配置寸法の導出"""

    #: 16:9 における従来の固定値（この値を再現できることが移行の前提）
    LEGACY_16_9 = {
        "content_left": 1.0,
        "content_top": 1.5,
        "content_width": 8.0,
        "content_height": 3.8,
        "split_body_width": 4.8,
        "code_split_body_width": 4.5,
        "split_image_left": 5.2,
        "split_image_width": 4.5,
        "table_split_top": 2.8,
        "table_split_body_height": 2.0,
        "code_split_left": 5.0,
        "code_split_width": 4.5,
        "code_center_left": 1.5,
        "code_center_width": 7.0,
        "code_full_top": 2.0,
        "code_full_height": 3.0,
    }

    @pytest.mark.parametrize("name, expected", sorted(LEGACY_16_9.items()))
    def test_16_9_reproduces_legacy_values(self, name, expected):
        """16:9では従来の固定値と一致する（既存資料の見た目を変えないための回帰テスト）"""
        layout = SlideLayout(Inches(10), Inches(5.625))
        assert Emu(getattr(layout, name)).inches == pytest.approx(expected, abs=0.001)

    @pytest.mark.parametrize(
        "width_in, height_in",
        [(10, 5.625), (10, 7.5), (10, 6.25), (11.69, 8.27)],  # 16:9 / 4:3 / 16:10 / A4
    )
    def test_content_area_follows_slide_size(self, width_in, height_in):
        """コンテンツ領域はスライドサイズに追従し、左右・下の余白は一定になる"""
        layout = SlideLayout(Inches(width_in), Inches(height_in))

        assert Emu(layout.content_width).inches == pytest.approx(width_in - 2.0, abs=0.001)
        assert Emu(layout.content_height).inches == pytest.approx(
            height_in - 1.5 - 0.325, abs=0.001
        )

    @pytest.mark.parametrize("width_in", [10, 11.69, 13.333])
    def test_split_image_reaches_right_margin(self, width_in):
        """2カラム時の図は、スライド幅によらず右端の余白まで広がる"""
        layout = SlideLayout(Inches(width_in), Inches(7.5))
        right_edge = Emu(layout.split_image_left + layout.split_image_width).inches
        assert right_edge == pytest.approx(width_in - 0.3, abs=0.001)

    def test_body_height_fits_in_slide(self):
        """本文枠の高さはスライド下端に余白を残す"""
        layout = SlideLayout(Inches(10), Inches(7.5))
        assert Emu(layout.body_height_for(Inches(1.75))).inches == pytest.approx(5.25)

    def test_built_from_presentation(self, base_config):
        """ジェネレーターはプレゼンテーションの実サイズからレイアウトを構築する"""
        base_config["slides"]["layout"] = "A4"
        gen = PPTXGenerator(base_config)

        assert gen.layout.width == gen.prs.slide_width
        assert gen.layout.height == gen.prs.slide_height


class TestLayoutAdaptsToSlideSize:
    """画角を変えたときに図・表が追従することの結合テスト"""

    MD = "## 図のスライド\n\n![img]({png})\n\n## 表のスライド\n\n| A | B |\n|---|---|\n| 1 | 2 |\n"

    def _generate(self, base_config, layout_name, tmp_path, png_file):
        base_config["slides"]["layout"] = layout_name
        gen = PPTXGenerator(base_config)
        gen.generate(self.MD.format(png=png_file), str(tmp_path / "out.pptx"))
        return gen

    @pytest.mark.parametrize("layout_name", ["16:9", "4:3", "A4"])
    def test_content_stays_inside_the_slide(self, base_config, layout_name, tmp_path, png_file):
        """どの画角でも図・表がスライドからはみ出さない"""
        gen = self._generate(base_config, layout_name, tmp_path, png_file)

        for slide in gen.prs.slides:
            for shape in slide.shapes:
                assert shape.left + shape.width <= gen.prs.slide_width
                assert shape.top + shape.height <= gen.prs.slide_height

    def test_wide_slide_uses_the_extra_width(self, base_config, tmp_path, png_file):
        """A4のような横長の画角では、表が広がった幅を使う"""
        narrow = self._generate(base_config, "16:9", tmp_path, png_file)
        wide = self._generate(base_config, "A4", tmp_path, png_file)

        def table_width(gen):
            for slide in gen.prs.slides:
                for shape in slide.shapes:
                    if shape.has_table:
                        return shape.width
            raise AssertionError("表が見つかりません")

        assert table_width(wide) > table_width(narrow)


# =====================================================================
# mermaid_renderer.py
# =====================================================================


class TestMermaidConfig:
    def test_conf_is_read_from_config(self, base_config):
        """config.yaml の mermaid セクションを読み取る"""
        base_config["mermaid"] = {"renderer": "off"}
        assert mermaid_conf(PPTXGenerator(base_config)) == {"renderer": "off"}

    def test_missing_section_is_empty(self, gen):
        """mermaid セクションが無い場合は空の設定になる"""
        assert mermaid_conf(gen) == {}

    @pytest.mark.parametrize(
        "endpoint, expected",
        [
            ("https://kroki.io", True),
            ("https://mermaid.ink/img/", True),
            ("http://kroki.internal.example.com", False),
            ("http://localhost:8000", False),
        ],
    )
    def test_public_endpoint_detection(self, endpoint, expected):
        """公開サービスかどうかを判定できる（警告とフォールバック可否の基準）"""
        assert mermaid_renderer.is_public_endpoint(endpoint) is expected


class TestRenderMermaid:
    MMD = "graph TD; A-->B;"

    def test_off_skips_generation(self, capsys):
        """renderer: off では図を生成せず None を返す"""
        assert render_mermaid({"renderer": "off"}, self.MMD) is None
        assert "スキップ" in capsys.readouterr().out

    def test_unknown_renderer_raises(self):
        """未知のレンダラー名はエラーにする（設定ミスを黙って無視しない）"""
        with pytest.raises(MermaidRenderError, match="未知の mermaid.renderer"):
            render_mermaid({"renderer": "magic"}, self.MMD)

    @patch("requests.get")
    def test_default_uses_public_kroki(self, mock_get, mock_response):
        """既定では公開Krokiを使う（従来どおりの動作）"""
        mock_get.return_value = mock_response
        assert render_mermaid({}, self.MMD) == TINY_PNG

        url = mock_get.call_args[0][0]
        assert url.startswith(f"{mermaid_renderer.DEFAULT_KROKI_ENDPOINT}/mermaid/png/")
        assert mock_get.call_args[1]["timeout"] == mermaid_renderer.HTTP_TIMEOUT_SEC

    @patch("requests.get")
    def test_warns_before_sending_to_public_service(self, mock_get, mock_response, capsys):
        """公開サービスへ送信する前に警告を表示する"""
        mock_get.return_value = mock_response
        render_mermaid({}, self.MMD)

        out = capsys.readouterr().out
        assert "外部サービス" in out
        assert "kroki.io" in out

    @patch("requests.get")
    def test_warning_can_be_disabled(self, mock_get, mock_response, capsys):
        """警告は warn_on_external: false で抑制できる"""
        mock_get.return_value = mock_response
        render_mermaid({"warn_on_external": False}, self.MMD)
        assert "外部サービス" not in capsys.readouterr().out

    @patch("requests.get")
    def test_self_hosted_endpoint_is_used(self, mock_get, mock_response, capsys):
        """自己ホストのKrokiを指定でき、その場合は警告を出さない"""
        mock_get.return_value = mock_response
        render_mermaid({"endpoint": "http://kroki.internal/"}, self.MMD)

        assert mock_get.call_args[0][0].startswith("http://kroki.internal/mermaid/png/")
        assert "外部サービス" not in capsys.readouterr().out

    @patch("requests.get")
    def test_public_endpoint_falls_back_to_mermaid_ink(self, mock_get, mock_response):
        """公開Kroki利用時は従来どおり mermaid.ink にフォールバックする"""
        mock_get.side_effect = [OSError("kroki down"), mock_response]

        assert render_mermaid({}, self.MMD) == TINY_PNG
        assert mock_get.call_args[0][0].startswith(mermaid_renderer.MERMAID_INK_URL)

    @patch("requests.get")
    def test_self_hosted_never_falls_back_to_public(self, mock_get):
        """自己ホスト指定時は公開APIへフォールバックしない（情報漏洩の防止）

        社内Krokiが落ちた際に、機密を含み得る図が公開サービスへ送られるのを防ぐ。
        """
        mock_get.side_effect = OSError("社内Krokiが停止")

        with pytest.raises(MermaidRenderError, match="情報漏洩"):
            render_mermaid({"endpoint": "http://kroki.internal"}, self.MMD)

        assert mock_get.call_count == 1  # mermaid.ink は呼ばれない

    @patch("requests.get")
    def test_fallback_can_be_forced_for_self_hosted(self, mock_get, mock_response):
        """明示的に許可した場合のみ、自己ホストでもフォールバックする"""
        mock_get.side_effect = [OSError("down"), mock_response]

        result = render_mermaid(
            {"endpoint": "http://kroki.internal", "fallback_to_public": True}, self.MMD
        )
        assert result == TINY_PNG
        assert mock_get.call_count == 2

    @patch("requests.get")
    def test_fallback_can_be_disabled_for_public(self, mock_get):
        """公開Kroki利用時でもフォールバックを禁止できる"""
        mock_get.side_effect = OSError("down")

        with pytest.raises(MermaidRenderError):
            render_mermaid({"fallback_to_public": False}, self.MMD)
        assert mock_get.call_count == 1


class TestLocalMermaidRendering:
    """mermaid-cli によるオフライン生成（外部送信なし）"""

    MMD = "graph TD; A-->B;"
    CONF = {"renderer": "local"}

    def _fake_run(self, png_bytes=TINY_PNG, returncode=0):
        """mmdc の代わりに出力ファイルを書くダミー"""
        def run(command, **kwargs):
            if returncode == 0:
                output_path = command[command.index("-o") + 1]
                with open(output_path, "wb") as f:
                    f.write(png_bytes)
            stderr = "パースに失敗しました".encode("utf-8")
            return subprocess.CompletedProcess(command, returncode, b"", stderr)
        return run

    def test_renders_without_network(self, mocker):
        """ローカル生成ではネットワークを一切使わない"""
        mocker.patch("subprocess.run", side_effect=self._fake_run())
        mock_get = mocker.patch("requests.get")

        assert render_mermaid(self.CONF, self.MMD) == TINY_PNG
        mock_get.assert_not_called()

    def test_diagram_source_is_passed_to_cli(self, mocker):
        """図のソースを一時ファイル経由でCLIへ渡す"""
        captured = {}

        def run(command, **kwargs):
            source_path = command[command.index("-i") + 1]
            with open(source_path, encoding="utf-8") as f:
                captured["source"] = f.read()
            with open(command[command.index("-o") + 1], "wb") as f:
                f.write(TINY_PNG)
            return subprocess.CompletedProcess(command, 0, b"", b"")

        mocker.patch("subprocess.run", side_effect=run)
        render_mermaid(self.CONF, self.MMD)
        assert captured["source"] == self.MMD

    def test_custom_cli_path(self, mocker):
        """cli_path で実行コマンドを差し替えられる"""
        mock_run = mocker.patch("subprocess.run", side_effect=self._fake_run())
        render_mermaid({"renderer": "local", "cli_path": "/opt/bin/mmdc"}, self.MMD)
        assert mock_run.call_args[0][0][0] == "/opt/bin/mmdc"

    def test_missing_cli_is_reported(self, mocker):
        """mermaid-cli が未インストールの場合は導入方法を案内する"""
        mocker.patch("subprocess.run", side_effect=FileNotFoundError)

        with pytest.raises(MermaidRenderError, match="mermaid-cli"):
            render_mermaid(self.CONF, self.MMD)

    def test_cli_failure_is_reported(self, mocker):
        """CLIが異常終了した場合は標準エラー出力を添えて報告する"""
        mocker.patch("subprocess.run", side_effect=self._fake_run(returncode=1))

        with pytest.raises(MermaidRenderError, match="失敗しました"):
            render_mermaid(self.CONF, self.MMD)

    def test_timeout_is_reported(self, mocker):
        """タイムアウトした場合も分かるメッセージにする"""
        mocker.patch(
            "subprocess.run", side_effect=subprocess.TimeoutExpired("mmdc", 60)
        )

        with pytest.raises(MermaidRenderError, match="タイムアウト"):
            render_mermaid(self.CONF, self.MMD)

    def test_missing_output_is_reported(self, mocker):
        """CLIが正常終了しても画像が無い場合はエラーにする"""
        mocker.patch(
            "subprocess.run",
            return_value=subprocess.CompletedProcess(["mmdc"], 0, b"", b""),
        )

        with pytest.raises(MermaidRenderError, match="画像を出力しませんでした"):
            render_mermaid(self.CONF, self.MMD)


# =====================================================================
# generator.py
# =====================================================================


class TestInitialization:
    @pytest.mark.parametrize(
        "layout, expected",
        [
            ("16:9", (Inches(10), Inches(5.625))),
            ("4:3", (Inches(10), Inches(7.5))),
            ("16:10", (Inches(10), Inches(6.25))),
            ("A4", (Inches(11.69), Inches(8.27))),
            ("unknown", (Inches(10), Inches(5.625))),
        ],
    )
    def test_get_slide_size(self, gen, layout, expected):
        """スライドサイズの計算ロジック（未知の値は16:9にフォールバック）"""
        assert gen._get_slide_size(layout) == expected

    def test_slide_size_applied_from_config(self, base_config):
        """config.yamlの画角がプレゼンテーションに反映される"""
        base_config["slides"]["layout"] = "4:3"
        gen = PPTXGenerator(base_config)
        assert (gen.prs.slide_width, gen.prs.slide_height) == (Inches(10), Inches(7.5))

    @pytest.mark.parametrize(
        "config", [None, {}, {"slides": None, "fonts": None, "images": None}]
    )
    def test_empty_config_is_accepted(self, config):
        """設定が空・Noneでも既定値で初期化できる"""
        assert PPTXGenerator(config).prs.slide_height == Inches(5.625)

    def test_template_is_loaded_when_exists(self, base_config, tmp_path):
        """テンプレートが存在する場合は読み込み、画角設定を上書きしない"""
        template = tmp_path / "template.pptx"
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        prs.save(str(template))

        base_config["slides"]["template_path"] = str(template)
        assert PPTXGenerator(base_config).prs.slide_width == Inches(13.333)

    def test_missing_template_falls_back_to_default(self, base_config):
        """テンプレートのパスが不正な場合はデフォルトのプレゼンテーションを使う"""
        base_config["slides"]["template_path"] = "no_such_template.pptx"
        assert PPTXGenerator(base_config).prs.slide_height == Inches(5.625)


class TestFrontMatter:
    FM = (
        '---\ntitle: "資料タイトル"\nsubtitle: "サブ"\n'
        'author: "著者"\ndate: "2026-05-10"\n---\n\n## 中身\n本文\n'
    )

    def test_title_slide_is_generated(self, gen, tmp_path):
        """フロントマターからタイトルスライドが自動生成される"""
        gen.generate(self.FM, str(tmp_path / "out.pptx"))

        title_slide = gen.prs.slides[0]
        assert title_slide.shapes.title.text == "資料タイトル"
        assert title_slide.placeholders[1].text == "サブ\n著者\n2026-05-10"
        assert len(gen.prs.slides) == 2

    def test_title_font_is_applied(self, gen, tmp_path):
        """タイトルには title_h1 の設定が適用される"""
        gen.generate(self.FM, str(tmp_path / "out.pptx"))
        run = gen.prs.slides[0].shapes.title.text_frame.paragraphs[0].runs[0]
        assert run.font.size == Pt(44)

    def test_broken_front_matter_is_reported(self, gen, tmp_path, capsys):
        """壊れたフロントマターは警告を出して処理を続行する"""
        gen.generate("---\ntitle: [壊れた\n---\n\n## 中身\n", str(tmp_path / "out.pptx"))
        assert "フロントマターの解析に失敗しました" in capsys.readouterr().out

    def test_front_matter_without_title_is_ignored(self, gen, tmp_path):
        """titleが無いフロントマターではタイトルスライドを作らない"""
        gen.generate("---\nauthor: 著者\n---\n\n## 中身\n", str(tmp_path / "out.pptx"))
        assert len(gen.prs.slides) == 1


class TestLayoutComments:
    @patch("processors.insert_image_fit")
    def test_two_column_comment(self, mock_fit, gen, tmp_path, png_file):
        """<!-- layout: 2-column --> で画像が右側に寄る"""
        md = f"## 見出し\n\n<!-- layout: 2-column -->\n\n![img]({png_file})\n"
        gen.generate(md, str(tmp_path / "out.pptx"))

        assert gen.forced_layout == "2-column"
        assert mock_fit.call_args[0][2] == Inches(5.2)

    @patch("processors.insert_image_fit")
    def test_center_comment(self, mock_fit, gen, tmp_path, png_file):
        """<!-- layout: center --> で画像が中央に配置される"""
        md = f"## 見出し\n\n本文\n\n<!-- layout: center -->\n\n![img]({png_file})\n"
        gen.generate(md, str(tmp_path / "out.pptx"))
        assert mock_fit.call_args[0][2] == Inches(1.0)

    def test_layout_is_reset_on_new_slide(self, gen, tmp_path):
        """新しい見出しでレイアウト指定はリセットされる"""
        gen.generate("## 1枚目\n\n<!-- layout: center -->\n\n## 2枚目\n", str(tmp_path / "out.pptx"))
        assert gen.forced_layout is None


class TestSlideNumbers:
    def test_numbers_are_added_except_first_slide(self, base_config, tmp_path):
        """先頭スライドを除いてページ番号が挿入される"""
        base_config["slides"]["show_slide_number"] = True
        gen = PPTXGenerator(base_config)
        gen.generate("# 表紙\n\n## 2枚目\n\n## 3枚目\n", str(tmp_path / "out.pptx"))

        assert textboxes_of(gen.prs.slides[0]) == []
        assert textboxes_of(gen.prs.slides[1])[0].text_frame.text == "1"
        assert textboxes_of(gen.prs.slides[2])[0].text_frame.text == "2"

    def test_numbers_can_be_disabled(self, base_config, tmp_path):
        """show_slide_number: false で無効化できる"""
        base_config["slides"]["show_slide_number"] = False
        gen = PPTXGenerator(base_config)
        gen.generate("# 表紙\n\n## 2枚目\n", str(tmp_path / "out.pptx"))
        assert textboxes_of(gen.prs.slides[1]) == []

    def test_enabled_by_default(self, tmp_path):
        """設定が無い場合はページ番号を表示する"""
        gen = PPTXGenerator({"slides": {"layout": "16:9"}})
        gen.generate("# 表紙\n\n## 2枚目\n", str(tmp_path / "out.pptx"))
        assert textboxes_of(gen.prs.slides[1])[0].text_frame.text == "1"


class TestGenerate:
    @patch("processors.insert_image_fit")
    @patch("requests.get")
    def test_markdown_integration(
        self, mock_get, mock_insert_image, base_config, mock_response, tmp_path
    ):
        """Markdownのパースからスライド生成までの一連の結合テスト"""
        mock_get.return_value = mock_response

        md_content = """
# タイトルスライド
ここはタイトルです。

## テキストと装飾のテスト
### 小見出し
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

## コードのテスト
```python
print(1)
```
"""
        gen = PPTXGenerator(base_config)
        output_path = tmp_path / "test_output.pptx"
        gen.generate(md_content, str(output_path))

        # ファイルが生成され、h1/h2の数だけスライドが作られる
        assert os.path.exists(output_path)
        assert len(gen.prs.slides) == 6

        # ノートが正しいスライドに追加される
        notes = gen.prs.slides[1].notes_slide.notes_text_frame.text
        assert "これはスピーカーノートです。" in notes

        # 外部通信と画像挿入は 画像1回 + Mermaid1回
        assert mock_get.call_count == 2
        assert mock_insert_image.call_count == 2

    def test_content_before_first_heading_is_ignored(self, gen, tmp_path):
        """見出しより前の本文は配置先が無いため無視される"""
        gen.generate("本文だけ\n\n# タイトル\n中身", str(tmp_path / "out.pptx"))

        assert len(gen.prs.slides) == 1
        assert gen.current_body.text == "中身"

    def test_list_items_are_converted(self, gen, tmp_path):
        """箇条書きは1項目につき1段落になる"""
        gen.generate("# タイトル\n\n* 項目1\n* 項目2\n", str(tmp_path / "out.pptx"))
        assert [p.text for p in gen.current_body.paragraphs] == ["項目1", "項目2"]

    def test_list_item_paragraph_is_not_duplicated(self, gen, tmp_path):
        """段落を含むリスト（loose list）でもpタグが二重に出力されない"""
        gen.generate("# タイトル\n\n* 項目1\n\n* 項目2\n", str(tmp_path / "out.pptx"))
        assert [p.text.strip() for p in gen.current_body.paragraphs] == ["項目1", "項目2"]

    def test_blockquote_paragraph_is_not_duplicated(self, gen, tmp_path):
        """引用内のpタグは本文に出力されない"""
        gen.generate("# タイトル\n\n> ノート\n", str(tmp_path / "out.pptx"))

        assert gen.current_body.text == ""
        assert gen.current_slide.notes_slide.notes_text_frame.text == "ノート"

    def test_hr_creates_additional_slide(self, gen, tmp_path):
        """水平線でスライドが追加される"""
        gen.generate("## 1枚目\n\n本文\n\n---\n\n続き\n", str(tmp_path / "out.pptx"))
        assert len(gen.prs.slides) == 2


# =====================================================================
# md2pptx.py（CLI）
# =====================================================================


class TestApplyTheme:
    def test_accent_and_text_colors(self):
        """テーマ色が見出し系・本文系のフォント設定に展開される"""
        config = apply_theme(
            {"theme": {"accent_color": [1, 2, 3], "text_color": [4, 5, 6]}, "fonts": {}}
        )

        assert config["fonts"]["title_h1"]["color_rgb"] == [1, 2, 3]
        assert config["fonts"]["table_header"]["color_rgb"] == [1, 2, 3]
        assert config["fonts"]["body"]["color_rgb"] == [4, 5, 6]
        assert config["fonts"]["table_body"]["color_rgb"] == [4, 5, 6]

    def test_existing_font_settings_are_kept(self):
        """既存のフォント設定（サイズ等）は保持したまま色だけ上書きする"""
        config = apply_theme(
            {
                "theme": {"accent_color": [1, 2, 3]},
                "fonts": {"title_h1": {"name": "Meiryo", "size_pt": 44}},
            }
        )

        assert config["fonts"]["title_h1"]["name"] == "Meiryo"
        assert config["fonts"]["title_h1"]["color_rgb"] == [1, 2, 3]

    def test_without_fonts_section(self):
        """fontsセクションが無くても展開できる"""
        config = apply_theme({"theme": {"accent_color": [1, 2, 3]}})
        assert config["fonts"]["title_h1"]["color_rgb"] == [1, 2, 3]

    @pytest.mark.parametrize("config", [{}, {"theme": None}, {"theme": {}}])
    def test_without_theme_is_noop(self, config):
        """テーマが無い場合は設定を変更しない"""
        assert apply_theme(config) == config


class TestCli:
    def test_parse_args_defaults(self):
        """出力先と設定ファイルの既定値"""
        args = parse_args(["input.md"])
        assert (args.input, args.output, args.config) == (
            "input.md",
            "output.pptx",
            "config.yaml",
        )

    def test_parse_args_options(self):
        """オプション指定が反映される"""
        args = parse_args(["in.md", "-o", "out.pptx", "-c", "my.yaml"])
        assert (args.output, args.config) == ("out.pptx", "my.yaml")

    def test_load_config(self, tmp_path):
        """YAML設定ファイルを辞書として読み込む"""
        path = tmp_path / "c.yaml"
        path.write_text("slides:\n  layout: '4:3'\n", encoding="utf-8")
        assert load_config(str(path)) == {"slides": {"layout": "4:3"}}

    def test_load_config_empty_file(self, tmp_path):
        """空のYAMLは空の辞書として扱う（Noneを返さない）"""
        path = tmp_path / "empty.yaml"
        path.write_text("", encoding="utf-8")
        assert load_config(str(path)) == {}

    def test_read_text_file(self, tmp_path):
        """UTF-8のテキストを読み込む"""
        path = tmp_path / "a.md"
        path.write_text("# 日本語", encoding="utf-8")
        assert read_text_file(str(path)) == "# 日本語"

    def _write_project(self, tmp_path, config_text="slides:\n  layout: '16:9'\n"):
        md = tmp_path / "in.md"
        md.write_text("# タイトル\n本文\n", encoding="utf-8")
        conf = tmp_path / "c.yaml"
        conf.write_text(config_text, encoding="utf-8")
        return str(md), str(conf), str(tmp_path / "out.pptx")

    def test_main_success(self, tmp_path, capsys):
        """正常系では0を返しファイルを生成する"""
        md, conf, out = self._write_project(tmp_path)

        assert main([md, "-o", out, "-c", conf]) == 0
        assert os.path.exists(out)
        assert "Success" in capsys.readouterr().out

    def test_main_with_empty_config(self, tmp_path):
        """空の設定ファイルでも変換できる"""
        md, conf, out = self._write_project(tmp_path, config_text="")
        assert main([md, "-o", out, "-c", conf]) == 0

    def test_main_applies_theme(self, tmp_path, mocker):
        """テーマ設定がジェネレーターに渡る設定へ反映される"""
        md, conf, out = self._write_project(
            tmp_path, config_text="theme:\n  accent_color: [1, 2, 3]\n"
        )
        spy = mocker.spy(md2pptx, "PPTXGenerator")

        assert main([md, "-o", out, "-c", conf]) == 0
        assert spy.call_args[0][0]["fonts"]["title_h1"]["color_rgb"] == [1, 2, 3]

    def test_main_uses_sys_argv_by_default(self, tmp_path, mocker):
        """引数を渡さない場合はsys.argvを解析する"""
        md, conf, out = self._write_project(tmp_path)
        mocker.patch.object(md2pptx.sys, "argv", ["md2pptx.py", md, "-o", out, "-c", conf])
        assert main() == 0

    def test_main_missing_input(self, tmp_path, capsys):
        """入力ファイルが無い場合は1を返す"""
        _, conf, out = self._write_project(tmp_path)

        assert main(["no_such.md", "-o", out, "-c", conf]) == 1
        assert "入力ファイル" in capsys.readouterr().out

    def test_main_missing_config(self, tmp_path, capsys):
        """設定ファイルが無い場合は1を返す"""
        md, _, out = self._write_project(tmp_path)

        assert main([md, "-o", out, "-c", "no_such.yaml"]) == 1
        assert "設定ファイル" in capsys.readouterr().out

    def test_main_permission_error(self, tmp_path, mocker, capsys):
        """出力先に書き込めない場合は専用のメッセージを表示する"""
        md, conf, out = self._write_project(tmp_path)
        mocker.patch.object(PPTXGenerator, "generate", side_effect=PermissionError)

        assert main([md, "-o", out, "-c", conf]) == 1
        assert "書き込めません" in capsys.readouterr().out

    def test_main_unexpected_error(self, tmp_path, mocker, capsys):
        """想定外の例外はメッセージとトレースバックを表示して1を返す"""
        md, conf, out = self._write_project(tmp_path)
        mocker.patch.object(PPTXGenerator, "generate", side_effect=ValueError("boom"))

        assert main([md, "-o", out, "-c", conf]) == 1
        captured = capsys.readouterr()
        assert "予期せぬエラー" in captured.out
        assert "ValueError" in captured.err
