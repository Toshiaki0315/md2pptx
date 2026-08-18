"""MarkdownをPowerPointに変換するGUI

コマンドラインを使わずに変換できる画面を提供する。

    python gui.py

画面で指定した内容は一時的なYAML設定に書き出し、CLI（md2pptx.py）を
そのまま呼び出して変換する。変換処理の実体を二重に持たないため、
CLIとGUIで結果が食い違わない。
"""

from __future__ import annotations

import os
import queue
import sys
import threading
from functools import partial
from typing import Any

import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, ttk

from gui_deps import exit_if_missing

#: このスクリプトが置かれているディレクトリ（CLIと既定の設定ファイルの場所）
APP_DIR = os.path.dirname(os.path.abspath(__file__))

exit_if_missing(APP_DIR)

import yaml  # noqa: E402  （上の確認を通ってから読み込む）

from make_template import create_template
from gui_runner import (
    DEFAULT_CONFIG_PATH,
    hex_to_rgb,
    open_in_file_manager,
    rgb_to_hex,
    run_conversion,
)
from gui_config import (
    ASPECTS,
    FONT_LABELS,
    H3_CHOICES,
    MERMAID_CHOICES,
    RGB,
    FontSetting,
    GuiSettings,
    build_config,
    default_output_path,
    settings_from_config,
    validate_inputs,
)

PADDING = 8


class App:
    """変換画面"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title('md2pptx - MarkdownをPowerPointに変換')

        # 変換中のプロセスからの出力を受け取る（別スレッド → 画面）
        self.log_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.running = False
        self.last_output: str = ''
        # 出力ファイル名を利用者が編集したか（未編集ならMarkdown名から自動で決める）
        self.output_edited = False

        self._build_variables()
        self._build_widgets()
        # 画面部品を触るため、部品を作り終えてから連動を有効にする
        self._register_traces()
        self._load_default_config()
        self._sync_enabled()

        self.root.after(100, self._drain_log_queue)

    # ------------------------------------------------------------------
    # 画面の組み立て
    # ------------------------------------------------------------------
    def _build_variables(self) -> None:
        defaults = GuiSettings()

        self.input_var = tk.StringVar()
        self.use_template_var = tk.BooleanVar(value=defaults.use_template)
        self.template_var = tk.StringVar(value=defaults.template_path)
        self.output_var = tk.StringVar(value=defaults.output_path)

        self.aspect_var = tk.StringVar(value=defaults.aspect)
        self.h3_var = tk.StringVar(value=H3_CHOICES[0][0])
        self.slide_number_var = tk.BooleanVar(value=defaults.show_slide_number)
        self.template_fonts_var = tk.BooleanVar(value=defaults.use_template_fonts)
        self.layout_title_var = tk.StringVar(value=defaults.layout_title)
        self.layout_content_var = tk.StringVar(value=defaults.layout_content)

        self.footer_text_var = tk.StringVar(value=defaults.footer_text)
        self.footer_date_var = tk.BooleanVar(value=defaults.footer_date)
        self.footer_on_title_var = tk.BooleanVar(value=defaults.footer_on_title)

        self.colors: dict[str, RGB] = {
            'accent_color': defaults.accent_color,
            'text_color': defaults.text_color,
            'code_bg_color': defaults.code_bg_color,
        }

        self.font_name_vars: dict[str, tk.StringVar] = {}
        self.font_size_vars: dict[str, tk.StringVar] = {}
        for key, _label in FONT_LABELS:
            font = defaults.fonts.get(key, FontSetting())
            self.font_name_vars[key] = tk.StringVar(value=font.name)
            self.font_size_vars[key] = tk.StringVar(value=self._format_number(font.size_pt))

        self.mermaid_var = tk.StringVar(value=MERMAID_CHOICES[0][0])
        self.endpoint_var = tk.StringVar(value=defaults.mermaid_endpoint)

        self.image_height_var = tk.StringVar(value=self._format_number(defaults.image_height_inches))
        self.image_downscale_var = tk.BooleanVar(value=defaults.image_downscale)
        self.image_dpi_var = tk.StringVar(value=self._format_number(defaults.image_dpi))

        self.status_var = tk.StringVar(value='Markdownファイルを選んで「生成」を押してください。')

    def _register_traces(self) -> None:
        """入力に応じて他の項目を追従させる"""
        self.input_var.trace_add('write', lambda *_: self._on_input_changed())
        self.use_template_var.trace_add('write', lambda *_: self._sync_enabled())
        self.output_var.trace_add('write', lambda *_: self._sync_enabled())

    def _build_widgets(self) -> None:
        main = ttk.Frame(self.root, padding=PADDING * 2)
        main.pack(fill='both', expand=True)
        main.columnconfigure(0, weight=1)

        self._build_basic(main)

        self.detail_button = ttk.Button(main, text='詳細設定を表示 ▼', command=self._toggle_details)
        self.detail_button.grid(row=1, column=0, sticky='w', pady=(PADDING, 0))

        self.details = ttk.Frame(main)
        self.details.columnconfigure(0, weight=1)
        self._build_details(self.details)
        self.details_visible = False

        self._build_actions(main)

        main.rowconfigure(4, weight=1)

    def _build_basic(self, parent: ttk.Widget) -> None:
        frame = ttk.LabelFrame(parent, text='基本設定', padding=PADDING)
        frame.grid(row=0, column=0, sticky='ew')
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text='Markdownファイル').grid(row=0, column=0, sticky='w', pady=4)
        ttk.Entry(frame, textvariable=self.input_var).grid(row=0, column=1, sticky='ew', padx=PADDING)
        ttk.Button(frame, text='参照...', command=self._browse_input).grid(row=0, column=2)

        ttk.Checkbutton(
            frame, text='テンプレート（社内フォーマットなど）を使う', variable=self.use_template_var
        ).grid(row=1, column=0, columnspan=3, sticky='w', pady=(PADDING, 0))

        self.template_label = ttk.Label(frame, text='テンプレートファイル')
        self.template_label.grid(row=2, column=0, sticky='w', pady=4)
        self.template_entry = ttk.Entry(frame, textvariable=self.template_var)
        self.template_entry.grid(row=2, column=1, sticky='ew', padx=PADDING)
        self.template_button = ttk.Button(frame, text='参照...', command=self._browse_template)
        self.template_button.grid(row=2, column=2)

        ttk.Button(
            frame,
            text='完成した資料からテンプレートを作る...',
            command=self._create_template_from_deck,
        ).grid(row=3, column=1, sticky='w', padx=PADDING, pady=(0, 4))

        ttk.Label(frame, text='出力ファイル').grid(row=4, column=0, sticky='w', pady=4)
        output_entry = ttk.Entry(frame, textvariable=self.output_var)
        output_entry.grid(row=4, column=1, sticky='ew', padx=PADDING)
        output_entry.bind('<Key>', lambda _event: setattr(self, 'output_edited', True))
        ttk.Button(frame, text='参照...', command=self._browse_output).grid(row=4, column=2)

    def _build_details(self, parent: ttk.Widget) -> None:
        notebook = ttk.Notebook(parent)
        notebook.grid(row=0, column=0, sticky='ew', pady=(PADDING, 0))

        self._build_slide_tab(notebook)
        self._build_design_tab(notebook)
        self._build_figure_tab(notebook)

        buttons = ttk.Frame(parent)
        buttons.grid(row=1, column=0, sticky='e', pady=(4, 0))
        ttk.Button(buttons, text='設定を読み込み...', command=self._load_config_dialog).pack(
            side='left', padx=(0, 4)
        )
        ttk.Button(buttons, text='設定を保存...', command=self._save_config_dialog).pack(side='left')

    def _build_slide_tab(self, notebook: ttk.Notebook) -> None:
        tab = ttk.Frame(notebook, padding=PADDING)
        notebook.add(tab, text='スライド')
        tab.columnconfigure(1, weight=1)

        ttk.Label(tab, text='画角（縦横比）').grid(row=0, column=0, sticky='w', pady=4)
        self.aspect_combo = ttk.Combobox(
            tab, textvariable=self.aspect_var, values=list(ASPECTS), state='readonly', width=12
        )
        self.aspect_combo.grid(row=0, column=1, sticky='w', padx=PADDING)
        self.aspect_note = ttk.Label(tab, text='※テンプレート使用時はテンプレート側の設定に従います')
        self.aspect_note.grid(row=1, column=0, columnspan=2, sticky='w')

        ttk.Label(tab, text='### の扱い').grid(row=2, column=0, sticky='w', pady=4)
        ttk.Combobox(
            tab,
            textvariable=self.h3_var,
            values=[label for label, _ in H3_CHOICES],
            state='readonly',
            width=28,
        ).grid(row=2, column=1, sticky='w', padx=PADDING)

        ttk.Checkbutton(tab, text='ページ番号を表示する（右下）', variable=self.slide_number_var).grid(
            row=3, column=0, columnspan=2, sticky='w', pady=4
        )

        self.template_fonts_check = ttk.Checkbutton(
            tab, text='テンプレートのフォントをそのまま使う', variable=self.template_fonts_var
        )
        self.template_fonts_check.grid(row=4, column=0, columnspan=2, sticky='w')

        self.layout_title_label = ttk.Label(tab, text='表紙のレイアウト名')
        self.layout_title_label.grid(row=5, column=0, sticky='w', pady=4)
        self.layout_title_entry = ttk.Entry(tab, textvariable=self.layout_title_var)
        self.layout_title_entry.grid(row=5, column=1, sticky='ew', padx=PADDING)

        self.layout_content_label = ttk.Label(tab, text='本文のレイアウト名')
        self.layout_content_label.grid(row=6, column=0, sticky='w', pady=4)
        self.layout_content_entry = ttk.Entry(tab, textvariable=self.layout_content_var)
        self.layout_content_entry.grid(row=6, column=1, sticky='ew', padx=PADDING)

        footer = ttk.LabelFrame(tab, text='フッター', padding=PADDING)
        footer.grid(row=7, column=0, columnspan=2, sticky='ew', pady=(PADDING, 0))
        footer.columnconfigure(1, weight=1)
        ttk.Label(footer, text='中央に出す文言').grid(row=0, column=0, sticky='w')
        ttk.Entry(footer, textvariable=self.footer_text_var).grid(
            row=0, column=1, sticky='ew', padx=PADDING
        )
        ttk.Checkbutton(footer, text='変換日を表示する', variable=self.footer_date_var).grid(
            row=1, column=0, columnspan=2, sticky='w', pady=(4, 0)
        )
        ttk.Checkbutton(footer, text='表紙にも表示する', variable=self.footer_on_title_var).grid(
            row=2, column=0, columnspan=2, sticky='w'
        )

    def _build_design_tab(self, notebook: ttk.Notebook) -> None:
        tab = ttk.Frame(notebook, padding=PADDING)
        notebook.add(tab, text='配色・フォント')
        tab.columnconfigure(3, weight=1)

        # 色見本は tk.Label で描く（macOSでは tk.Button の背景色が反映されないため）
        self.color_swatches: dict[str, tk.Label] = {}
        labels = (
            ('accent_color', '見出しの色'),
            ('text_color', '本文の色'),
            ('code_bg_color', 'コードの背景色'),
        )
        for row, (key, label) in enumerate(labels):
            ttk.Label(tab, text=label).grid(row=row, column=0, sticky='w', pady=3)
            swatch = tk.Label(tab, width=6, relief='groove', background=rgb_to_hex(self.colors[key]))
            swatch.grid(row=row, column=1, sticky='w', padx=PADDING)
            swatch.bind('<Button-1>', partial(self._on_swatch_click, key))
            self.color_swatches[key] = swatch
            ttk.Button(tab, text='色を選ぶ...', command=partial(self._pick_color, key)).grid(
                row=row, column=2, sticky='w'
            )

        fonts = ttk.LabelFrame(tab, text='フォント（空欄にすると指定しません）', padding=PADDING)
        fonts.grid(row=len(labels), column=0, columnspan=4, sticky='ew', pady=(PADDING, 0))
        fonts.columnconfigure(1, weight=1)

        ttk.Label(fonts, text='書体').grid(row=0, column=1, sticky='w', padx=PADDING)
        ttk.Label(fonts, text='サイズ(pt)').grid(row=0, column=2, sticky='w')
        for row, (key, label) in enumerate(FONT_LABELS, start=1):
            ttk.Label(fonts, text=label).grid(row=row, column=0, sticky='w', pady=3)
            ttk.Entry(fonts, textvariable=self.font_name_vars[key]).grid(
                row=row, column=1, sticky='ew', padx=PADDING
            )
            ttk.Spinbox(
                fonts, from_=6, to=200, increment=1, width=6, textvariable=self.font_size_vars[key]
            ).grid(row=row, column=2, sticky='w')

    def _build_figure_tab(self, notebook: ttk.Notebook) -> None:
        tab = ttk.Frame(notebook, padding=PADDING)
        notebook.add(tab, text='図・画像')
        tab.columnconfigure(1, weight=1)

        ttk.Label(tab, text='Mermaid図の生成方法').grid(row=0, column=0, sticky='w', pady=4)
        self.mermaid_combo = ttk.Combobox(
            tab,
            textvariable=self.mermaid_var,
            values=[label for label, _ in MERMAID_CHOICES],
            state='readonly',
            width=32,
        )
        self.mermaid_combo.grid(row=0, column=1, sticky='w', padx=PADDING)
        self.mermaid_combo.bind('<<ComboboxSelected>>', lambda _event: self._sync_enabled())
        ttk.Label(
            tab, text='※Krokiは図の内容を外部サービスへ送信します（社内Krokiのアドレス指定も可）'
        ).grid(row=1, column=0, columnspan=2, sticky='w')

        self.endpoint_label = ttk.Label(tab, text='Krokiのアドレス')
        self.endpoint_label.grid(row=2, column=0, sticky='w', pady=4)
        self.endpoint_entry = ttk.Entry(tab, textvariable=self.endpoint_var)
        self.endpoint_entry.grid(row=2, column=1, sticky='ew', padx=PADDING)

        images = ttk.LabelFrame(tab, text='画像', padding=PADDING)
        images.grid(row=3, column=0, columnspan=2, sticky='ew', pady=(PADDING, 0))
        images.columnconfigure(1, weight=1)
        ttk.Label(images, text='既定の高さ（インチ）').grid(row=0, column=0, sticky='w', pady=3)
        ttk.Spinbox(
            images, from_=0.5, to=10, increment=0.5, width=6, textvariable=self.image_height_var
        ).grid(row=0, column=1, sticky='w', padx=PADDING)
        ttk.Checkbutton(
            images,
            text='大きすぎる画像を自動で縮小する（ファイルサイズ削減）',
            variable=self.image_downscale_var,
        ).grid(row=1, column=0, columnspan=2, sticky='w')
        ttk.Label(images, text='縮小後の解像度（dpi）').grid(row=2, column=0, sticky='w', pady=3)
        ttk.Spinbox(
            images, from_=48, to=600, increment=6, width=6, textvariable=self.image_dpi_var
        ).grid(row=2, column=1, sticky='w', padx=PADDING)

    def _build_actions(self, parent: ttk.Widget) -> None:
        actions = ttk.Frame(parent)
        actions.grid(row=3, column=0, sticky='ew', pady=(PADDING * 2, 0))
        actions.columnconfigure(1, weight=1)

        self.generate_button = ttk.Button(actions, text='生成', command=self._generate)
        self.generate_button.grid(row=0, column=0)
        ttk.Label(actions, textvariable=self.status_var).grid(row=0, column=1, sticky='w', padx=PADDING)
        self.open_button = ttk.Button(
            actions, text='出力先を開く', command=self._open_output, state='disabled'
        )
        self.open_button.grid(row=0, column=2)

        log_frame = ttk.LabelFrame(parent, text='ログ', padding=4)
        log_frame.grid(row=4, column=0, sticky='nsew', pady=(PADDING, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = tk.Text(log_frame, height=8, wrap='word', state='disabled')
        self.log.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(log_frame, orient='vertical', command=self.log.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.log.configure(yscrollcommand=scrollbar.set)

    # ------------------------------------------------------------------
    # 画面の状態
    # ------------------------------------------------------------------
    def _toggle_details(self) -> None:
        if self.details_visible:
            self.details.grid_forget()
            self.detail_button.configure(text='詳細設定を表示 ▼')
        else:
            self.details.grid(row=2, column=0, sticky='ew')
            self.detail_button.configure(text='詳細設定を隠す ▲')
        self.details_visible = not self.details_visible

    def _sync_enabled(self) -> None:
        """他の設定によって意味を持たなくなる項目を無効化する"""
        use_template = self.use_template_var.get()
        template_state = 'normal' if use_template else 'disabled'
        for widget in (self.template_entry, self.template_button):
            widget.configure(state=template_state)
        self.template_label.configure(state=template_state)

        self.aspect_combo.configure(state='disabled' if use_template else 'readonly')
        self.aspect_note.configure(state='normal' if use_template else 'disabled')
        self.template_fonts_check.configure(state=template_state)
        for widget in (self.layout_title_entry, self.layout_content_entry):
            widget.configure(state=template_state)
        for label in (self.layout_title_label, self.layout_content_label):
            label.configure(state=template_state)

        kroki = self._selected_value(self.mermaid_var, MERMAID_CHOICES) == 'kroki'
        self.endpoint_entry.configure(state='normal' if kroki else 'disabled')
        self.endpoint_label.configure(state='normal' if kroki else 'disabled')

        if not self.running:
            ready = bool(self.input_var.get().strip() and self.output_var.get().strip())
            self.generate_button.configure(state='normal' if ready else 'disabled')

    def _on_input_changed(self) -> None:
        if not self.output_edited:
            # Markdownのファイル名から自動で決める（利用者が触ったら追従しない）
            self.output_var.set(default_output_path(self.input_var.get()))
        self._sync_enabled()

    # ------------------------------------------------------------------
    # ファイル選択
    # ------------------------------------------------------------------
    def _browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title='Markdownファイルを選択',
            filetypes=[('Markdown', '*.md *.markdown'), ('すべてのファイル', '*.*')],
        )
        if path:
            self.input_var.set(path)

    def _browse_template(self) -> None:
        path = filedialog.askopenfilename(
            title='テンプレートを選択',
            filetypes=[('PowerPointテンプレート', '*.pptx *.potx'), ('すべてのファイル', '*.*')],
        )
        if path:
            self.template_var.set(path)

    def _browse_output(self) -> None:
        current = self.output_var.get()
        path = filedialog.asksaveasfilename(
            title='保存先を選択',
            defaultextension='.pptx',
            initialfile=os.path.basename(current) or 'output.pptx',
            initialdir=os.path.dirname(current) or None,
            filetypes=[('PowerPoint', '*.pptx')],
        )
        if path:
            self.output_var.set(path)
            self.output_edited = True

    def _on_swatch_click(self, key: str, _event: 'tk.Event[tk.Label]') -> None:
        self._pick_color(key)

    def _create_template_from_deck(self) -> None:
        """完成した資料からテンプレート（書式だけのファイル）を作る

        資料をそのままテンプレートに指定すると中身が出力に残るため、
        画面から作れるようにしてある。
        """
        source = filedialog.askopenfilename(
            title='元になるPowerPoint資料を選択',
            filetypes=[('PowerPoint', '*.pptx *.potx'), ('すべてのファイル', '*.*')],
        )
        if not source:
            return

        base = os.path.splitext(os.path.basename(source))[0]
        destination = filedialog.asksaveasfilename(
            title='テンプレートの保存先',
            defaultextension='.pptx',
            initialfile=f'{base}-template.pptx',
            initialdir=os.path.dirname(source) or None,
            filetypes=[('PowerPoint', '*.pptx')],
        )
        if not destination:
            return

        try:
            removed = create_template(source, destination)
        except Exception as e:
            messagebox.showerror('テンプレートの作成', f'テンプレートを作成できませんでした。\n{e}')
            return

        # 作ったテンプレートをそのまま使えるようにする
        self.template_var.set(destination)
        self.use_template_var.set(True)
        self.status_var.set(
            f'{os.path.basename(destination)} を作成しました（スライド{removed}枚を取り除きました）。'
        )

    def _pick_color(self, key: str) -> None:
        chosen = colorchooser.askcolor(color=rgb_to_hex(self.colors[key]))
        if chosen and chosen[1]:
            self.colors[key] = hex_to_rgb(str(chosen[1]))
            self.color_swatches[key].configure(background=rgb_to_hex(self.colors[key]))

    # ------------------------------------------------------------------
    # 設定値の受け渡し
    # ------------------------------------------------------------------
    @staticmethod
    def _format_number(value: float) -> str:
        return str(int(value)) if float(value).is_integer() else str(value)

    @staticmethod
    def _read_float(var: tk.StringVar, default: float) -> float:
        try:
            return float(var.get())
        except ValueError:
            return default

    @staticmethod
    def _selected_value(var: tk.StringVar, choices: tuple[tuple[str, str], ...]) -> str:
        for label, value in choices:
            if label == var.get():
                return value
        return choices[0][1]

    @staticmethod
    def _label_of(value: str, choices: tuple[tuple[str, str], ...]) -> str:
        for label, choice in choices:
            if choice == value:
                return label
        return choices[0][0]

    def collect_settings(self) -> GuiSettings:
        """画面の入力内容を設定オブジェクトにまとめる"""
        defaults = GuiSettings()
        fonts = {
            key: FontSetting(
                name=self.font_name_vars[key].get(),
                size_pt=self._read_float(self.font_size_vars[key], 0),
            )
            for key, _label in FONT_LABELS
        }
        return GuiSettings(
            input_path=self.input_var.get().strip(),
            use_template=self.use_template_var.get(),
            template_path=self.template_var.get().strip(),
            output_path=self.output_var.get().strip(),
            aspect=self.aspect_var.get(),
            h3_as=self._selected_value(self.h3_var, H3_CHOICES),
            show_slide_number=self.slide_number_var.get(),
            use_template_fonts=self.template_fonts_var.get(),
            layout_title=self.layout_title_var.get(),
            layout_content=self.layout_content_var.get(),
            footer_text=self.footer_text_var.get(),
            footer_date=self.footer_date_var.get(),
            footer_on_title=self.footer_on_title_var.get(),
            accent_color=self.colors['accent_color'],
            text_color=self.colors['text_color'],
            code_bg_color=self.colors['code_bg_color'],
            fonts=fonts,
            mermaid_renderer=self._selected_value(self.mermaid_var, MERMAID_CHOICES),
            mermaid_endpoint=self.endpoint_var.get(),
            image_height_inches=self._read_float(self.image_height_var, defaults.image_height_inches),
            image_downscale=self.image_downscale_var.get(),
            image_dpi=self._read_float(self.image_dpi_var, defaults.image_dpi),
        )

    def apply_settings(self, settings: GuiSettings) -> None:
        """設定オブジェクトの内容を画面に反映する（詳細設定のみ）"""
        self.use_template_var.set(settings.use_template)
        if settings.template_path:
            self.template_var.set(settings.template_path)
        self.aspect_var.set(settings.aspect)
        self.h3_var.set(self._label_of(settings.h3_as, H3_CHOICES))
        self.slide_number_var.set(settings.show_slide_number)
        self.template_fonts_var.set(settings.use_template_fonts)
        self.layout_title_var.set(settings.layout_title)
        self.layout_content_var.set(settings.layout_content)
        self.footer_text_var.set(settings.footer_text)
        self.footer_date_var.set(settings.footer_date)
        self.footer_on_title_var.set(settings.footer_on_title)

        for key in self.colors:
            self.colors[key] = getattr(settings, key)
            self.color_swatches[key].configure(background=rgb_to_hex(self.colors[key]))

        for key, _label in FONT_LABELS:
            font = settings.fonts.get(key, FontSetting())
            self.font_name_vars[key].set(font.name)
            self.font_size_vars[key].set(self._format_number(font.size_pt) if font.size_pt else '')

        self.mermaid_var.set(self._label_of(settings.mermaid_renderer, MERMAID_CHOICES))
        self.endpoint_var.set(settings.mermaid_endpoint)
        self.image_height_var.set(self._format_number(settings.image_height_inches))
        self.image_downscale_var.set(settings.image_downscale)
        self.image_dpi_var.set(self._format_number(settings.image_dpi))
        self._sync_enabled()

    def _load_default_config(self) -> None:
        """同じフォルダの config.yaml があれば初期値として読み込む"""
        if os.path.isfile(DEFAULT_CONFIG_PATH):
            self._load_config_file(DEFAULT_CONFIG_PATH, quiet=True)

    def _load_config_file(self, path: str, quiet: bool = False) -> None:
        try:
            with open(path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
        except (OSError, yaml.YAMLError) as e:
            if not quiet:
                messagebox.showerror('設定の読み込み', f'設定ファイルを読み込めませんでした。\n{e}')
            return
        self.apply_settings(settings_from_config(config))
        if not quiet:
            self.status_var.set(f'{os.path.basename(path)} を読み込みました。')

    def _load_config_dialog(self) -> None:
        path = filedialog.askopenfilename(
            title='設定ファイルを選択',
            filetypes=[('YAML', '*.yaml *.yml'), ('すべてのファイル', '*.*')],
        )
        if path:
            self._load_config_file(path)

    def _save_config_dialog(self) -> None:
        path = filedialog.asksaveasfilename(
            title='設定の保存先',
            defaultextension='.yaml',
            initialfile='my-config.yaml',
            filetypes=[('YAML', '*.yaml *.yml')],
        )
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                yaml.safe_dump(
                    build_config(self.collect_settings()), f, allow_unicode=True, sort_keys=False
                )
        except OSError as e:
            messagebox.showerror('設定の保存', f'設定ファイルを保存できませんでした。\n{e}')
            return
        self.status_var.set(f'{os.path.basename(path)} に保存しました。')

    # ------------------------------------------------------------------
    # 変換の実行
    # ------------------------------------------------------------------
    def _generate(self) -> None:
        if self.running:
            return

        settings = self.collect_settings()
        errors = validate_inputs(settings)
        if errors:
            messagebox.showerror('入力内容を確認してください', '\n'.join(errors))
            return

        self._clear_log()
        self.running = True
        self.generate_button.configure(state='disabled')
        self.open_button.configure(state='disabled')
        self.status_var.set('変換中です...')

        thread = threading.Thread(target=self._run_conversion, args=(settings,), daemon=True)
        thread.start()

    def _run_conversion(self, settings: GuiSettings) -> None:
        """別スレッドでCLIを呼び出し、出力を画面へ送る"""
        returncode, output_path = run_conversion(
            settings, lambda line: self.log_queue.put(('log', line))
        )
        self.log_queue.put(('done', (returncode, output_path)))

    def _drain_log_queue(self) -> None:
        """別スレッドからの出力を画面に反映する（100msごと）"""
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == 'log':
                    self._append_log(str(payload))
                else:
                    returncode, output_path = payload
                    self._finish(int(returncode), str(output_path))
        except queue.Empty:
            pass
        self.root.after(100, self._drain_log_queue)

    def _finish(self, returncode: int, output_path: str) -> None:
        self.running = False
        self.last_output = output_path
        self._sync_enabled()
        if returncode == 0:
            self.status_var.set(f'完了しました: {os.path.basename(output_path)}')
            self.open_button.configure(state='normal')
        else:
            self.status_var.set('変換に失敗しました。ログを確認してください。')
            messagebox.showerror(
                '変換に失敗しました', '下のログにエラーの内容が表示されています。'
            )

    def _open_output(self) -> None:
        if self.last_output and os.path.exists(self.last_output):
            open_in_file_manager(self.last_output)

    def _append_log(self, text: str) -> None:
        self.log.configure(state='normal')
        self.log.insert('end', text + '\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def _clear_log(self) -> None:
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')


def force_initial_draw(root: tk.Tk) -> None:
    """起動直後にウィンドウの中身が描画されないことへの対処

    macOS標準のTk 8.5では、最初の描画が行われず真っ白なウィンドウになることがある。
    要求サイズでウィンドウを確定させ、直後に高さを1ピクセルだけ変えて再描画を促す。
    """
    root.update_idletasks()
    width, height = root.winfo_reqwidth(), root.winfo_reqheight()
    root.minsize(width, height)
    root.geometry(f'{width}x{height}')
    root.after(100, lambda: root.geometry(f'{width}x{height + 1}'))


def main() -> int:
    """GUIを起動する"""
    root = tk.Tk()
    App(root)
    force_initial_draw(root)
    root.mainloop()
    return 0


if __name__ == '__main__':  # pragma: no cover
    sys.exit(main())
