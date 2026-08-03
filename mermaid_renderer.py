"""Mermaid記法の図をPNG画像に変換する

図のソースには社内システムの構成やデータフローが含まれ得るため、
どこへ送信するか（あるいは送信しないか）を config.yaml で制御できるようにしている。
"""

from __future__ import annotations

import base64
import os
import subprocess
import tempfile
import zlib
from typing import TYPE_CHECKING, Any

import requests

if TYPE_CHECKING:
    from generator import PPTXGenerator

#: 既定のKrokiエンドポイント（公開サービス）
DEFAULT_KROKI_ENDPOINT = "https://kroki.io"

#: Krokiが応答しない場合の代替（公開サービス）
MERMAID_INK_URL = "https://mermaid.ink/img/"

#: 外部APIのタイムアウト秒数
HTTP_TIMEOUT_SEC = 15

#: ローカルレンダリング（mermaid-cli）のタイムアウト秒数
LOCAL_RENDER_TIMEOUT_SEC = 60

#: 既定のmermaid-cliコマンド
DEFAULT_CLI_PATH = "mmdc"

#: 公開サービスとして扱うホスト（警告およびフォールバック可否の判定に使う）
PUBLIC_HOSTS = ("kroki.io", "mermaid.ink")

EXTERNAL_WARNING = (
    "Warning: Mermaid図の内容を外部サービス（{endpoint}）へ送信します。\n"
    "         社外に出せない情報を含む場合は、config.yaml で\n"
    "         mermaid.renderer を 'local' または 'off' にするか、\n"
    "         mermaid.endpoint に自己ホストしたKrokiのURLを指定してください。"
)


class MermaidRenderError(Exception):
    """Mermaid図の生成に失敗したことを表す例外"""


def mermaid_conf(generator: PPTXGenerator) -> dict[str, Any]:
    """config.yaml の mermaid セクションを取得する"""
    return generator.config.get('mermaid') or {}


def is_public_endpoint(endpoint: str) -> bool:
    """指定のエンドポイントが公開サービスかどうかを判定する"""
    return any(host in endpoint for host in PUBLIC_HOSTS)


def _fallback_allowed(conf: dict[str, Any], endpoint_is_public: bool) -> bool:
    """公開API（mermaid.ink）へのフォールバックを許可するか

    明示指定が無い場合は「元から公開サービスを使っているときだけ許可」する。
    自己ホストのKrokiを指定しているのに、障害時だけ公開APIへ送ってしまうと
    情報漏洩になるため、既定では自動フォールバックしない。
    """
    setting = conf.get('fallback_to_public')
    if setting is None:
        return endpoint_is_public
    return bool(setting)


def render_mermaid(conf: dict[str, Any], text: str) -> bytes | None:
    """Mermaid記法をPNGのバイト列に変換する（renderer が 'off' の場合は None）"""
    renderer = str(conf.get('renderer') or 'kroki').lower()

    if renderer == 'off':
        print("INFO: mermaid.renderer が 'off' のため、Mermaid図の生成をスキップしました。")
        return None
    if renderer == 'local':
        return _render_locally(conf, text)
    if renderer == 'kroki':
        return _render_with_kroki(conf, text)

    raise MermaidRenderError(
        f"未知の mermaid.renderer です: '{renderer}'（kroki / local / off のいずれかを指定してください）"
    )


def _render_with_kroki(conf: dict[str, Any], text: str) -> bytes:
    """Kroki（既定は公開サービス、自己ホスト可）で変換する"""
    endpoint = str(conf.get('endpoint') or DEFAULT_KROKI_ENDPOINT).rstrip('/')
    endpoint_is_public = is_public_endpoint(endpoint)

    if endpoint_is_public and conf.get('warn_on_external', True):
        print(EXTERNAL_WARNING.format(endpoint=endpoint))

    compressed = zlib.compress(text.encode('utf-8'), 9)
    encoded = base64.urlsafe_b64encode(compressed).decode('ascii')

    try:
        response = requests.get(f"{endpoint}/mermaid/png/{encoded}", timeout=HTTP_TIMEOUT_SEC)
        response.raise_for_status()
        return response.content
    except Exception as e_kroki:
        if not _fallback_allowed(conf, endpoint_is_public):
            raise MermaidRenderError(
                f"Krokiでの生成に失敗しました: {e_kroki}\n"
                "         公開APIへの自動フォールバックは、情報漏洩を避けるため行いません"
                "（許可する場合は mermaid.fallback_to_public: true）。"
            ) from e_kroki

        print(f"INFO: Kroki APIが応答しませんでした。代替API(mermaid.ink)を試行します... ({e_kroki})")
        encoded_ink = base64.urlsafe_b64encode(text.encode('utf-8')).decode('ascii')
        response = requests.get(f"{MERMAID_INK_URL}{encoded_ink}", timeout=HTTP_TIMEOUT_SEC)
        response.raise_for_status()
        return response.content


def _render_locally(conf: dict[str, Any], text: str) -> bytes:
    """mermaid-cli（mmdc）でオフライン変換する（外部送信なし）"""
    cli_path = str(conf.get('cli_path') or DEFAULT_CLI_PATH)

    with tempfile.TemporaryDirectory() as work_dir:
        source_path = os.path.join(work_dir, 'diagram.mmd')
        output_path = os.path.join(work_dir, 'diagram.png')
        with open(source_path, 'w', encoding='utf-8') as f:
            f.write(text)

        command = [cli_path, '-i', source_path, '-o', output_path]
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                timeout=LOCAL_RENDER_TIMEOUT_SEC,
                check=False,
            )
        except FileNotFoundError as e:
            raise MermaidRenderError(
                f"'{cli_path}' が見つかりません。"
                "npm install -g @mermaid-js/mermaid-cli で導入するか、"
                "mermaid.cli_path にパスを指定してください。"
            ) from e
        except subprocess.TimeoutExpired as e:
            raise MermaidRenderError(
                f"ローカルでの図の生成が {LOCAL_RENDER_TIMEOUT_SEC} 秒でタイムアウトしました。"
            ) from e

        if result.returncode != 0:
            stderr = (result.stderr or b'').decode('utf-8', errors='replace').strip()
            raise MermaidRenderError(f"{cli_path} が異常終了しました: {stderr}")

        if not os.path.exists(output_path):
            raise MermaidRenderError(f"{cli_path} が画像を出力しませんでした。")

        with open(output_path, 'rb') as f:
            return f.read()
