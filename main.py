#!/usr/bin/env python3
"""
Teleprompter para PowerPoint — orquestrador principal.

O que este script faz:
  1. Sobe um servidor local (HTTP + WebSocket) que serve as páginas em ui/
     e distribui o estado do slide atual para quem estiver conectado.
  2. Conecta ao PowerPoint (via COM no Windows, via AppleScript no Mac) e
     empurra o slide/notas atuais para o servidor sempre que mudam.
  3. Abre um painel de controle (janela nativa via pywebview) de onde você
     escolhe em qual monitor abrir a tela de teleprompter em tela cheia.

Uso:
    python main.py             # modo normal, com painel de controle nativo
    python main.py --no-gui    # só o servidor; abra http://127.0.0.1:8765/control.html
                                # manualmente no navegador (modo manual / debug)

Veja o README.md para instruções completas de instalação e uso.
"""

import argparse
import json
import os
import sys
import time
from dataclasses import asdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from connector import get_connector, SlideState  # noqa: E402
from server import TeleprompterServer  # noqa: E402

HTTP_PORT = 8765
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UI_DIR = os.path.join(BASE_DIR, "ui")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")


def _load_config():
    """Lê config.json (se existir) para usar como configuração inicial.
    Copie config.example.json para config.json e ajuste como preferir."""
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as exc:
            print(f"[aviso] Não foi possível ler config.json ({exc}); usando padrões.")
    return {}


class Api:
    """Métodos expostos ao JavaScript do painel de controle (control.html)
    através da ponte js_api do pywebview.

    IMPORTANTE: `self._monitors` vem de `webview.screens` (não de
    `screeninfo`). As duas bibliotecas descrevem a posição dos monitores em
    sistemas de coordenadas diferentes no macOS (uma conta a partir do
    canto superior esquerdo, a outra do inferior esquerdo), então misturar
    "listar com uma, posicionar com a outra" é o que causava abrir na tela
    errada. Usando só `webview.screens` para as duas coisas, a lista que o
    usuário vê e a posição onde a janela abre usam exatamente os mesmos
    números.
    """

    def __init__(self, monitors, http_port):
        self._monitors = monitors
        self._http_port = http_port
        self._teleprompter_window = None

    def list_monitors(self):
        result = []
        for i, m in enumerate(self._monitors):
            result.append(
                {
                    "index": i,
                    "x": m.x,
                    "y": m.y,
                    "width": m.width,
                    "height": m.height,
                    "is_primary": bool(getattr(m, "is_primary", i == 0)),
                }
            )
        return result

    def launch_teleprompter(self, index):
        if index is None or index < 0 or index >= len(self._monitors):
            return {"ok": False, "error": "Monitor inválido."}
        m = self._monitors[index]
        try:
            import webview

            # Sem frameless/on_top: mantemos a barra de título nativa (com
            # botão de fechar) como saída de emergência garantida, e a
            # tecla Q (tratada em teleprompter.html via este mesmo js_api)
            # como forma rápida de fechar sem precisar caçar o cursor.
            win = webview.create_window(
                "Teleprompter",
                url=f"http://127.0.0.1:{self._http_port}/teleprompter.html",
                x=m.x,
                y=m.y,
                width=m.width,
                height=m.height,
                fullscreen=True,
                js_api=self,
            )
            self._teleprompter_window = win
            return {"ok": True}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    def close_teleprompter(self):
        if self._teleprompter_window is not None:
            try:
                self._teleprompter_window.destroy()
            except Exception:
                pass
            self._teleprompter_window = None
        return {"ok": True}


def _detect_monitors():
    """Usa webview.screens (não screeninfo) — ver o comentário na classe
    Api sobre por que as duas bibliotecas não podem ser misturadas aqui."""
    try:
        import webview

        return list(webview.screens)
    except Exception as exc:
        print(f"[aviso] Não foi possível detectar monitores automaticamente ({exc}).")
        return []


def main():
    parser = argparse.ArgumentParser(description="Teleprompter para PowerPoint")
    parser.add_argument("--no-gui", action="store_true", help="Não abrir janela nativa; só o servidor local.")
    parser.add_argument("--port", type=int, default=HTTP_PORT, help="Porta HTTP local (padrão 8765).")
    args = parser.parse_args()

    server = TeleprompterServer(ui_dir=UI_DIR, http_port=args.port, initial_settings=_load_config())
    server.start()
    print(f"Servidor local rodando em http://127.0.0.1:{args.port}")

    def on_update(state: SlideState):
        server.push_state(asdict(state))

    try:
        connector = get_connector(on_update)
    except RuntimeError as exc:
        print(f"\n[erro] {exc}\n")
        server.stop()
        return
    connector.start()

    if args.no_gui:
        print(f"Painel de controle: http://127.0.0.1:{args.port}/control.html")
        print(f"Teleprompter (abra manualmente na 2ª tela): http://127.0.0.1:{args.port}/teleprompter.html")
        print("Pressione Ctrl+C para encerrar.")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass
        connector.stop()
        server.stop()
        return

    try:
        import webview
    except ImportError:
        print("\n[aviso] pywebview não está instalado — não é possível abrir o painel nativo")
        print("        nem a abertura automática por monitor.")
        print("        Instale com: pip install pywebview")
        print(f"        Ou use o modo manual: abra http://127.0.0.1:{args.port}/control.html no navegador.\n")
        return main_no_gui_fallback(server, connector, args.port)

    monitors = _detect_monitors()
    api = Api(monitors, args.port)

    webview.create_window(
        "Painel de Controle — Teleprompter",
        url=f"http://127.0.0.1:{args.port}/control.html",
        js_api=api,
        width=520,
        height=780,
        x=40,
        y=40,
        confirm_close=False,
    )

    try:
        webview.start()
    finally:
        connector.stop()
        server.stop()


def main_no_gui_fallback(server, connector, port):
    print("Pressione Ctrl+C para encerrar.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    connector.stop()
    server.stop()


if __name__ == "__main__":
    main()
