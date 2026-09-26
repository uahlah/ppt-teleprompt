"""
Servidor HTTP + WebSocket local (aiohttp) que:
  - serve as páginas estáticas em ui/ (teleprompter.html e control.html);
  - mantém o último estado do slide (índice, notas, etc.) e as configurações
    (velocidade de rolagem, tamanho da fonte, espelhar, tema...);
  - transmite (broadcast) qualquer atualização para todos os clientes
    conectados via WebSocket em /ws — tanto a janela de teleprompter quanto
    o painel de controle escutam o mesmo canal.

Roda inteiramente em localhost; nada sai da máquina do usuário.
"""

from __future__ import annotations

import asyncio
import threading

from aiohttp import web, WSMsgType

DEFAULT_SETTINGS = {
    "scroll_speed": 40,     # pixels por segundo
    "font_size": 48,        # px
    "line_height": 1.5,
    "theme": "dark",        # "dark" | "light"
    "mirror": False,        # espelhar horizontalmente (uso com vidro de teleprompter)
    "auto_scroll": True,
}

DEFAULT_STATE = {
    "connected": False,
    "in_slideshow": False,
    "slide_index": 0,
    "slide_count": 0,
    "notes": "",
    "slide_title": "",
    "platform": "",
    "error": "Aguardando conexão com o PowerPoint...",
}


class TeleprompterServer:
    def __init__(self, ui_dir: str, http_port: int = 8765, initial_settings: dict | None = None):
        self.ui_dir = ui_dir
        self.http_port = http_port
        self.loop: asyncio.AbstractEventLoop | None = None
        self.runner: web.AppRunner | None = None
        self.clients: set[web.WebSocketResponse] = set()
        self.state = dict(DEFAULT_STATE)
        self.settings = dict(DEFAULT_SETTINGS)
        if initial_settings:
            self.settings.update(initial_settings)

        self._thread: threading.Thread | None = None
        self._ready = threading.Event()

    # ------------------------------------------------------------------
    # ciclo de vida
    # ------------------------------------------------------------------
    def start(self):
        self._thread = threading.Thread(target=self._thread_main, daemon=True)
        self._thread.start()
        if not self._ready.wait(timeout=5):
            raise RuntimeError("Servidor local não iniciou a tempo.")

    def _thread_main(self):
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        self.loop.run_until_complete(self._start_app())
        self._ready.set()
        self.loop.run_forever()

    async def _start_app(self):
        app = web.Application()
        app.router.add_get("/ws", self._ws_handler)
        app.router.add_get("/api/settings", self._get_settings)
        app.router.add_post("/api/settings", self._post_settings)
        app.router.add_get("/api/state", self._get_state)
        app.router.add_static("/", self.ui_dir, show_index=True)

        self.runner = web.AppRunner(app)
        await self.runner.setup()
        site = web.TCPSite(self.runner, "127.0.0.1", self.http_port)
        await site.start()

    def stop(self):
        if self.loop and self.loop.is_running():
            fut = asyncio.run_coroutine_threadsafe(self._shutdown(), self.loop)
            try:
                fut.result(timeout=5)
            except Exception:
                pass

    async def _shutdown(self):
        for ws in list(self.clients):
            await ws.close()
        if self.runner:
            await self.runner.cleanup()
        self.loop.stop()

    # ------------------------------------------------------------------
    # chamadas thread-safe (usadas pelo conector e pela API do pywebview,
    # que correm em outras threads que não a do event loop do aiohttp)
    # ------------------------------------------------------------------
    def push_state(self, state_dict: dict):
        self.state = state_dict
        self._broadcast_threadsafe({"type": "state", "data": self.state})

    def update_settings(self, partial: dict) -> dict:
        self.settings.update(partial)
        self._broadcast_threadsafe({"type": "settings", "data": self.settings})
        return self.settings

    def _broadcast_threadsafe(self, message: dict):
        if self.loop and self.loop.is_running():
            asyncio.run_coroutine_threadsafe(self._broadcast(message), self.loop)

    async def _broadcast(self, message: dict):
        dead = []
        for ws in list(self.clients):
            try:
                await ws.send_json(message)
            except Exception:
                dead.append(ws)
        for d in dead:
            self.clients.discard(d)

    # ------------------------------------------------------------------
    # handlers HTTP / WS
    # ------------------------------------------------------------------
    async def _ws_handler(self, request):
        ws = web.WebSocketResponse(heartbeat=20)
        await ws.prepare(request)
        self.clients.add(ws)
        try:
            await ws.send_json({"type": "state", "data": self.state})
            await ws.send_json({"type": "settings", "data": self.settings})
            async for msg in ws:
                if msg.type == WSMsgType.TEXT:
                    await self._handle_client_message(msg.json())
                elif msg.type == WSMsgType.ERROR:
                    break
        finally:
            self.clients.discard(ws)
        return ws

    async def _handle_client_message(self, payload: dict):
        # Permite que a própria tela de teleprompter (atalhos de teclado)
        # ou o painel de controle alterem configurações e sincronizem entre si.
        if payload.get("type") == "settings_update":
            self.settings.update(payload.get("data", {}))
            await self._broadcast({"type": "settings", "data": self.settings})

    async def _get_settings(self, request):
        return web.json_response(self.settings)

    async def _post_settings(self, request):
        try:
            data = await request.json()
        except Exception:
            data = {}
        settings = self.update_settings(data)
        return web.json_response(settings)

    async def _get_state(self, request):
        return web.json_response(self.state)
