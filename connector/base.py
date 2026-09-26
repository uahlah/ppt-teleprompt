"""
Interface comum para os "conectores" de PowerPoint (Windows/Mac).

Cada conector é responsável por:
  1. Detectar se há uma apresentação de slides em execução.
  2. Descobrir o índice do slide atual e o total de slides.
  3. Extrair o texto das notas do slide atual.
  4. Chamar `on_update(state)` sempre que algo mudar.

O restante da aplicação (servidor local + tela de teleprompter) não sabe
nem se importa se está rodando no Windows ou no Mac — só recebe objetos
`SlideState` através do callback.
"""

from dataclasses import dataclass, field
from typing import Callable, Optional
import threading
import time


@dataclass
class SlideState:
    connected: bool = False           # há um PowerPoint com slideshow ativo?
    in_slideshow: bool = False        # está em modo de apresentação (F5)?
    slide_index: int = 0              # 1-based, como o PowerPoint numera
    slide_count: int = 0
    notes: str = ""
    slide_title: str = ""
    platform: str = ""
    error: Optional[str] = None
    updated_at: float = field(default_factory=time.time)


class BaseConnector:
    """Classe base. Subclasses implementam `_poll_once` e/ou eventos nativos."""

    #: intervalo de polling de segurança (segundos) — usado como rede de
    #: proteção mesmo em conectores que também escutam eventos nativos.
    SAFETY_POLL_INTERVAL = 1.0

    def __init__(self, on_update: Callable[[SlideState], None]):
        self.on_update = on_update
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._last_state: Optional[SlideState] = None

    # --- API pública -----------------------------------------------------
    def start(self):
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop_event.set()
        if self._thread:
            self._thread.join(timeout=2)

    # --- a implementar nas subclasses ------------------------------------
    def _run(self):
        """Loop principal da thread. Implementação default: polling simples."""
        while not self._stop_event.is_set():
            try:
                state = self._poll_once()
            except Exception as exc:  # nunca deixar a thread morrer silenciosamente
                state = SlideState(connected=False, error=str(exc), platform=self.platform_name())
            self._emit(state)
            self._stop_event.wait(self.SAFETY_POLL_INTERVAL)

    def _poll_once(self) -> SlideState:
        raise NotImplementedError

    def platform_name(self) -> str:
        return "unknown"

    # --- utilitário --------------------------------------------------------
    def _emit(self, state: SlideState):
        """Só chama on_update se o estado realmente mudou (evita spam)."""
        if self._last_state is None or self._state_changed(self._last_state, state):
            self._last_state = state
            self.on_update(state)

    @staticmethod
    def _state_changed(a: SlideState, b: SlideState) -> bool:
        return (
            a.connected != b.connected
            or a.in_slideshow != b.in_slideshow
            or a.slide_index != b.slide_index
            or a.slide_count != b.slide_count
            or a.notes != b.notes
            or a.slide_title != b.slide_title
            or a.error != b.error
        )
