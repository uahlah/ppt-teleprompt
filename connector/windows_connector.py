"""
Conector Windows: usa automação COM (pywin32) para falar com uma instância
do PowerPoint que já esteja aberta (não abre uma nova).

Estratégia:
  - Conecta via `GetActiveObject` (só funciona se o PowerPoint já estiver rodando).
  - Assina os eventos `SlideShowNextSlide` / `SlideShowBegin` / `SlideShowEnd`
    do objeto `Application`, que disparam a cada troca de slide durante a
    apresentação (inclusive avançar, voltar ou saltar para um slide específico).
  - Além dos eventos, faz um polling de segurança a cada ~0,5s, para cobrir
    qualquer cenário em que o evento COM não seja entregue.

Requer: pip install pywin32
"""

from .base import BaseConnector, SlideState

try:
    import pythoncom
    import win32com.client
    import pywintypes

    PYWIN32_AVAILABLE = True
except ImportError:  # ambiente sem pywin32 (ex.: quem só quer olhar o código)
    PYWIN32_AVAILABLE = False

# Valor da constante ppPlaceholderNotesSlideImage — usado para pular a
# miniatura do slide que aparece na página de notas (não é texto).
NOTES_SLIDE_IMAGE_PLACEHOLDER = 101
MSO_PLACEHOLDER = 14


class WindowsConnector(BaseConnector):
    SAFETY_POLL_INTERVAL = 0.5

    def platform_name(self):
        return "windows"

    # ------------------------------------------------------------------
    def _run(self):
        if not PYWIN32_AVAILABLE:
            self.on_update(
                SlideState(
                    connected=False,
                    platform="windows",
                    error="pywin32 não está instalado. Rode: pip install pywin32",
                )
            )
            return

        pythoncom.CoInitialize()
        try:
            while not self._stop_event.is_set():
                app = self._connect()
                if app is None:
                    self._emit(
                        SlideState(
                            connected=False,
                            platform="windows",
                            error="PowerPoint não encontrado em execução. Abra o arquivo e inicie a apresentação (F5).",
                        )
                    )
                    self._stop_event.wait(self.SAFETY_POLL_INTERVAL)
                    continue

                # Assina eventos; se falhar, seguimos só com polling.
                try:
                    events = win32com.client.WithEvents(app, _PPTEvents)
                    events.set_connector(self)
                except Exception:
                    pass

                # Loop interno: bombeia mensagens COM (para os eventos
                # chegarem) e faz polling de segurança.
                while not self._stop_event.is_set():
                    try:
                        pythoncom.PumpWaitingMessages()
                    except Exception:
                        pass
                    try:
                        state = self._poll_once(app)
                    except pywintypes.com_error:
                        # PowerPoint provavelmente foi fechado; tenta reconectar.
                        break
                    except Exception as exc:
                        state = SlideState(connected=False, platform="windows", error=str(exc))
                    self._emit(state)
                    self._stop_event.wait(0.25)
        finally:
            pythoncom.CoUninitialize()

    # ------------------------------------------------------------------
    def _connect(self):
        try:
            return win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            return None

    def _poll_once(self, app=None):
        if app is None:
            app = self._connect()
        if app is None:
            return SlideState(connected=False, platform="windows", error="PowerPoint não está em execução.")
        try:
            if app.SlideShowWindows.Count == 0:
                return SlideState(
                    connected=True,
                    in_slideshow=False,
                    platform="windows",
                    error="PowerPoint aberto, mas nenhuma apresentação em modo slideshow (pressione F5).",
                )
            wn = app.SlideShowWindows(1)
            view = wn.View
            idx = view.CurrentShowPosition
            pres = wn.Presentation
            total = pres.Slides.Count
            slide = pres.Slides(idx)
            notes = self._extract_notes(slide)
            title = self._extract_title(slide)
            return SlideState(
                connected=True,
                in_slideshow=True,
                slide_index=idx,
                slide_count=total,
                notes=notes,
                slide_title=title,
                platform="windows",
            )
        except Exception as exc:
            return SlideState(connected=True, in_slideshow=False, platform="windows", error=str(exc))

    @staticmethod
    def _extract_notes(slide):
        try:
            notes_page = slide.NotesPage
        except Exception:
            return ""
        parts = []
        for shape in notes_page.Shapes:
            try:
                if shape.Type == MSO_PLACEHOLDER and getattr(shape, "PlaceholderFormat", None) is not None:
                    if shape.PlaceholderFormat.Type == NOTES_SLIDE_IMAGE_PLACEHOLDER:
                        continue  # é a miniatura do slide, não texto
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    text = shape.TextFrame.TextRange.Text
                    if text and text.strip():
                        parts.append(text.strip())
            except Exception:
                continue
        return "\n\n".join(parts)

    @staticmethod
    def _extract_title(slide):
        try:
            for shape in slide.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText and getattr(shape, "Type", None) == MSO_PLACEHOLDER:
                    ph_type = shape.PlaceholderFormat.Type
                    if ph_type in (13, 1):  # ppPlaceholderTitle / ppPlaceholderCenterTitle
                        return shape.TextFrame.TextRange.Text.strip()
        except Exception:
            pass
        return ""


if PYWIN32_AVAILABLE:

    class _PPTEvents:
        """Sink de eventos COM do PowerPoint.Application.

        Os nomes dos métodos (OnSlideShowNextSlide etc.) são impostos pela
        biblioteca de tipos do PowerPoint — não são livres para renomear.
        """

        def set_connector(self, connector):
            self._connector = connector

        def OnSlideShowNextSlide(self, Wn):
            self._push()

        def OnSlideShowBegin(self, Wn):
            self._push()

        def OnSlideShowEnd(self, Pres):
            self._push()

        def _push(self):
            connector = getattr(self, "_connector", None)
            if connector is None:
                return
            try:
                state = connector._poll_once()
                connector._emit(state)
            except Exception:
                pass
