"""
Conector macOS: usa AppleScript (via `osascript`) para consultar o
PowerPoint em execução.

Reescrito depois de confirmar, diretamente no dicionário AppleScript do
PowerPoint (Script Editor → File → Open Dictionary), os nomes reais das
classes/propriedades usadas aqui:
  - `application "Microsoft PowerPoint" is running` (fora de um `tell`,
    idioma padrão do AppleScript — nunca lança o app à toa)
  - `slide show window`, com propriedade `presentation`
  - `slideshow view` (uma palavra "slideshow" + "view" — NÃO é "slide show
    view", esse foi o bug que te fez perder tempo antes), com propriedade
    `slide` que devolve o objeto do slide atual diretamente
  - `notes page` do slide, com `shapes`, `text frame`, `text range`,
    `content`

Filosofia desta versão: cada etapa tem seu próprio `try` que, se falhar,
já devolve uma mensagem de erro ESPECÍFICA dizendo qual propriedade
quebrou — em vez de um "deu erro" genérico. Se algo ainda não funcionar
na sua versão do Office, a mensagem que aparecer na tela do app já aponta
exatamente onde ajustar aqui.

Estratégia: como não é prático assinar eventos COM no macOS a partir do
Python, este conector faz *polling* (padrão: a cada 400ms) chamando o
AppleScript.
"""

import base64
import subprocess
import tempfile
import os

from .base import BaseConnector, SlideState

APPLESCRIPT = r"""
on run
    try
        if not (application "Microsoft PowerPoint" is running) then
            return my buildResult(false, false, 0, 0, "PowerPoint não está em execução.")
        end if
    on error errMsg
        return my buildResult(false, false, 0, 0, "falha ao verificar se o PowerPoint está aberto: " & errMsg)
    end try

    set outIdx to 0
    set outTotal to 0
    set outNotes to ""
    set outTitle to ""

    tell application "Microsoft PowerPoint"
        set ssw to missing value
        try
            set ssw to slide show window 1
        on error errMsg
            return my buildResult(true, false, 0, 0, "PowerPoint aberto, mas nenhuma apresentação em modo slideshow (pressione F5). [" & errMsg & "]")
        end try

        set thePresentation to missing value
        try
            set thePresentation to presentation of ssw
        on error errMsg
            return my buildResult(true, true, 0, 0, "falha em 'presentation of slide show window': " & errMsg)
        end try

        try
            set outTotal to count of slides of thePresentation
        on error errMsg
            return my buildResult(true, true, 0, 0, "falha em 'count of slides': " & errMsg)
        end try

        set sv to missing value
        try
            set sv to slideshow view of ssw
        on error errMsg
            return my buildResult(true, true, 0, outTotal, "falha em 'slideshow view of slide show window': " & errMsg)
        end try

        set theSlide to missing value
        try
            set theSlide to slide of sv
        on error errMsg
            return my buildResult(true, true, 0, outTotal, "falha em 'slide of slideshow view': " & errMsg)
        end try

        try
            set outIdx to slide index of theSlide
        on error
            -- não crítico para as notas: seguimos com 0 (a UI mostra "?" )
            set outIdx to 0
        end try

        try
            set outNotes to my extractNotes(theSlide)
        on error errMsg
            return my buildResultFull(true, true, outIdx, outTotal, "", "", "falha ao extrair notas: " & errMsg)
        end try

        -- título é "nice to have": qualquer falha aqui não deve derrubar o
        -- resto do resultado (notas/índice já extraídos com sucesso).
        try
            set outTitle to my extractTitle(theSlide)
        on error
            set outTitle to ""
        end try
    end tell

    return my buildResultFull(true, true, outIdx, outTotal, outNotes, outTitle, "")
end run

-- Extrai o texto das notas preservando quebras de parágrafo.
--
-- `content of text range of tf` devolve o texto do quadro inteiro como uma
-- única string "achatada" (sem preservar as quebras de parágrafo internas
-- do PowerPoint) — essa era a causa da nota sair tudo junto, sem separação
-- entre frases/parágrafos. Aqui iteramos `paragraphs of (text range)`, que
-- devolve cada parágrafo separadamente, e juntamos com uma quebra de linha
-- real (ASCII 10) entre eles.
--
-- Também filtramos explicitamente `missing value`: `t is not ""` é
-- verdadeiro mesmo quando `t` é `missing value` (já que `missing value` é
-- diferente de string vazia), então um parágrafo/quadro que devolvesse
-- `missing value` acabava sendo transformado no texto literal
-- "missing value" ao juntar a lista — esse era o texto estranho que
-- aparecia no topo da nota.
on extractNotes(theSlide)
    set noteParts to {}
    tell application "Microsoft PowerPoint"
        set np to notes page of theSlide
        set shapeList to shapes of np
        repeat with s in shapeList
            try
                set tf to text frame of s
                set tr to text range of tf
                set shapeText to my extractParagraphs(tr)
                if shapeText is not "" then set end of noteParts to shapeText
            end try
        end repeat
    end tell
    set oldDelims to AppleScript's text item delimiters
    set AppleScript's text item delimiters to (ASCII character 10) & (ASCII character 10)
    set joined to noteParts as string
    set AppleScript's text item delimiters to oldDelims
    return joined
end extractNotes

-- Junta os parágrafos de um `text range`, um por linha. Se a versão do
-- Office instalada não suportar `paragraphs of`, OU se por qualquer razão
-- essa abordagem não devolver nada de útil (mesmo sem erro), cai de volta
-- para o texto "achatado" (`content of tr`) — a abordagem antiga, que já
-- sabemos que funciona. Nunca deixamos "paragraphs of" sozinho decidir se
-- a nota está vazia: só confiamos nesse resultado se ele de fato produziu
-- texto; caso contrário usamos o fallback, para nunca voltar a mostrar
-- "este slide não tem notas" para um slide que tem.
on extractParagraphs(tr)
    set paraParts to {}
    tell application "Microsoft PowerPoint"
        try
            set paraList to paragraphs of tr
            repeat with p in paraList
                try
                    set pText to my safeTextOf(content of p)
                    if pText is not "" then set end of paraParts to pText
                end try
            end repeat
        end try
    end tell

    if (count of paraParts) = 0 then
        -- "paragraphs of" não existe nesta versão do Office, ou devolveu
        -- vazio, ou todo parágrafo individual falhou ao ler — tenta o
        -- texto achatado do quadro inteiro antes de desistir.
        tell application "Microsoft PowerPoint"
            try
                set flatText to my safeTextOf(content of tr)
                if flatText is not "" then set end of paraParts to flatText
            end try
        end tell
    end if

    set oldDelims to AppleScript's text item delimiters
    set AppleScript's text item delimiters to (ASCII character 10)
    set joined to paraParts as string
    set AppleScript's text item delimiters to oldDelims
    return joined
end extractParagraphs

-- Converte com segurança um valor de texto do PowerPoint para string do
-- AppleScript: devolve "" para `missing value` (em vez de deixar a
-- coerção "as string" transformar isso no texto literal "missing value"),
-- e remove a quebra de parágrafo (CR/LF) que o PowerPoint deixa no final
-- do texto de cada parágrafo.
--
-- IMPORTANTE: esta função nunca lança erro. Uma versão anterior usava
-- `text 1 thru -2 of t` para cortar o último caractere, o que quebra
-- (erro em tempo de execução) quando `t` tem exatamente 1 caractere (ex.:
-- um parágrafo vazio que é só a marca de parágrafo, "\r") — e como esse
-- erro não estava protegido em todo lugar em que a função era chamada,
-- ele conseguia apagar a nota inteira de um slide. Agora cada corte usa
-- só índices positivos (nunca -2) e todo o laço está dentro de um `try`,
-- então na pior das hipóteses devolvemos o texto sem cortar a quebra
-- final, mas nunca perdemos o conteúdo.
on safeTextOf(v)
    if v is missing value then return ""
    try
        set t to v as string
    on error
        return ""
    end try
    try
        repeat
            set n to (count of t)
            if n = 0 then exit repeat
            set lastChar to text n thru n of t
            if lastChar is (ASCII character 13) or lastChar is (ASCII character 10) then
                if n = 1 then
                    set t to ""
                    exit repeat
                else
                    set t to text 1 thru (n - 1) of t
                end if
            else
                exit repeat
            end if
        end repeat
    end try
    return t
end safeTextOf

-- Título do slide (opcional, "nice to have"). Evitamos comparar com
-- constantes nomeadas do dicionário do PowerPoint (ex.: "placeholder type
-- title") porque um nome de constante inexistente na sua versão do Office
-- quebra a COMPILAÇÃO do script inteiro (erro -2741), e um `try` não
-- protege contra isso — só protege contra erros em tempo de execução.
-- Por isso usamos só termos genéricos já confirmados (`shapes`, `name`,
-- `text frame`, `text range`, `content`) e reconhecemos o placeholder de
-- título pelo nome do shape, que o PowerPoint atribui automaticamente
-- (ex.: "Title 1", ou "Título 1" em Office em português).
on extractTitle(theSlide)
    tell application "Microsoft PowerPoint"
        try
            set shapeList to shapes of theSlide
            repeat with s in shapeList
                try
                    set shapeName to (name of s) as string
                    if shapeName contains "Title" or shapeName contains "Título" then
                        set tf to text frame of s
                        set t to my safeTextOf(content of text range of tf)
                        if t is not "" then return t
                    end if
                end try
            end repeat
        end try
    end tell
    return ""
end extractTitle

on toBase64(theText)
    if theText is "" then return ""
    try
        return do shell script "printf %s " & quoted form of theText & " | base64 | tr -d '\n'"
    on error
        return ""
    end try
end toBase64

on buildResult(isConnected, isInShow, idx, total, errText)
    return my buildResultFull(isConnected, isInShow, idx, total, "", "", errText)
end buildResult

on buildResultFull(isConnected, isInShow, idx, total, notesText, titleText, errText)
    set notesB64 to my toBase64(notesText)
    set titleB64 to my toBase64(titleText)
    set errB64 to my toBase64(errText)
    return (isConnected as string) & "|" & (isInShow as string) & "|" & (idx as string) & "|" & (total as string) & "|" & notesB64 & "|" & titleB64 & "|" & errB64
end buildResultFull
"""


class MacConnector(BaseConnector):
    SAFETY_POLL_INTERVAL = 0.4

    def __init__(self, on_update):
        super().__init__(on_update)
        self._script_path = None

    def platform_name(self):
        return "mac"

    def start(self):
        fd, path = tempfile.mkstemp(suffix=".applescript")
        # encoding explícito: o script tem "Título" (acento) e não podemos
        # depender do encoding padrão do sistema para escrever isso certo.
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(APPLESCRIPT)
        self._script_path = path
        super().start()

    def stop(self):
        super().stop()
        if self._script_path and os.path.exists(self._script_path):
            try:
                os.remove(self._script_path)
            except OSError:
                pass

    def _run(self):
        while not self._stop_event.is_set():
            try:
                state = self._poll_once()
            except Exception as exc:
                state = SlideState(connected=False, platform="mac", error=str(exc))
            self._emit(state)
            self._stop_event.wait(self.SAFETY_POLL_INTERVAL)

    def _poll_once(self) -> SlideState:
        try:
            result = subprocess.run(
                ["osascript", self._script_path],
                capture_output=True,
                text=True,
                timeout=5,
            )
        except subprocess.TimeoutExpired:
            return SlideState(connected=False, platform="mac", error="Tempo esgotado ao consultar o PowerPoint (osascript).")
        except FileNotFoundError:
            return SlideState(connected=False, platform="mac", error="'osascript' não encontrado — isso só funciona no macOS.")

        if result.returncode != 0:
            return SlideState(connected=False, platform="mac", error=result.stderr.strip() or "Falha ao executar o AppleScript.")

        raw = result.stdout.strip()
        parts = raw.split("|")
        if len(parts) < 7:
            return SlideState(connected=False, platform="mac", error=f"Resposta inesperada do AppleScript: {raw!r}")

        connected_s, in_show_s, idx_s, total_s, notes_b64, title_b64, err_b64 = parts[:7]

        def _b64(s):
            try:
                return base64.b64decode(s).decode("utf-8", errors="replace") if s else ""
            except Exception:
                return ""

        try:
            idx = int(float(idx_s))
        except ValueError:
            idx = 0
        try:
            total = int(float(total_s))
        except ValueError:
            total = 0

        return SlideState(
            connected=(connected_s.lower() == "true"),
            in_slideshow=(in_show_s.lower() == "true"),
            slide_index=idx,
            slide_count=total,
            notes=_b64(notes_b64),
            slide_title=_b64(title_b64),
            platform="mac",
            error=_b64(err_b64) or None,
        )
