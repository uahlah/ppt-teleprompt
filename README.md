# Teleprompter para PowerPoint

Utilitário que se conecta a uma apresentação do PowerPoint **já aberta e em
execução**, identifica automaticamente qual slide está sendo exibido,
extrai o texto das notas daquele slide e projeta esse texto em uma segunda
tela no formato de teleprompter — com rolagem automática configurável —
sincronizado em tempo real com o avanço dos slides.

## Como funciona (visão geral)

```
PowerPoint (aberto, em modo de apresentação)
        │  COM (Windows) / AppleScript (Mac)
        ▼
  connector/  →  detecta slide atual + extrai notas
        │
        ▼
  server/     →  servidor local (HTTP + WebSocket em 127.0.0.1)
        │
        ├──► ui/control.html      (painel de controle — fica na tela principal)
        └──► ui/teleprompter.html (tela de teleprompter — fica na 2ª tela, em tela cheia)
```

Tudo roda localmente na sua máquina. Nenhuma nota ou conteúdo do slide sai
do seu computador.

## Instalação

**Pré-requisitos:** Python 3.9+ e o PowerPoint desktop (Windows ou macOS).
Não funciona com o PowerPoint Online / versão web, pois ele não expõe a
automação COM/AppleScript necessária.

### Windows

1. Extraia esta pasta em qualquer lugar do seu computador.
2. Dê duplo clique em `run_windows.bat`.
   - Na primeira execução, ele cria um ambiente virtual Python e instala as
     dependências automaticamente (`aiohttp`, `pywebview`, `pywin32`).
3. Abra sua apresentação no PowerPoint e inicie o modo de apresentação
   (**F5** ou **Shift+F5**).
4. O painel de controle deve abrir automaticamente. Veja a seção "Usando"
   abaixo.

### macOS

1. Extraia esta pasta em qualquer lugar do seu computador.
2. No Terminal, dê permissão de execução (só na primeira vez) e execute:
   ```bash
   chmod +x run_mac.command
   ./run_mac.command
   ```
   (ou dê duplo clique em `run_mac.command` no Finder)
3. Na primeira vez, o macOS vai pedir permissão de **Automação** para que
   este script controle o Microsoft PowerPoint. Autorize em
   *Preferências do Sistema → Privacidade e Segurança → Automação*.
4. Abra sua apresentação no PowerPoint e inicie o modo de apresentação.

> ⚠️ **Sobre o suporte a Mac**: o conector do Windows usa a API COM oficial
> do PowerPoint e é bem confiável. O conector do Mac usa o dicionário
> AppleScript do PowerPoint. As classes e propriedades usadas
> (`slide show window`, `slideshow view`, `notes page` etc.) foram
> confirmadas diretamente no dicionário do PowerPoint (Script Editor →
> File → Open Dictionary), não são só uma suposição — mas dicionários de
> AppleScript variam entre versões do Office, então ainda pode ser que
> algo precise de ajuste na sua instalação. Se falhar, cada etapa do
> script devolve uma mensagem de erro específica (ex: "falha em
> 'slideshow view of slide show window'") em vez de um erro genérico —
> essa mensagem já diz exatamente qual trecho ajustar. Veja "Ajustando o
> conector no Mac" mais abaixo.

## Instalação manual (qualquer sistema, sem os scripts .bat/.command)

```bash
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## Usando

1. Com o PowerPoint em modo de apresentação, execute o programa. Um
   **painel de controle** vai abrir na sua tela principal, mostrando:
   - o status da conexão com o PowerPoint;
   - o slide atual e uma prévia da nota;
   - a lista de monitores conectados, cada um com um botão
     **"Abrir teleprompter →"**.
2. Clique no botão do monitor onde está o teleprompter (o monitor virado
   para você, não o que a audiência vê). A tela de teleprompter abre
   automaticamente em tela cheia naquele monitor.
3. Avance os slides normalmente no PowerPoint (seta, clique, controle
   remoto de apresentação). O teleprompter atualiza a nota e volta a
   rolagem para o topo a cada troca de slide.
4. Ajuste velocidade de rolagem e tamanho da fonte pelo painel de controle
   a qualquer momento — a mudança é aplicada instantaneamente na tela do
   teleprompter.

### Atalhos de teclado (com o foco na janela do teleprompter)

| Tecla         | Ação                                   |
|---------------|-----------------------------------------|
| `Espaço`      | Pausa / retoma a rolagem automática     |
| `↑` / `↓`     | Rola manualmente (pausa o automático)   |
| `+` / `-`     | Aumenta / diminui a velocidade          |
| `M`           | Espelha o texto horizontalmente         |
| `F`           | Alterna tela cheia                      |
| `Q`           | Fecha a janela do teleprompter          |
| duplo clique  | Pausa / retoma (útil em tela touch)     |

O topo da tela do teleprompter mostra **"Slide X / Y"** em destaque e,
quando o slide tiver um título (placeholder de título nomeado "Title..."
ou "Título..." — o nome padrão que o PowerPoint atribui), o título também
aparece logo acima da nota, em destaque.

O espelhamento (`M`) é para quem usa um vidro de teleprompter de verdade
(beam-splitter) na frente da câmera — nesse caso a imagem precisa ser
espelhada para aparecer correta no reflexo. Se você só está olhando direto
para um segundo monitor, deixe desligado.

### Modo manual (sem detecção automática de monitor)

Se preferir não usar a janela nativa (ou o `pywebview` não estiver
disponível na sua máquina), rode:

```bash
python main.py --no-gui
```

Isso só liga o servidor local. Abra no navegador:
- Painel de controle: `http://127.0.0.1:8765/control.html`
- Teleprompter: `http://127.0.0.1:8765/teleprompter.html` — abra em uma
  nova janela, arraste-a para a segunda tela e pressione `F` para tela
  cheia.

### Configuração padrão

Copie `config.example.json` para `config.json` para definir os valores
iniciais (velocidade, tamanho de fonte, tema, espelhamento) sem precisar
ajustar tudo de novo a cada execução:

```bash
cp config.example.json config.json
```

## Estrutura do projeto

```
ppt_teleprompter/
  main.py                  orquestrador: sobe servidor, conector e painel
  requirements.txt
  config.example.json
  connector/
    base.py                interface comum (SlideState, BaseConnector)
    windows_connector.py   automação COM (pywin32) — eventos + polling de segurança
    mac_connector.py       automação AppleScript (osascript) — polling
  server/
    app.py                 servidor local aiohttp (HTTP + WebSocket)
  ui/
    teleprompter.html      tela de teleprompter (2ª tela)
    control.html           painel de controle (tela principal)
  run_windows.bat
  run_mac.command
```

## Ajustando o conector no Mac

Se ao abrir o painel de controle o status ficar em erro mesmo com o
PowerPoint em modo de apresentação, o motivo mais provável é um nome de
propriedade do AppleScript diferente do esperado. Para conferir e corrigir:

1. Veja a mensagem de erro exibida no painel de controle ou na tela do
   teleprompter — ela já indica qual etapa falhou (ex: "falha em 'count
   of slides'", "falha ao extrair notas: ...").
2. Abra o app **Script Editor** (Editor de Script) no Mac e confirme o
   nome certo da propriedade: menu **File → Open Dictionary… → Microsoft
   PowerPoint**, procure a classe mencionada no erro (`slide show
   window`, `slideshow view`, `notes page`, `text frame`...) e veja suas
   propriedades reais na sua versão do Office.
3. Abra `connector/mac_connector.py` e ajuste só o trecho correspondente
   dentro da variável `APPLESCRIPT` — cada etapa tem seu próprio bloco
   `try`/`on error`, então a correção fica isolada, sem afetar as outras.
4. Para testar rapidamente sem precisar reiniciar o app inteiro, copie só
   o miolo do bloco `tell application "Microsoft PowerPoint" ... end
   tell` do `on run`, cole no Script Editor e aperte "Run" enquanto a
   apresentação estiver ativa — o resultado/erro aparece no painel
   inferior do Script Editor.

## Limitações conhecidas

- Não funciona com o PowerPoint Online (versão web) nem com o PowerPoint
  do iPad/Android — apenas a versão desktop (Windows/macOS), que expõe a
  automação necessária.
- No modo "Apresentar Online" ou em alguns modos de apresentação em rede,
  o comportamento não foi validado.
- A detecção e o posicionamento automático de monitor usam só
  `webview.screens` (de propósito uma única biblioteca, para evitar
  incompatibilidade entre sistemas de coordenadas). Ainda assim, o
  comportamento exato de "tela cheia num monitor específico" pode variar
  um pouco entre versões do macOS/Windows. Se o teleprompter não abrir no
  monitor certo, a janela tem barra de título normal (com botão de
  fechar) e pode ser arrastada manualmente para a tela certa — F alterna
  tela cheia, Q fecha a janela.
- Este projeto não foi testado contra uma instância real do PowerPoint
  durante o desenvolvimento (o ambiente onde foi construído não tem
  Windows/PowerPoint disponível). A lógica foi revisada com cuidado e
  testada isoladamente onde possível (parsing, máquina de estados,
  detecção de mudança de slide), mas o primeiro teste "de verdade" com o
  PowerPoint aberto será o seu. Se algo não funcionar como esperado, o
  ponto de partida mais provável para o ajuste é `_poll_once()` em
  `windows_connector.py` ou a string `APPLESCRIPT` em `mac_connector.py`.

## Privacidade

Todo o tráfego fica em `127.0.0.1` (localhost) — o servidor não expõe
nada para a rede local nem para a internet. O conteúdo das notas nunca sai
da sua máquina.
