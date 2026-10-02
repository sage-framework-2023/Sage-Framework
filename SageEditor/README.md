# SageEditor

Add-in .NET do editor do VBA (VBE), carregado dentro do Excel. Ele roda independente do VBA, então continua funcionando com o código pausado na depuração ou resetado.

- **Menu Sage**, entre "Janela" e "Ajuda", com a opção **Configurações...**.
- **Configurações** no estilo do VS Code. Por enquanto, só *Aparência: Tema de Cores*.
- **Temas para o VBE inteiro**: menus, barras de ferramentas, menus de contexto, Projeto, Propriedades, Verificação imediata, código, bordas e barras de título, a Caixa de ferramentas e o fundo do designer de UserForms (o formulário em si mantém as cores dele).
  - Temas disponíveis: *Padrão do VBE*, *Dark Modern*, *Dark+*, *Sage* (escuro em tons de sálvia) e *Light Modern*.
- **Realce de sintaxe** no estilo do VS Code, além das cores que o VBE já tem:
  - nomes de Sub/Function/Property/Enum e chamadas de procedimentos em amarelo;
  - variáveis e propriedades em ciano;
  - strings em laranja e números em verde;
  - tipos (`As Long`, `As Sage.ListS`) em verde-água;
  - controle de fluxo (`If`, `For`, `Select Case`, `Exit`...) em roxo.
- **Números de linha** à esquerda do código, com a linha atual em destaque (*Editor: Números de Linha* nas Configurações).
- **Abas** das janelas abertas (módulos, classes e formulários) no topo da área de código, como no VS Code (*Editor: Mostrar Abas* nas Configurações). Funciona com qualquer tema, inclusive o *Padrão do VBE*:
  - clique ativa a janela; arrastar muda a aba de lugar; o `×` ou o botão do meio fecha;
  - o botão direito abre *Fechar*, *Fechar Outras*, *Fechar à Direita* e *Fechar Todas*;
  - formulários mostram o tipo ao lado do nome (`frmPrincipal  UserForm`); módulos com o mesmo nome em projetos diferentes mostram o projeto.

- **Vários cursores**, como no VS Code (*Editor: Vários Cursores* nas Configurações):
  - **Alt+Clique** acrescenta um cursor (ou tira, se já houver um ali);
  - **Ctrl+Alt+Seta para Cima/Baixo** acrescenta um cursor na linha de cima/de baixo;
  - **Shift+Alt+arrastar** seleciona em coluna (um cursor por linha);
  - digitar, Backspace, Delete, Tab, setas (Shift+Seta seleciona na linha), Home e End valem para todos; **Esc** volta a um cursor; outras teclas (Enter, Ctrl+...) e cliques sem Alt também voltam a um cursor.
  - As edições são feitas pelo `CodeModule`: o Ctrl+Z do VBE não as desfaz, e o VBE arruma a linha enquanto você digita (`x=1` vira `x = 1`), não só ao sair dela. Precisa de fonte de largura fixa no editor, como a padrão.
- **Comandos na Verificação imediata**, como num terminal: digite `Clear` e Enter para apagar tudo, ou `Exit` e Enter para fechar a janela (o histórico continua lá ao reabrir). Sem diferenciar maiúsculas. Como o VBE não dá acesso ao texto da janela, o add-in acompanha o que é digitado desde o começo da linha; se o cursor for movido (setas, clique) antes do Enter, a linha segue normalmente para o VBE. (`Exit` é palavra reservada do VBA: não daria para ser uma função.) O Ctrl+J que abre e fecha a janela fica no SageShortcuts.
- **Referência automática** à biblioteca Sage Framework (`Sage.StringS`, do SageTypes): marcada em *Ferramentas > Referências* do projeto cujo código você abre (*Editor: Referência Automática* nas Configurações). Projetos protegidos, em execução ou depuração, ou que usam o `Sage.xlam` ficam como estão. Como qualquer mudança no projeto, a pasta de trabalho passa a pedir para ser salva.
- **Idioma** da interface do Sage igual ao do Office/Excel: português ou inglês (os demais idiomas usam inglês). Os textos ficam em `src\Strings.cs`; para outro idioma, basta um método com a tradução.

## Instalar

1. Feche o Excel.
2. Rode:

   ```
   powershell -ExecutionPolicy Bypass -File install.ps1
   ```

   O script compila `src\*.cs` com o `csc` do .NET Framework 4, que já vem no Windows, e copia o DLL para `%LOCALAPPDATA%\Sage\Editor`. O registro fica só no usuário atual (HKCU) e não precisa de administrador.
3. Abra o editor do VBA (Alt+F11) e use **Sage > Configurações**.

Para remover: `install.ps1 -Uninstall`.

- **Configurações:** `%APPDATA%\Sage\settings.json`, no formato do VS Code, por exemplo `"workbench.colorTheme": "Dark Modern"`, `"editor.lineNumbers": "off"` ou `"workbench.editor.showTabs": "none"`. Para forçar um idioma diferente do Office: `"locale": "en"` (ou `"pt-BR"`).
- **Log:** `%APPDATA%\Sage\SageEditor.log`.

## Como o tema funciona

Tudo roda na thread de interface do Excel e só age **enquanto uma janela do VBE está se pintando**: um contador é mantido pelo subclassing das janelas do VBE. Por isso o Excel não muda.

| Parte | Técnica |
|---|---|
| Cores de sistema (fundos, texto, seleção) | Desvio de `GetSysColor`/`GetSysColorBrush`/`FillRect` na tabela de importação do `VBE7.DLL`, do `VBEUI.DLL` e do `comctl32` |
| Cores do código (palavra-chave, comentário...) | `SetTextColor`/`SetBkColor` do `VBE7.DLL`: as 16 cores do "Formato do editor" viram as cores do tema (texto e fundo separados) |
| Menus e barras (desenhados pelo `VBEUI.DLL`) | Cores fixas do Office têm a luminosidade invertida (`SetTextColor`, `CreateSolidBrush`, `CreatePen`, `SetDC*Color`, `GetStockObject`) |
| Bordas 3D e botões "X" das janelas encaixadas | `DrawEdge`/`DrawFrameControl` substituídos por versões lisas |
| Abas das Propriedades e botões de modo de exibição | Desenho próprio em `Painters.cs` |
| Números de linha | Faixa reservada na área não-cliente da janela de código (`WM_NCCALCSIZE`), onde os números são desenhados (`WM_NCPAINT`). O VBE continua cuidando de clique, cursor e rolagem no espaço restante. A altura das linhas vem do desenho do código; a primeira linha visível e a linha atual vêm do `CodePane`, associado à janela pelo título |
| Abas | Faixa reservada no topo da área não-cliente do `MDIClient`, o contêiner das janelas de código; as janelas, maximizadas ou não, ficam no espaço que sobra. A lista vem dos filhos do `MDIClient` (Win32, sem COM); os cliques chegam como `WM_NC*BUTTON*` (o hit-test da faixa responde `HTBORDER`) e viram `WM_MDIACTIVATE` ou `SC_CLOSE`. O arrasto usa um loop próprio que acompanha o cursor (`GetCursorPos`), porque o loop de mensagens do VBE não entrega `WM_MOUSEMOVE` com a captura ativa. Subclassing próprio (`EditorTabs.cs`), independente do tema |
| Realce de sintaxe | `ExtTextOutA`/`TextOutA` do `VBE7.DLL`: cada trecho de texto normal ou de palavra-chave é dividido em tokens (`Syntax.cs`) e redesenhado em pedaços. O VBE usa `TA_UPDATECP`, então os pedaços saem em sequência. Os nomes de procedimentos vêm dos módulos não protegidos e das declarações que aparecem na tela |
| Barras de rolagem, caixas e barra de título | Tema escuro do Windows (`DarkMode_Explorer`, `DarkMode_CFD`) e DWM |
| Caixa de ferramentas | Desenhada pelo FM20.DLL (Microsoft Forms), que não lê as cores pelo `GetSysColor` importado. Depois de cada desenho, os cinzas têm a luminosidade invertida (`Painters.InvertGrays`): branco vira o fundo do tema, preto vira o texto e cores (ícones, seleção) ficam. Ela é criada sem dono, então é encontrada pela verificação periódica e não pelo hook CBT. Os UserForms não são tocados |
| Fundo do designer | `WM_ERASEBKGND` do `DesignerWindow` com a cor do editor |
| Vários cursores | Subclassing próprio das janelas de código (`MultiCursor.cs`): com mais de um cursor, as teclas viram edições pelo `CodeModule.ReplaceLine`, e a coluna de cada cursor é recalculada pelos caracteres que não são espaço (o VBE reformata a linha). Os cursores extras são desenhados com XOR no `WM_PAINT`, só na área repintada. A posição na tela vem de uma referência medida com o cursor do VBE sem seleção (com seleção ele o esconde) mais a largura do caractere e a altura da linha. O Shift+Alt+arrastar usa um laço próprio, como o arraste das abas |

## Limitações

- Diálogos (Opções, Referências, Localizar...) continuam no visual padrão.
- Quem usa cores personalizadas em *Ferramentas > Opções > Formato do editor* vê as cores do tema no lugar das 16 cores padrão.
- Janelas criadas por outros add-ins do VBE não são tratadas de forma especial.
