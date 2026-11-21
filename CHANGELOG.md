# Histórico de versões do Sage Framework

O que entrou em cada versão e o que está planejado. A versão segue `maior.menor.correção`: novidades aumentam o número do meio; só correções, o último.

## Planejado

### Controles para UserForms
- **TreeView**: mostra um JSON (ou um `Dictionary`/`List`) como árvore, para navegar por dados hierárquicos.
- **ComboBox**: a caixa de seleção do Sage, no mesmo estilo do TextBox e do DateBox.

## Ideias para o SageEditor

Ainda sem versão.

### Aba Problemas: próximos passos
A aba existe desde a 1.6.0, com os erros de estrutura (encontrados lendo o texto, sem o compilador). Próximos passos, pela mesma análise própria (nada de compilar cópias):

- **Variável não definida** (com `Option Explicit`) e **Sub ou Function não definida**: lendo também os nomes das bibliotecas referenciadas (as bibliotecas de tipos delas), para não dar alarmes falsos.
- **Aviso** (o código roda, mas provavelmente está errado): módulo sem `Option Explicit`; `Dim a, b As Long` (só o `b` é `Long`); variável ou parâmetro nunca usado; procedimento privado nunca chamado; `On Error Resume Next` sem `On Error GoTo 0`; nomes de antes da 1.5 (`Sage.StringS`...).
- **Informação:** `TODO`/`FIXME`; `Dim x` ou `Function F()` sem tipo; parâmetro sem `ByVal`/`ByRef`; `Scripting.FileSystemObject` e `Scripting.Dictionary`, com a sugestão do `Sage.Path` e do `Sage.Dictionary`.

**Regra principal: não engessar o código.** A aba só informa; nada bloqueia, obriga ou muda o jeito de escrever:
- nada é sublinhado nem pisca enquanto se digita; os avisos ficam só na aba, atualizados alguns instantes depois de parar de digitar, e só no módulo que mudou;
- nenhum aviso impede rodar, compilar ou salvar, e nenhum muda o código sozinho;
- cada regra pode ser desligada nas Configurações, e uma linha pode ser ignorada com o comentário `'@Ignore`;
- a aba pode ficar fechada: quem não quiser, não a vê;
- o que pode ter outra explicação (procedimento "nunca chamado", que pode ser chamado por um botão da planilha ou por `Application.Run`) entra como Informação, e não como Aviso.

## Ideias para o SageTypes

Ainda sem versão.

### Funções de planilha
Funções novas para usar direto nas células, como `=ENTRE(A1; 1; 10)` (VERDADEIRO se o valor está entre os dois limites).

- **Como:** um suplemento de automação COM, em C#, no mesmo pipeline de registro do SageTypes; o instalador o ativa para cada usuário. Uma XLL (com o Excel-DNA) só se um dia precisarmos de descrição dos argumentos no assistente de função ou de mais velocidade.
- **Nomes nos dois idiomas:** o Excel traduz sozinho só as funções nativas (o arquivo as guarda em inglês, e cada Excel mostra no seu idioma de exibição). Uma função personalizada fica com o nome digitado. Por isso cada função existe com os dois nomes, `ENTRE` e `BETWEEN`, e uma planilha funciona em qualquer Excel com o Sage. Não registrar só o nome do idioma do Office (a planilha de um colega de outro idioma abriria com `#NOME?`) nem reescrever as fórmulas ao abrir o arquivo.
- **ENTRE:** inclui as pontas (como o `BETWEEN` do SQL) e aceita os limites em qualquer ordem; um quarto argumento opcional exclui as pontas.
- **Sem o Sage:** a planilha mostra `#NOME?`, como com qualquer suplemento.

## Ideias: o que é antigo no VBA e pode ser modernizado

Ainda sem versão. Coisas que os UserForms não têm, ou têm de um jeito antigo.

### Sem equivalente no VBA
- **Chart**: gráfico de linha, barras ou pizza direto de um `DataFrame`, dentro do form.
- **ProgressBar**: barra de progresso para macros longas, sem improvisar com um Label.
- **Toast**: aviso que aparece e some sozinho ("Salvo!"), sem o `MsgBox` que trava a macro.
- **DataGrid editável**: editar células, acrescentar e apagar linhas e receber o `DataFrame` alterado.
- **Image moderna**: PNG com transparência, SVG e imagens da web (o Image do VBA só aceita BMP, JPG e GIF).

### Entradas
- **Toggle**: interruptor liga/desliga, no lugar do CheckBox.
- **NumberBox**: número com botões + e −, mínimo, máximo, passo e formato.
- **Slider**: escolher um valor ou um intervalo arrastando (o ScrollBar do VBA não serve para isso).
- **MultiSelect**: escolher vários itens, mostrados como etiquetas com ×, no lugar da ListBox com caixas.
- **FileBox**: arrastar e soltar um arquivo no form, ou escolher pelo botão, e receber o caminho.
- **ColorBox**: escolher uma cor e receber o `RGB`.

### Organização e exibição
- **Metric**: cartão de indicador, com valor grande, rótulo e variação colorida (▲ +12%).
- **Expander**: seção que recolhe e abre, para forms longos.
- **Tabs**: abas no visual atual (o MultiPage do VBA tem visual antigo).
- **Validação**: ícone de erro com a mensagem ao lado do controle inválido.

## 1.6.0

### SageEditor
- Aba **Problemas** no Painel (a primeira da faixa): lista de uma vez os erros de sintaxe e de estrutura do código de todos os projetos abertos, lendo o texto, sem compilar e sem mexer no código: as linhas que o VBE marca em vermelho (sobra no cabeçalho do procedimento, como `Sub Teste 1()`, ou numa declaração, parênteses desencontrados, `If` sem `Then`, `For` sem `=` ou `To`, `x = a b`), blocos desencontrados (`If`/`End If`, `For`/`Next`, `Do`/`Loop`, `With`, `Select Case`, `Sub`/`Function`/`Property`...), `Next` com a variável de outro `For`, `Case`/`Else`/`Exit` fora do bloco, rótulo não definido, declaração duplicada, nome ambíguo e código fora de procedimento, com as mensagens do VBA e sem erros em cascata. Verifica ao abrir a aba, ao salvar e pelo botão Verificar; o clique leva ao trecho.
- **Erros no Terminal**: um erro do código (em tempo de execução ou de compilação) não abre mais a caixa do VBA; vai para a aba Terminal: primeiro o lugar, como o caminho na janela Projeto (`Pasta1.xlsm\Módulos\Modulo1\Calcular, linha 12`), depois o código da linha e, por último, a mensagem; o cursor fica na linha do erro. *Editor: Erros no Terminal* nas Configurações: *Mostrar e Parar* (padrão), *Mostrar e Depurar* ou *Desativado*. Com o editor do VBA fechado, fica a caixa de sempre.
- Configurações: o índice à esquerda é uma árvore, com **Editor > Aparência** (tema, números de linha, abas, ícones, decoradores de cor, Explorador) e **Editor > Recursos** (vários cursores, erros no Terminal, referências, controles da Caixa de ferramentas). O clique em *Editor* recolhe ou abre o grupo.
- Menu **Sage > Exibir**: abre e fecha o Explorador (Ctrl+R) e o Painel (Ctrl+J). O que está à vista fica marcado.
- Vários cursores: **Ctrl+V** (ou Shift+Insert) cola em todos os cursores, no lugar da seleção de cada um (antes, só no principal). Com uma linha na área de transferência para cada cursor, cada um recebe a sua; senão, todos recebem o texto inteiro.
- Correção: com vários cursores (Ctrl+F2, Alt+Clique), as seleções extras às vezes apareciam deslocadas ou largas demais. A largura do caractere agora vem da fonte do código, e não do cursor do VBE. Os cursores extras ficam com a largura do cursor do Windows, a mesma do cursor do VBE (1 px no padrão, antes 2 px), e piscam junto com ele: no mesmo ritmo, recomeçando visíveis a cada movimento.
- O tema *Dark+* passa a se chamar **Dark** (quem já usa continua nele).
- Aba Terminal: a tela é uma caixa só, como num terminal comum, com o prompt e a linha de comando na última linha. Dá para selecionar e copiar qualquer trecho, inclusive da saída até a linha de comando; só a linha de comando pode ser editada. O clique direito tem *Copiar*, *Copiar Tudo*, *Colar* e *Selecionar Tudo*, e o Ctrl+A seleciona tudo.
- Aba Terminal: o prompt sai na cor do texto do tema, e não mais na cor de destaque (azul).
- O Painel (Ctrl+J, *Sage > Exibir > Painel*) abre sempre na aba **Terminal**.
- Correção: a aba Terminal às vezes abria sem o prompt e sem a linha de comando, e só aparecia depois de trocar de aba e voltar.
- A biblioteca aparece em *Ferramentas > Referências* como **Sage Framework 1.6**.

## 1.5.0

### SageEditor
- Layout fixo, como no VS Code, com o Explorador unificado ligado:
  - O Explorador ocupa a altura inteira, à esquerda ou à direita, e o Painel (Ctrl+J) fica sempre embaixo do código, ao lado dele. As janelas não se desacoplam mais: uma janela solta volta a acoplar.
  - Clique direito na barra do Explorador: **Mover a Barra Lateral Primária para a Esquerda/Direita** e **Ocultar Barra Lateral Primária** (Ctrl+R).
  - Clique direito nas abas do Painel: **Ocultar Painel** (Ctrl+J).
  - Os divisores continuam ajustando os tamanhos, que ficam salvos nas configurações.
  - O item *Encaixável* sai dos menus do VBE enquanto o Sage está carregado.
- Propriedades do Explorador: a borda de cima da faixa mostra a seta dupla e arrasta para mudar a altura.
- A biblioteca aparece em *Ferramentas > Referências* como **Sage Framework 1.5**.

### SageTypes
- **Console.PrintF**: escreve na aba Terminal do Painel, com os marcadores do `printf` do C (`%d`, `%s`, `%.2f`, `%x`...). Cada chamada termina a linha, como um print (sem linha em branco se o texto já terminar com `\n`).
- **Path**: arquivos e pastas sem a biblioteca Scripting (`FileSystemObject`). Ler e gravar texto (`ReadText`, `ReadLines`, `WriteText`, `WriteLines`, `AppendText`, com UTF-8 e detecção de BOM/ANSI), partes do caminho (`Name`, `Stem`, `Extension`, `Parent`, `Join`, `WithName`, `WithExtension`), informações (`Exists`, `IsFile`, `IsFolder`, `Size`, `Modified`, `Created`), listagem (`Files`, `Folders`, `Glob` com `**`), `Mkdir`, `CopyTo`, `MoveTo` e `Delete`. Um caminho sem pasta é procurado na pasta da pasta de trabalho ativa, como no `DataFrame`.
- Global **PathOf**: um `Path` numa linha, de um texto ou de uma pasta do Windows (`sgDesktop`, `sgDocuments`, `sgDownloads`, `sgTemp`...).
- **Datetime.StrFTime** aceita os códigos do `Format` do VBA (`"dd/mm/yyyy hh:nn"`), pelo mesmo motor e com o mesmo resultado, mais `f` para as frações de segundo (`"hh:mm:ss.fff"`). Com `%` no formato, continuam os códigos do strftime (`"%d/%m/%Y"`).
- **Nomes novos**: `StringS`, `ListS`, `DictionaryS` e `DatetimeS` passam a se chamar **`String`**, **`List`**, **`Dictionary`** e **`Datetime`** (`Sage.String`, `CreateObject("Sage.List")`...). No código antigo, troque os nomes com *Editar > Substituir* do VBE, em *Projeto atual*.

## 1.4.0

### SageEditor
- Temas novos: **Azure (Dark)**, inspirado no Dark+ com o azul do Azure; **Azure (Light)**, branco e azul-claro, inspirado no portal do Azure; e **Sage (Light)**. O tema *Sage* passa a se chamar **Sage (Dark)** (quem já usa continua nele). O *Padrão do VBE* passa a se chamar **Light (Padrão do VBA)**, e a lista fica em ordem alfabética. Escolha em *Sage > Configurações > Aparência: Tema de Cores*.
- Explorador unificado:
  - Botões **Novo Formulário**, **Novo Módulo** e **Nova Classe** na linha do projeto ativo, à direita, como os de Novo Arquivo do VS Code: os ícones do Sage com um "+". O clique cria o componente e já o abre.
  - Linhas mais altas, como no VS Code (22 px na escala de 100%), com os ícones no mesmo tamanho dos do VS Code.
  - A barra de título passa a mostrar **Explorador** (*Explorer* em inglês) no lugar de *Projeto - VBAProject*, porque a janela tem a árvore e as Propriedades.
- Código: a roda do mouse rola além da última linha, como no VS Code, até ela ficar no topo da janela, com espaço vazio embaixo. Redimensionar a janela não tira o código do lugar, e o clique no espaço vazio põe o cursor no fim do código.
- Janela Terminal (Ctrl+J): ao lado do ×, como no VS Code:
  - Aba Terminal: lixeira **Fechar Terminal**, o mesmo que digitar `exit`.
  - Aba Resultado do DataFrame: **Limpar Resultado** e **Salvar**, que grava todas as linhas do DataFrame mostrado em **CSV**, **Excel** (.xlsx), **JSON**, **Markdown** ou **XML**. No Excel, no JSON e no XML, números, datas e sim/não continuam com o tipo; o CSV usa o separador do Windows (`;` no português), como o `df.ToCsv`.
- A biblioteca aparece em *Ferramentas > Referências* como **Sage Framework 1.4**.

### Correções
- Código: as caixas de objeto e de procedimento, os botões de modo de exibição e as barras de rolagem piscavam ao rolar, porque a cada rolagem a janela inteira era redesenhada. Agora só o texto é.
- Explorador: depois de *Alternar pastas*, as pastas voltavam com o ícone comum do VBE até serem expandidas. Agora cada uma já aparece com o seu ícone.
- Explorador: depois do **Ctrl+R**, a barra de título voltava a ser a do VBE.
- SageShortcuts: os atalhos novos de cada versão (como o **Ctrl+R**, que abre e fecha a janela Projeto) não chegavam a quem já tinha instalado antes, porque o `%APPDATA%\Sage\keybindings.txt` só era criado na primeira vez. Agora os atalhos novos entram no fim do arquivo, uma vez cada: um atalho apagado não volta, e um atalho que você já usa para outra ação fica como está.
- Números de linha: com a fonte do editor maior (*Ferramentas > Opções > Formato do editor*), os números ficavam com o espaçamento da fonte anterior e se desalinhavam das linhas do código. Agora eles acompanham o tamanho da fonte, e a faixa se alarga para caber os números maiores.

## 1.3.0

### SageEditor
- **Decoradores de cor**: cada `RGB(255, 128, 0)` e cada hexadecimal do VBA (`&H80FF&`, `&H8000000F`) fica sublinhado na própria cor. As cores do sistema aparecem com a cor real do Windows.
  - O clique na cor abre um seletor de cor (saturação e brilho, matiz, `#RRGGBB`, R, G, B) que reescreve a cor no código, no mesmo formato, enquanto você escolhe.
  - Enter ou clicar fora mantém; Esc volta a cor original.
  - Opção *Editor: Decoradores de Cor* nas Configurações (`"editor.colorDecorators"`).
- A aba *Resultado DataFrame* da janela Terminal passa a se chamar **Resultado do DataFrame** em português.
- Aba Terminal: `exit` fecha a janela e recomeça o PowerShell, com a tela, o histórico de comandos e as variáveis apagados.
- A biblioteca aparece em *Ferramentas > Referências* como **Sage Framework 1.3**.

### Instalador
- Nova opção **Permitir os controles do Sage nos UserForms (ActiveX do Office)**, marcada por padrão. O Microsoft 365 passou a vir com *Desabilitar todos os controles sem notificação*, que bloqueia o TextBox, o DataGrid e o DateBox com o erro "Não há confiabilidade no requerente para a ação especificada". Com a opção marcada, cada usuário do computador passa para *Avisar-me antes de habilitar controles UFI com restrições adicionais e controles SFI com restrições mínimas*, e os controles do Sage abrem sem aviso.
  - A desinstalação volta a configuração anterior.
  - Se uma política de grupo bloquear o ActiveX, o instalador avisa (a política tem prioridade).
  - Instalação sem janela: `/quiet /noactivex` não mexe na configuração.

### Correções
- Aba Terminal: a linha de comando atual tinha letras um pouco maiores e ficava deslocada em relação aos comandos já executados. Agora ela fica alinhada como a próxima linha do histórico.

## 1.2.0

### SageTypes
- **Sage DateBox**: caixa de data para UserForms, com calendário.
  - `Value` devolve um `Date`, ou `Empty` quando está vazia.
  - Digitação rápida: `12102026`, `121026`, `1503` (ano atual), `5` (mês atual), com o separador entrando sozinho. Data inválida ou fora dos limites é recusada.
  - Calendário pelo botão, F4 ou Alt+Seta para baixo, sem tirar o foco do form: dias, meses e anos, *Hoje* e *Limpar*, navegação pelo teclado.
  - `Format` (formatos do VBA), `MinDate`, `MaxDate`, `Placeholder`, cores, fonte e `DropDown()`.
  - Eventos `Change` (só quando a data muda), `DropButtonClick`, `KeyDown` e `KeyPress`.
- A biblioteca aparece em *Ferramentas > Referências* como **Sage Framework 1.2**.

### SageEditor e SageShortcuts
- **Página "Sage" na Caixa de ferramentas**: os controles do Sage ficam numa página própria, ao lado da página *Controles*. Quem tinha os controles na página *Controles* (versões anteriores) os vê passar para a página Sage.
- O DateBox entra sozinho na página Sage e na opção *Controles na Caixa de Ferramentas* das Configurações.
- A desinstalação apaga a página Sage e tira os controles do Sage da Caixa de ferramentas de todos os usuários do computador.

### SageEditor
- Com as abas, as janelas de código e de formulário ocupam a área toda, sem barra de título e sem os botões nativos de minimizar, restaurar e fechar.

## 1.1.0

### SageTypes
- **Sage TextBox**: caixa de texto para UserForms com correção ortográfica do Windows (sublinhado, sugestões, *Ignorar*, *Adicionar ao dicionário*), `Placeholder`, `Language`, `Misspellings()` e borda moderna.
- **Sage DataGrid**: mostra um `DataFrame` num UserForm, com milhões de linhas.
  - Paginação automática (`AutoPaging`), com a barra *linhas 1–20 de N ◀ Página 1 de N ▶*.
  - Formatos de coluna com os formatos do VBA (`ColumnFormats`).
  - Seleção de células, linhas e colunas com mouse e teclado, e Ctrl+C pronto para colar no Excel.
  - Cor da seleção (`SelectionColor`); valores nulos como "Vazio"/"Empty".
- **`Resizable`**: o UserForm ganha borda para redimensionar, com os controles ancorados (`Fill`, `Width`, `Bottom Right`...).
- **`OfficeTheme()`**: o tema do Office em uso (`"Colorful"`, `"Dark Gray"`, `"Black"`, `"White"`), resolvendo *Usar configuração do sistema* para o tema real.
- A biblioteca aparece em *Ferramentas > Referências* como **Sage Framework 1.1**; arquivos com a referência à 1.0 passam a usar a 1.1 sozinhos.

### SageEditor
- Configurações novas: *Referência no Projeto Ativo* (adiciona ou remove a referência do Sage no projeto aberto) e *Controles na Caixa de Ferramentas*.
- Aba Resultado DataFrame: valores nulos como "Vazio" em português e "Empty" nos outros idiomas.

### SageShortcuts e instalador
- Os controles do Sage entram sozinhos na Caixa de ferramentas dos UserForms (uma vez por usuário).
- O instalador apaga os caches dos controles no MSForms (`Sage*.exd`) ao instalar e desinstalar.

### Correções
- Com os temas escuros do Office, os itens dos menus do editor do VBA apareciam cinza, como se estivessem desabilitados.
- *Controles adicionais* da Caixa de ferramentas não abria depois que um controle era tirado da lista.

## 1.0.0 (2026-10-10)

Primeira versão, com um instalador único para Windows 10 e 11 com Excel de 64 bits.

### SageEditor (add-in do editor do VBA)
- Menu **Sage** e tela de **Configurações**, com o arquivo `%APPDATA%\Sage\settings.json`.
- Temas para o editor inteiro: *Padrão do VBE*, *Dark Modern*, *Dark+*, *Sage* e *Light Modern*.
- Realce de sintaxe ampliado (procedimentos, variáveis, strings, números, tipos, controle de fluxo).
- Números de linha, com a linha atual em destaque.
- Abas das janelas abertas, com arrastar, fechar e menu de contexto.
- Vários cursores (Alt+Clique, Ctrl+Alt+Seta, Shift+Alt+arrastar, Ctrl+F2).
- Explorador unificado: as Propriedades dentro da janela Projeto.
- Ícones coloridos na janela Projeto e nas abas.
- Janela **Terminal**: Verificação imediata, Variáveis locais, Inspeção de variáveis, um PowerShell que executa o VBA e a aba Resultado DataFrame.
- Comando `Clear` na Verificação imediata.
- Referência automática à biblioteca Sage em toda pasta de trabalho nova.
- Interface em português ou inglês, conforme o idioma do Office.

### SageShortcuts
- Atalhos configuráveis em `keybindings.txt`: Ctrl+K, Ctrl+C / Ctrl+K, Ctrl+U para comentar, Alt+Seta para mover linhas, Shift+Alt+Seta para duplicar, Ctrl+J para o Terminal, Ctrl+R para a janela Projeto, F12 para ir à definição.

### SageTypes (biblioteca Sage)
- `StringS`, `ListS`, `DictionaryS` e `DatetimeS`: métodos encadeáveis, índices negativos, fatias e erros com nome (`KeyError`, `IndexError`, `TypeError`).
- `DataFrame` sobre o DuckDB, para milhões de linhas.
- `Json`, `Requests` (com `Response` e `Session`) e `SqlEngine` (PostgreSQL, MySQL, SQLite, SQL Server, Access, ODBC).
- Globais `GetUser` e `Sleep`.
