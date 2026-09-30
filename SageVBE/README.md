# SageVBE

Add-in .NET do editor do VBA (VBE), carregado dentro do Excel. Ele roda independente do VBA, então continua funcionando com o código pausado na depuração ou resetado.

- **Menu Sage**, entre "Janela" e "Ajuda", com a opção **Configurações...**.
- **Configurações** no estilo do VS Code. Por enquanto, só *Aparência: Tema de Cores*.
- **Temas para o VBE inteiro**: menus, barras de ferramentas, menus de contexto, Projeto, Propriedades, Verificação imediata, código, bordas e barras de título.
  - Temas disponíveis: *Padrão do VBE*, *Dark Modern*, *Dark+* e *Light Modern*.

## Instalar

1. Feche o Excel.
2. Rode:

   ```
   powershell -ExecutionPolicy Bypass -File install.ps1
   ```

   O script compila `src\*.cs` com o `csc` do .NET Framework 4, que já vem no Windows, e copia o DLL para `%LOCALAPPDATA%\Sage\VBE`. O registro fica só no usuário atual (HKCU) e não precisa de administrador.
3. Abra o editor do VBA (Alt+F11) e use **Sage > Configurações**.

Para remover: `install.ps1 -Uninstall`.

- **Configurações:** `%APPDATA%\Sage\settings.json`, no formato do VS Code, por exemplo `"workbench.colorTheme": "Dark Modern"`.
- **Log:** `%APPDATA%\Sage\SageVBE.log`.

## Como o tema funciona

Tudo roda na thread de interface do Excel e só age **enquanto uma janela do VBE está se pintando**: um contador é mantido pelo subclassing das janelas do VBE. Por isso o Excel não muda.

| Parte | Técnica |
|---|---|
| Cores de sistema (fundos, texto, seleção) | Desvio de `GetSysColor`/`GetSysColorBrush`/`FillRect` na tabela de importação do `VBE7.DLL`, do `VBEUI.DLL` e do `comctl32` |
| Cores do código (palavra-chave, comentário...) | `SetTextColor`/`SetBkColor` do `VBE7.DLL`: as 16 cores do "Formato do editor" viram as cores do tema (texto e fundo separados) |
| Menus e barras (desenhados pelo `VBEUI.DLL`) | Cores fixas do Office têm a luminosidade invertida (`SetTextColor`, `CreateSolidBrush`, `CreatePen`, `SetDC*Color`, `GetStockObject`) |
| Bordas 3D e botões "X" das janelas encaixadas | `DrawEdge`/`DrawFrameControl` substituídos por versões lisas |
| Abas das Propriedades e botões de modo de exibição | Desenho próprio em `Painters.cs` |
| Barras de rolagem, caixas e barra de título | Tema escuro do Windows (`DarkMode_Explorer`, `DarkMode_CFD`) e DWM |

## Limitações

- Diálogos (Opções, Referências, Localizar...) continuam no visual padrão.
- Quem usa cores personalizadas em *Ferramentas > Opções > Formato do editor* vê as cores do tema no lugar das 16 cores padrão.
- Janelas criadas por outros add-ins do VBE não são tratadas de forma especial.
