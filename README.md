# Sage Framework

**O VBA, modernizado, sem sair do Excel.** O Sage Framework traz para o editor do VBA a experiência de um IDE atual e, para a linguagem, tipos modernos para texto, listas, datas, tabelas, JSON, APIs, bancos de dados e controles novos para UserForms. Tudo num instalador só.

📖 **Documentação:** https://sageframework.onrender.com
⬇️ **Download:** [Releases](https://github.com/sage-framework-2023/Sage-Framework/releases) (`SageFramework-Setup-x.y.z.exe`)
📝 **Novidades:** [CHANGELOG.md](CHANGELOG.md)

## O que vem no pacote

### SageEditor: o editor do VBA renovado
- **Temas** para o editor inteiro: *Dark Modern*, *Dark+*, *Sage*, *Light Modern* ou o padrão do VBE.
- **Realce de sintaxe** mais completo, **números de linha**, **abas** das janelas abertas e **vários cursores**.
- **Explorador unificado** (as Propriedades dentro da janela Projeto) e **ícones** para cada tipo de componente.
- **Janela Terminal**: Verificação imediata, Variáveis locais, Inspeção de variáveis, um PowerShell que executa o seu VBA e a aba *Resultado DataFrame*.
- **Referência automática** à biblioteca Sage em toda pasta de trabalho nova.
- Interface em português ou inglês, conforme o idioma do Office.

### SageShortcuts: atalhos de teclado
Configuráveis num arquivo de texto: `Ctrl+K, Ctrl+C` para comentar, `Alt+Seta` para mover linhas, `Shift+Alt+Seta` para duplicar, `Ctrl+J` para o Terminal, `F12` para ir à definição e outros.

### SageTypes: a biblioteca **Sage**
| Tipo | Para quê |
|---|---|
| `StringS`, `ListS`, `DictionaryS` | Texto, listas e dicionários com métodos encadeáveis, índices negativos, fatias e erros claros (`KeyError`, `IndexError`) |
| `DatetimeS` | Data e hora com fuso, dias úteis, início e fim de período |
| `DataFrame` | Tabelas sobre o DuckDB, para milhões de linhas |
| `Json`, `Requests`, `SqlEngine` | Ler e gravar JSON, chamar APIs e conectar a PostgreSQL, MySQL, SQLite, SQL Server, Access e ODBC |
| `Resizable`, `OfficeTheme` | UserForms redimensionáveis e o tema do Office em uso |

### Controles para UserForms
Na página **Sage** da Caixa de ferramentas:
- **TextBox** com correção ortográfica do Windows e texto de dica.
- **DataGrid** que mostra um `DataFrame` com paginação automática, formatos de coluna, seleção de células e Ctrl+C.
- **DateBox**: caixa de data com calendário, digitação rápida e limites de data.

## Um exemplo

```vb
Dim df As New Sage.DataFrame
df.ReadCsv "C:\dados\vendas.csv"

Dim resumo As Sage.DataFrame
Set resumo = df.Query("Valor > 100 And Cidade = ""SP""") _
               .GroupBy("Produto", "Total = sum(Valor), N = count(*)") _
               .SortValues("Total", False)

resumo.ToRange Sheets("Resumo").Range("A1")
Debug.Print resumo                     ' mostra a tabela formatada
```

## Instalação

**Requisitos:** Windows 10 ou 11 com **Excel de 64 bits**, .NET Framework 4.8 (já vem no Windows; se faltar, o instalador oferece instalar) e permissão de administrador.

1. Baixe o `SageFramework-Setup-x.y.z.exe` na [página de releases](https://github.com/sage-framework-2023/Sage-Framework/releases).
2. Feche o Excel e execute o instalador.
3. Abra o Excel, entre no editor do VBA (Alt+F11) e procure o menu **Sage**.

Para atualizar, execute o instalador da versão nova. Para desinstalar: *Configurações > Aplicativos > Sage Framework*.

> **Controles ActiveX no Microsoft 365:** para usar os controles do Sage nos UserForms, em *Arquivo > Opções > Central de Confiabilidade > Configurações da Central de Confiabilidade > Configurações do ActiveX*, marque *Avisar-me antes de habilitar controles UFI com restrições adicionais e controles SFI com restrições mínimas*. [Mais detalhes](https://sageframework.onrender.com/controles/#configuração-do-activex-no-office).

O instalador ainda não tem assinatura digital: se o Windows mostrar *"O Windows protegeu o computador"*, clique em **Mais informações** > **Executar assim mesmo**.

## Compilar a partir do código

Tudo é C# compilado com o `csc` do .NET Framework 4, que já vem no Windows: não precisa de Visual Studio nem de NuGet.

```powershell
# Cada componente, instalado só para o usuário atual (sem administrador)
powershell -ExecutionPolicy Bypass -File SageTypes\install.ps1
powershell -ExecutionPolicy Bypass -File SageEditor\install.ps1
powershell -ExecutionPolicy Bypass -File SageShortcuts\install.ps1

# O instalador da release (dist\SageFramework-Setup-<versão>.exe)
powershell -ExecutionPolicy Bypass -File build-release.ps1 -Version 1.2.0
```

Feche o Excel antes de instalar: ele mantém os DLLs carregados. A versão da release precisa ser a mesma do `AssemblyVersion` em `SageTypes\src\AssemblyInfo.cs`.

| Pasta | Componente |
|---|---|
| [`SageEditor/`](SageEditor/README.md) | Add-in COM do editor do VBA |
| [`SageTypes/`](SageTypes/README.md) | A biblioteca **Sage** e os controles para UserForms |
| `SageShortcuts/` | Os atalhos de teclado (roda como script do PowerShell) |
| `SageSetup/` | O instalador e o desinstalador |

## Contribuir

Encontrou um problema ou tem uma sugestão? Abra uma [issue](https://github.com/sage-framework-2023/Sage-Framework/issues). Se algo não funcionar, anexe o log do editor: `%APPDATA%\Sage\SageEditor.log`.

Os componentes de terceiros e suas licenças estão em `SageTypes/THIRD-PARTY-NOTICES.txt` (DuckDB) e `SageEditor/THIRD-PARTY-NOTICES.txt` (ícones).
