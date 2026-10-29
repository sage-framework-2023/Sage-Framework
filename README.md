# Sage Framework

**VBA, modernized, without leaving Excel.** Sage Framework brings the experience of a modern IDE to the VBA editor and adds modern types to the language for text, lists, dates, tables, JSON, APIs and databases, plus new controls for UserForms. All in a single installer.

📖 **Documentation:** https://sageframework.onrender.com
⬇️ **Download:** [Releases](https://github.com/sage-framework-2023/Sage-Framework/releases) (`SageFramework-Setup-x.y.z.exe`)
📝 **What's new:** [CHANGELOG.md](CHANGELOG.md)

## What's included

### SageEditor: a renewed VBA editor
- **Themes** for the whole editor: *Dark Modern*, *Dark+*, *Sage*, *Light Modern* or the VBE default.
- Richer **syntax highlighting**, **line numbers**, **tabs** for open windows and **multiple cursors**.
- **Color decorators**: every `RGB(...)` and `&H...&` is underlined in its own color; clicking it opens a color picker that rewrites the color in the code.
- **Unified explorer** (Properties inside the Project window) and **icons** for each component type.
- **Terminal window**: Immediate, Locals, Watches, a PowerShell that runs your VBA, and the *DataFrame Results* tab.
- **Automatic reference** to the Sage library in every new workbook.
- Interface in English or Portuguese, following the Office language.

### SageShortcuts: keyboard shortcuts
Configurable in a text file: `Ctrl+K, Ctrl+C` to comment, `Alt+Arrow` to move lines, `Shift+Alt+Arrow` to duplicate, `Ctrl+J` for the Terminal, `F12` to go to definition, and more.

### SageTypes: the **Sage** library
| Type | What for |
|---|---|
| `StringS`, `ListS`, `DictionaryS` | Text, lists and dictionaries with chainable methods, negative indexes, slices and clear errors (`KeyError`, `IndexError`) |
| `DatetimeS` | Date and time with time zones, business days, start and end of period |
| `DataFrame` | Tables backed by DuckDB, for millions of rows |
| `Json`, `Requests`, `SqlEngine` | Read and write JSON, call APIs and connect to PostgreSQL, MySQL, SQLite, SQL Server, Access and ODBC |
| `Resizable`, `OfficeTheme` | Resizable UserForms and the current Office theme |

### Controls for UserForms
On the **Sage** page of the Toolbox:
- **TextBox** with Windows spell checking and placeholder text.
- **DataGrid** that shows a `DataFrame` with automatic paging, column formats, cell selection and Ctrl+C.
- **DateBox**: a date box with a calendar, quick typing and date limits.

## An example

```vb
Dim df As New Sage.DataFrame
df.ReadCsv "C:\data\sales.csv"

Dim summary As Sage.DataFrame
Set summary = df.Query("Amount > 100 And City = ""NY""") _
                .GroupBy("Product", "Total = sum(Amount), N = count(*)") _
                .SortValues("Total", False)

summary.ToRange Sheets("Summary").Range("A1")
Debug.Print summary                    ' prints the formatted table
```

## Installation

**Requirements:** Windows 10 or 11 with **64-bit Excel**, .NET Framework 4.8 (included in Windows; if it is missing, the installer offers to install it) and administrator rights.

1. Download `SageFramework-Setup-x.y.z.exe` from the [releases page](https://github.com/sage-framework-2023/Sage-Framework/releases).
2. Close Excel and run the installer.
3. Open Excel, go to the VBA editor (Alt+F11) and look for the **Sage** menu.

To update, run the installer of the new version. To uninstall: *Settings > Apps > Sage Framework*.

> **ActiveX controls in Microsoft 365:** Microsoft 365 now ships with *Disable all controls without notification*, which blocks the Sage controls in UserForms with the error *"The subject is not trusted for the specified action"*. The installer option **Allow Sage controls in UserForms (Office ActiveX setting)**, checked by default, fixes this for every user on the computer. To do it by hand: *File > Options > Trust Center > Trust Center Settings > ActiveX Settings*, select *Prompt me before enabling Unsafe for Initialization (UFI) controls with additional restrictions and Safe for Initialization (SFI) controls with minimal restrictions*. [More details](https://sageframework.onrender.com/controles/#configuração-do-activex-no-office).

The installer is not digitally signed yet: if Windows shows *"Windows protected your PC"*, click **More info** > **Run anyway**.

## Building from source

Everything is C# compiled with the .NET Framework 4 `csc`, which ships with Windows: no Visual Studio or NuGet needed.

```powershell
# Each component, installed for the current user only (no administrator rights)
powershell -ExecutionPolicy Bypass -File SageTypes\install.ps1
powershell -ExecutionPolicy Bypass -File SageEditor\install.ps1
powershell -ExecutionPolicy Bypass -File SageShortcuts\install.ps1

# The release installer (dist\SageFramework-Setup-<version>.exe)
powershell -ExecutionPolicy Bypass -File build-release.ps1 -Version 1.3.0
```

Close Excel before installing: it keeps the DLLs loaded. The release version must match the `AssemblyVersion` in `SageTypes\src\AssemblyInfo.cs`.

| Folder | Component |
|---|---|
| [`SageEditor/`](SageEditor/README.md) | COM add-in for the VBA editor |
| [`SageTypes/`](SageTypes/README.md) | The **Sage** library and the UserForm controls |
| `SageShortcuts/` | The keyboard shortcuts (runs as a PowerShell script) |
| `SageSetup/` | The installer and uninstaller |

## Contributing

Found a problem or have a suggestion? Open an [issue](https://github.com/sage-framework-2023/Sage-Framework/issues). If something doesn't work, attach the editor log: `%APPDATA%\Sage\SageEditor.log`.

Third-party components and their licenses are listed in `SageTypes/THIRD-PARTY-NOTICES.txt` (DuckDB) and `SageEditor/THIRD-PARTY-NOTICES.txt` (icons).
