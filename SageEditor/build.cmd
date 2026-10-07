@echo off
rem Compila bin\SageEditor.dll com o csc do .NET Framework 4 (vem com o Windows).
setlocal
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
rem PowerShell 5.1 (aba Terminal): vem com o Windows, no GAC
set PS=%WINDIR%\Microsoft.NET\assembly\GAC_MSIL\System.Management.Automation\v4.0_3.0.0.0__31bf3856ad364e35\System.Management.Automation.dll
if not exist "%~dp0bin" mkdir "%~dp0bin"
rem Icones do vscode-icons (MIT) embutidos como recursos "icons.ARQUIVO.png". A lista vai num
rem arquivo de resposta: na linha de comando passaria do limite do Windows (8191 caracteres).
set RSP=%~dp0bin\resources.rsp
type nul > "%RSP%"
for %%f in ("%~dp0res\icons\*.png") do echo -resource:"%%f",icons.%%~nxf>> "%RSP%"
"%CSC%" -nologo -target:library -platform:anycpu -optimize+ -codepage:65001 -out:"%~dp0bin\SageEditor.dll" -r:System.dll -r:System.Core.dll -r:System.Drawing.dll -r:System.Windows.Forms.dll -r:Microsoft.CSharp.dll -r:"%PS%" @"%RSP%" "%~dp0src\*.cs"
exit /b %ERRORLEVEL%
