@echo off
rem Compila bin\SageTypes.dll com o csc do .NET Framework 4 (vem com o Windows).
setlocal
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%~dp0bin" mkdir "%~dp0bin"
"%CSC%" -nologo -target:library -platform:anycpu -optimize+ -codepage:65001 -out:"%~dp0bin\SageTypes.dll" -r:System.dll -r:System.Core.dll -r:System.Data.dll -r:System.Drawing.dll -r:System.Windows.Forms.dll -r:Microsoft.CSharp.dll -r:Microsoft.VisualBasic.dll "%~dp0src\*.cs"
exit /b %ERRORLEVEL%
