@echo off
rem Abre o Sage Framework Installer (instalar, atualizar, reparar e desinstalar).
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0Instalador\Instalador.ps1"
