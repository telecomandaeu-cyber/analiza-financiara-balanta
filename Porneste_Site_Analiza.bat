@echo off
chcp 65001 >nul
title Platforma Locala - Analiza Financiara (.DOCX)
echo ========================================================================
echo   Pornire Site Local - Analiza Pozitiei si Performantelor Financiare
echo ========================================================================

set "PY_EXE=%LOCALAPPDATA%\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.13_qbz5n2kfra8p0\python.exe"
if not exist "%PY_EXE%" set "PY_EXE=python"

"%PY_EXE%" "%~dp0site_analiza_financiara.py"
pause

