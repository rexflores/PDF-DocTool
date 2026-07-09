@echo off
title PDF Record Tool (CLI)
echo Starting PDF Record Tool...
cd /d "%~dp0"

:: The following command bypasses execution policies and runs the PowerShell script
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; & '.\PDFRecordTool.ps1' }"

echo.
pause
