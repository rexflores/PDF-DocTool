@echo off
title PDF Record Tool (GUI)
cd /d "%~dp0"

:: The following command bypasses execution policies and runs the PowerShell script
:: We use 'start ""' and '-WindowStyle Hidden' so the black command prompt window doesn't stay open in the background.
start "" powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -Command "& '.\PDFRecordTool-GUI.ps1'"
