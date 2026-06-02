@echo off
cd /d "%~dp0"
title AutoCAD Numbering
echo Запуск AutoCAD Numbering от имени администратора...
powershell -Command "Start-Process 'AutoCAD_Numbering.exe' -Verb RunAs -WorkingDirectory '%~dp0'"
