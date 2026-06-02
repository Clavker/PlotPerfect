@echo off
title AutoCAD Task Manager
echo Запуск AutoCAD Task Manager от имени администратора...
powershell -Command "Start-Process 'AutoCAD_Task_Manager.exe' -Verb RunAs -WorkingDirectory '%~dp0'"
