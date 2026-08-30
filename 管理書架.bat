@echo off
chcp 65001 >nul
title Picture Book Shelf Manager
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0shelf-tool.ps1" %*
