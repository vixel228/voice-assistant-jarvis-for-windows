@echo off
chcp 65001 >nul
title Jarvis
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "%~dp0jarvis.py" %*
) else (
    python "%~dp0jarvis.py" %*
)
pause