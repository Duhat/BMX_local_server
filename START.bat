@echo off
title BMX Graphics Server
color 0A
echo ========================================
echo    BMX BROADCAST GRAPHICS SERVER
echo ========================================
echo.
echo Проверка Node.js...
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ОШИБКА] Node.js не установлен!
    echo Скачайте с https://nodejs.org/
    pause
    exit /b 1
)
echo Node.js найден: 
node --version
echo.
echo Установка зависимостей...
call npm install
echo.
echo Запуск сервера...
echo.
call npm start
pause