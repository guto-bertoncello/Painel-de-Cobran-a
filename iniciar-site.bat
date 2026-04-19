@echo off
setlocal
set SITE_DIR=%~dp0
start "Painel de Cobrança - Servidor" cmd /k py -3 -m http.server 8080 --directory "%SITE_DIR%"
timeout /t 2 >nul
start "" http://localhost:8080
endlocal
