@echo off
:: ============================================================
::  auto_push_data.bat
::  Sube todos los archivos de esta carpeta a GitHub.
::  Corre automaticamente a las 10am via Programador de Tareas.
:: ============================================================

set DATA_DIR=C:\Users\FTQ\apoyoconsultoria.com\File Server - 2025-070-O IJM - Linea de base\04. Analisis\03. Programacion\06. Dashboard Seguimiento Web\data
set GIT="C:\Program Files\Git\cmd\git.exe"

cd /d "%DATA_DIR%"
if errorlevel 1 (
    echo ERROR: No se encontro la carpeta de datos.
    exit /b 1
)

echo [%date% %time%] Iniciando subida de datos...

%GIT% pull origin main --rebase
if errorlevel 1 (
    echo AVISO: No se pudo hacer pull. Continuando...
)

%GIT% add .
%GIT% diff --cached --quiet
if errorlevel 1 (
    %GIT% commit -m "Actualizacion automatica %date% %time%"
    %GIT% push origin main
    echo OK: Archivos subidos a GitHub.
) else (
    echo Sin cambios nuevos.
)
