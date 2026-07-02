@echo off
:: ============================================================
::  auto_push_data.bat  (CORREGIDO)
::  Sube los archivos de la carpeta data\ a GitHub SIN borrar
::  el codigo del dashboard (web\, scripts\) que se sube por la web.
::  Corre automaticamente a las 10am via Programador de Tareas.
:: ============================================================

set REPO_DIR=C:\Users\FTQ\apoyoconsultoria.com\File Server - 2025-070-O IJM - Linea de base\04. Analisis\03. Programacion\06. Dashboard Seguimiento Web
set GIT="C:\Program Files\Git\cmd\git.exe"
set PYTHON=C:\Users\FTQ\AppData\Local\anaconda3\python.exe

cd /d "%REPO_DIR%"
if errorlevel 1 (
    echo ERROR: No se encontro la carpeta del repositorio.
    exit /b 1
)

echo [%date% %time%] Iniciando subida de datos...

:: 1) Traer cambios remotos. --autostash guarda y restaura automaticamente
::    los cambios locales (Excels, etc.) sin perderlos.
%GIT% pull --rebase --autostash origin main
if errorlevel 1 (
    echo ERROR: No se pudo actualizar desde GitHub.
    exit /b 1
)

:: 2) Regenerar archivos JS del dashboard a partir de los Excels actualizados
"%PYTHON%" scripts\build_data.py
if errorlevel 1 (
    echo ERROR: Fallo build_data.py.
    exit /b 1
)

:: 3) Subir Excels y JS generados (sin --force)
%GIT% add data\ web\
%GIT% diff --cached --quiet
if errorlevel 1 (
    %GIT% commit -m "Actualizacion automatica %date% %time%"
    %GIT% push origin main
    if errorlevel 1 (
        echo ERROR: Fallo el push.
        exit /b 1
    )
    echo OK: Datos y dashboard subidos a GitHub.
) else (
    echo Sin cambios nuevos.
)
