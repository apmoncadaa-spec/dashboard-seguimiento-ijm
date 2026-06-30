@echo off
:: ============================================================
::  auto_push_data.bat  (CORREGIDO)
::  Sube los archivos de la carpeta data\ a GitHub SIN borrar
::  el codigo del dashboard (web\, scripts\) que se sube por la web.
::  Corre automaticamente a las 10am via Programador de Tareas.
:: ============================================================

set REPO_DIR=C:\Users\FTQ\apoyoconsultoria.com\File Server - 2025-070-O IJM - Linea de base\04. Analisis\03. Programacion\06. Dashboard Seguimiento Web
set GIT="C:\Program Files\Git\cmd\git.exe"

cd /d "%REPO_DIR%"
if errorlevel 1 (
    echo ERROR: No se encontro la carpeta del repositorio.
    exit /b 1
)

echo [%date% %time%] Iniciando subida de datos...

:: 1) Traer PRIMERO los cambios del repositorio (codigo del dashboard, etc.)
::    Asi esta copia local se mantiene al dia y no se pisa el trabajo de otros.
%GIT% pull --rebase origin main
if errorlevel 1 (
    echo ERROR: No se pudo actualizar desde GitHub ^(git pull --rebase^).
    echo        Resuelve el conflicto manualmente y vuelve a ejecutar.
    exit /b 1
)

:: 2) Subir SOLO los cambios de la carpeta data\  (sin --force)
%GIT% add data\
%GIT% diff --cached --quiet
if errorlevel 1 (
    %GIT% commit -m "Actualizacion automatica %date% %time%"
    %GIT% push origin main
    if errorlevel 1 (
        echo ERROR: Fallo el push. Revisa la conexion o vuelve a ejecutar.
        exit /b 1
    )
    echo OK: Archivos de datos subidos a GitHub.
) else (
    echo Sin cambios nuevos.
)
