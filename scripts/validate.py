# -*- coding: utf-8 -*-
"""
validate.py — Valida la base ANTES de regenerar el dashboard.

Si encuentra errores, termina con código de salida 1 y escribe un reporte
claro en validation_report.txt. El workflow de GitHub Actions usa ese código
para NO publicar una versión rota: el dashboard anterior queda en línea.

Uso:
    python scripts/validate.py
"""
import sys
import io
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
EXCEL = ROOT / "data" / "Base - Prevalencia y dependencia.xlsx"
REPORT = ROOT / "validation_report.txt"

# Categorías de resultado conocidas (si aparece una nueva, se avisa como ERROR
# para decidir conscientemente cómo clasificarla en las donas).
RESULTADOS_VALIDOS = {
    "Completa",
    "Incompleta",
    "Rechazada",
    "Ausente",
    "No hay población objetivo en la vivienda",
    "Discapacidad",
    "Cita - Aplazada",
}

COLS_BASE = ["Fecha", "DDJJ", "Grupo", "Departamento", "Sexo", "Edad", "Resultado", "Inicio", "Fin"]
COLS_CUOTAS = ["Distrito Judicial", "Grupo etario", "Sexo", "Cuota"]
COLS_AUDIOS = ["Id", "Grupo", "DDJJ", "Resultado", "Duracion_min"]

errors: list[str] = []
warnings: list[str] = []


def err(msg: str) -> None:
    errors.append(msg)


def warn(msg: str) -> None:
    warnings.append(msg)


def main() -> int:
    if not EXCEL.exists():
        err(f"No se encontró el archivo: {EXCEL.name} en la carpeta data/")
        return finish()

    try:
        xls = pd.ExcelFile(EXCEL)
    except Exception as e:
        err(f"No se pudo abrir el Excel (¿archivo corrupto o aún abierto en Excel?): {e}")
        return finish()

    # --- Hojas requeridas -------------------------------------------------
    for hoja in ["Base (0)", "Cuotas", "Audios"]:
        if hoja not in xls.sheet_names:
            err(f"Falta la hoja '{hoja}'. Hojas encontradas: {xls.sheet_names}")
    if errors:
        return finish()

    base = pd.read_excel(EXCEL, sheet_name="Base (0)")
    cuotas = pd.read_excel(EXCEL, sheet_name="Cuotas")
    audios = pd.read_excel(EXCEL, sheet_name="Audios")

    # --- Columnas ---------------------------------------------------------
    for col in COLS_BASE:
        if col not in base.columns:
            err(f"Hoja 'Base (0)': falta la columna '{col}'.")
    for col in COLS_CUOTAS:
        if col not in cuotas.columns:
            err(f"Hoja 'Cuotas': falta la columna '{col}'.")
    for col in COLS_AUDIOS:
        if col not in audios.columns:
            err(f"Hoja 'Audios': falta la columna '{col}'.")
    if errors:
        return finish()

    # --- Base (0): contenido ------------------------------------------------
    if len(base) == 0:
        err("Hoja 'Base (0)' está vacía.")
        return finish()

    # Fechas
    fechas = pd.to_datetime(base["Fecha"], dayfirst=True, errors="coerce")
    n_bad = fechas.isna().sum()
    if n_bad > 0:
        filas = (base.index[fechas.isna()] + 2).tolist()[:10]
        err(f"Hoja 'Base (0)': {n_bad} fecha(s) inválida(s) en columna 'Fecha' "
            f"(filas de Excel: {filas}{'...' if n_bad > 10 else ''}). Formato esperado: dd/mm/aaaa.")

    # Resultados
    res = base["Resultado"].dropna().unique()
    nuevos = sorted(set(map(str, res)) - RESULTADOS_VALIDOS)
    if nuevos:
        err("Hoja 'Base (0)': categorías de 'Resultado' NO reconocidas: "
            f"{nuevos}. Si son válidas, agregarlas a RESULTADOS_VALIDOS en scripts/validate.py "
            "y definir su clasificación en scripts/build_data.py.")
    n_sin_res = base["Resultado"].isna().sum()
    if n_sin_res > 0:
        warn(f"Hoja 'Base (0)': {n_sin_res} fila(s) sin 'Resultado' (se excluirán del dashboard).")

    # Grupos
    grupos = set(base["Grupo"].dropna().astype(str).unique())
    desconocidos = grupos - {"Mujer", "Adolescentes"}
    if desconocidos:
        err(f"Hoja 'Base (0)': valores de 'Grupo' no reconocidos: {sorted(desconocidos)} "
            "(se esperan 'Mujer' o 'Adolescentes').")

    # --- Cuotas -------------------------------------------------------------
    if not pd.api.types.is_numeric_dtype(cuotas["Cuota"]):
        err("Hoja 'Cuotas': la columna 'Cuota' tiene valores no numéricos.")
    elif cuotas["Cuota"].isna().any():
        err("Hoja 'Cuotas': hay celdas vacías en la columna 'Cuota'.")

    # Distritos de la base presentes en cuotas
    if not errors:
        ddjj_base = set(base["DDJJ"].dropna().astype(str).unique())
        ddjj_cuotas = set(cuotas["Distrito Judicial"].dropna().astype(str).unique())
        sin_cuota = ddjj_base - ddjj_cuotas
        if sin_cuota:
            err(f"Distritos judiciales con encuestas pero SIN cuota definida: {sorted(sin_cuota)}.")

    # --- Chequeos suaves ------------------------------------------------------
    dup = base.duplicated().sum()
    if dup > 0:
        warn(f"Hoja 'Base (0)': {dup} fila(s) completamente duplicada(s).")

    if not fechas.isna().all():
        hoy = pd.Timestamp.today().normalize()
        futuras = (fechas > hoy).sum()
        if futuras > 0:
            warn(f"Hoja 'Base (0)': {futuras} encuesta(s) con fecha futura.")

    return finish()


def finish() -> int:
    buf = io.StringIO()
    buf.write("REPORTE DE VALIDACIÓN — Base - Prevalencia y dependencia.xlsx\n")
    buf.write("=" * 64 + "\n\n")
    if errors:
        buf.write(f"ERRORES ({len(errors)}) — el dashboard NO se actualizará hasta corregirlos:\n")
        for i, e in enumerate(errors, 1):
            buf.write(f"  {i}. {e}\n")
        buf.write("\n")
    if warnings:
        buf.write(f"ADVERTENCIAS ({len(warnings)}) — no bloquean la publicación:\n")
        for i, w in enumerate(warnings, 1):
            buf.write(f"  {i}. {w}\n")
        buf.write("\n")
    if not errors and not warnings:
        buf.write("Sin errores ni advertencias. ✔\n")
    elif not errors:
        buf.write("Sin errores bloqueantes. ✔\n")

    texto = buf.getvalue()
    print(texto)
    REPORT.write_text(texto, encoding="utf-8")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
