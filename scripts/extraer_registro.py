# -*- coding: utf-8 -*-
"""
extraer_registro.py — Genera el extracto liviano que alimenta la pestaña
"Registro diario de encuestadoras" a partir de las DOS bases crudas.

Por qué existe
--------------
La pestaña antes se alimentaba de "Registro diario de encuestadoras.xlsx".
Ahora la fuente de verdad son las dos bases crudas de campo:

    DATA_ADOLESCENTES_V1_CSV.CSV   (grupo: Adolescentes)
    DATA_MUJERES_CSV.CSV           (grupo: Mujeres)

Esas bases pesan ~28-30 MB cada una (≈300 columnas del cuestionario completo),
demasiado para subirlas al repo por la web de GitHub (límite 25 MB). Este script
lee ambas y guarda en data/ un EXTRACTO compacto con SOLO las 5 columnas que el
dashboard necesita, así que pesa unos pocos cientos de KB y se sube igual que
antes. El dato es el mismo, sin el peso muerto de las ~295 columnas que no se usan.

Flujo de uso
------------
    python scripts/extraer_registro.py        # regenera el extracto en data/
    # luego sube  data/Registro - bases crudas.csv  a GitHub (igual que hoy)

La ubicación de las bases crudas se resuelve sola (carpeta hermana del repo). Se
puede forzar con la variable de entorno BASES_CRUDAS_DIR o pasando la ruta como
primer argumento:

    python scripts/extraer_registro.py "D:\\ruta\\a\\01. Bases crudas"
"""
import csv
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "Registro - bases crudas.csv"

# Cada base cruda y el grupo poblacional que representa.
FUENTES = {
    "DATA_ADOLESCENTES_V1_CSV.CSV": "Adolescentes",
    "DATA_MUJERES_CSV.CSV": "Mujeres",
}

# Columnas que conserva el extracto (se localizan POR NOMBRE en el encabezado,
# así no importa que cada base las tenga en posiciones distintas).
COLS_FUENTE = ["CODENCU", "ENCUESTADOR", "FECHAINICIOENC", "FECHAFINENC", "P800RESULTADO"]
# Nombres de salida (en minúscula, como los espera build_data.py / build_registro).
COLS_SALIDA = ["grupo", "codencu", "encuestador", "fechainicioenc", "fechafinenc", "p800resultado"]


def _ubicar_bases():
    """Devuelve la carpeta '01. Bases crudas'. Orden de búsqueda:
    1) argumento de línea de comandos, 2) variable BASES_CRUDAS_DIR,
    3) ubicación estándar relativa al repo (carpeta hermana 04. Bases de datos)."""
    candidatos = []
    if len(sys.argv) > 1:
        candidatos.append(Path(sys.argv[1]))
    if os.environ.get("BASES_CRUDAS_DIR"):
        candidatos.append(Path(os.environ["BASES_CRUDAS_DIR"]))
    # .../04. Analisis/03. Programacion/06. Dashboard...  ->  .../04. Analisis/04. Bases de datos/...
    candidatos.append(ROOT.parent.parent / "04. Bases de datos" / "01. Encuestas" / "01. Bases crudas")
    for c in candidatos:
        if c and c.is_dir():
            return c
    rutas = "\n  ".join(str(c) for c in candidatos)
    raise SystemExit(
        "ERROR: no se encontró la carpeta '01. Bases crudas'. Probé:\n  " + rutas +
        "\nPásala como argumento o define BASES_CRUDAS_DIR."
    )


# --- Reparación de filas con columnas corridas (2026-07-13) ------------------
# Algunas direcciones traen ';' dentro del texto (p. ej. "CALLE X 117; LOS
# OLIVOS"). Como ';' es el separador del CSV, esas filas quedan con columnas de
# MÁS y todos los campos posteriores se corren. Sin reparación, el código de la
# encuestadora cae en otra columna y el dashboard atribuye la encuesta a la
# persona equivocada (p. ej. la supervisora con código 01).
_RE_FECHA = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")
_RE_COD = re.compile(r"^\d{1,3}$")


def _tiene_letras(s):
    return bool(re.search(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]", s or ""))


def _realinear(row, hmap, extra):
    """Devuelve [codencu, encuestador, fechainicio, fechafin, p800] de una fila
    con `extra` columnas de más, validando el patrón de cada campo. El bloque
    CODENCU;ENCUESTADOR;CODSUP;SUPERVISOR (número, nombre, número, nombre) sirve
    de ancla; las fechas y el resultado se validan por formato."""
    ic = hmap["CODENCU"]
    # 1) ancla: buscar el patrón número/nombre/número/nombre cerca de CODENCU
    s_cod = None
    for k in range(0, extra + 1):
        j = ic + k
        if j + 3 < len(row) and _RE_COD.match(row[j].strip()) \
                and _tiene_letras(row[j + 1]) and not _RE_COD.match(row[j + 1].strip()) \
                and _RE_COD.match(row[j + 2].strip()) and _tiene_letras(row[j + 3]):
            s_cod = k
            break
    if s_cod is None:
        return None
    codencu = row[ic + s_cod].strip()
    encuestador = row[hmap["ENCUESTADOR"] + s_cod].strip()

    # 2) fechas: validar formato dd/mm/aaaa probando corrimientos
    def _fecha(col):
        i = hmap[col]
        for k in (extra, s_cod, 0):
            j = i + k
            if j < len(row) and _RE_FECHA.match(row[j].strip()):
                return row[j].strip()
        return ""

    fi, ff = _fecha("FECHAINICIOENC"), _fecha("FECHAFINENC")

    # 3) resultado: código de 2 dígitos, probando corrimientos
    p800 = ""
    ip = hmap["P800RESULTADO"]
    for k in (extra, s_cod, 0):
        j = ip + k
        if j < len(row) and re.match(r"^\d{2}$", row[j].strip()):
            p800 = row[j].strip()
            break
    return [codencu, encuestador, fi, ff, p800]


def main():
    base_dir = _ubicar_bases()
    filas = []
    resumen = {}
    for archivo, grupo in FUENTES.items():
        ruta = base_dir / archivo
        if not ruta.exists():
            raise SystemExit(f"ERROR: falta {archivo} en {base_dir}")
        # Las bases vienen en latin-1 y separadas por ';'.
        with open(ruta, encoding="latin-1", newline="") as fh:
            rd = csv.reader(fh, delimiter=";")
            header = next(rd, None)
            if header is None:
                continue
            hmap = {str(h).strip().upper(): i for i, h in enumerate(header)}
            faltan = [c for c in COLS_FUENTE if c not in hmap]
            if faltan:
                raise SystemExit(f"ERROR: {archivo} no tiene columnas {faltan}")
            idx = [hmap[c] for c in COLS_FUENTE]
            maxi = max(idx)
            ncols = len(header)
            reparadas = 0
            n = 0
            for row in rd:
                if len(row) <= maxi:
                    continue
                if len(row) != ncols:
                    # Fila con columnas corridas por ';' dentro de un texto:
                    # realinear validando el patrón de cada campo.
                    vals = _realinear(row, hmap, len(row) - ncols)
                    if vals is None:
                        print(f"   AVISO: fila corrida NO reparable en {archivo} "
                              f"(cols={len(row)}, esperadas={ncols}); se omite.")
                        continue
                    reparadas += 1
                else:
                    vals = [str(row[i]).strip() for i in idx]
                # Descarta filas totalmente vacías en las columnas clave.
                if not any(vals):
                    continue
                # Salta filas sin encuestadora NI código (no aportan a la matriz).
                if not vals[0] and not vals[1]:
                    continue
                filas.append([grupo] + vals)
                n += 1
            resumen[grupo] = n
            if reparadas:
                print(f"   {archivo}: {reparadas} fila(s) con columnas corridas reparadas.")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    # UTF-8 con BOM para que Excel lo abra bien si alguien lo revisa a mano.
    with open(OUT, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(COLS_SALIDA)
        w.writerows(filas)

    comp = sum(1 for f in filas if f[5] == "01")
    print(f"OK -> {OUT.relative_to(ROOT)}")
    print(f"   filas: {len(filas)} | por grupo: {resumen} | completas (01): {comp}")
    print(f"   tamaño: {OUT.stat().st_size/1024:.0f} KB")


if __name__ == "__main__":
    main()
