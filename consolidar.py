"""
Consolidador de PRODUCCIÓN DEL DIGITADOR — USPP Satipo
=========================================================
Lee los mismos Excel (RAtenDet_*.xlsx) que usa el proyecto de
"Producción del Personal de Salud", pero desde la mirada del DIGITADOR:
agrupa cada FUA por la fecha en la que se REGISTRÓ (columna "Fecha Registro"),
no por la fecha en la que se atendió al paciente.

Por eso el archivo de salida se llama dig_AAAA_MM.json (mes de REGISTRO),
a diferencia de base_AAAA_MM.json del otro proyecto (mes de ATENCIÓN).

IMPORTANTE — DEDUPLICACIÓN POR FUA:
Los reportes RAtenDet que exporta el SIS son "acumulados": cada vez que
vuelves a exportar el mismo establecimiento, el archivo trae TODO el
historial otra vez (no solo lo nuevo). Si simplemente sumáramos todas las
filas de todos los Excel, una misma FUA que aparece en 2 o 3 exportaciones
distintas se contaría 2 o 3 veces, inflando artificialmente la producción.

Por eso este script:
  1) Extrae el número de FUA de cada fila (columna "FUA").
  2) Ordena los archivos cronológicamente según la fecha/hora que traen en
     su propio nombre (ej. "RAtenDet_18-08-2026_14-59-15...xlsx").
  3) Si una misma FUA aparece en más de un archivo, se queda con la versión
     del archivo MÁS RECIENTE (por si el registro fue corregido después) y
     descarta las repetidas.

Este es un proyecto independiente: se instala en su propia carpeta,
con su propio index.html y su propia carpeta data/. Solo comparte con el
otro proyecto la carpeta de Excel de origen (EXCEL_DIR).
"""
import os
import re
import json
from datetime import datetime
import pandas as pd

# ==========================
# CONFIG
# ==========================
EXCEL_DIR = r"C:\Users\USUARIO\Desktop\BASE DE DATOS\Escritorio\USPP\BASE DE DATOS 2"  # <-- tu carpeta real de excels
OUT_DIR = os.path.join(os.path.dirname(__file__), "data")
LIMIT_FILES = 0  # 0 = sin límite

os.makedirs(OUT_DIR, exist_ok=True)

# ==========================
# HELPERS
# ==========================
def norm_col(s: str) -> str:
    s = str(s).strip().lower()
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    s = s.replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u").replace("ñ","n")
    return s

def canon_dni(v) -> str:
    s = "" if v is None else str(v).strip()
    return re.sub(r"\D", "", s)

def parse_date_any(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s[:10], fmt)
        except Exception:
            pass
    try:
        return pd.to_datetime(s, errors="coerce").to_pydatetime()
    except Exception:
        return None

def parse_hora_any(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%H:%M")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            total_min = round(float(v) * 24 * 60)
            hh = (total_min // 60) % 24
            mm = total_min % 60
            return f"{hh:02d}:{mm:02d}"
        except Exception:
            return ""
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return ""
    m = re.match(r"^(\d{1,2}):(\d{2})", s)
    if m:
        hh = int(m.group(1)) % 24
        mm = int(m.group(2))
        return f"{hh:02d}:{mm:02d}"
    return ""

def detect_columns(df: pd.DataFrame) -> dict:
    mp = {}
    for c in df.columns:
        mp[norm_col(c)] = c
    return mp

def read_excel_any(path: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, engine="openpyxl")
    except Exception:
        return pd.read_excel(path)

def buscar_col(colmap, candidatos):
    for cand in candidatos:
        if cand in colmap:
            return colmap[cand]
    return None

def extraer_timestamp_archivo(nombre: str):
    """Saca la fecha/hora de exportación del propio nombre del archivo,
    ej. 'RAtenDet_18-08-2026_14-59-15-puerto ocopa.xlsx' -> datetime(2026,8,18,14,59,15).
    Si no puede, devuelve None (esos archivos se procesan al final, en orden alfabético)."""
    m = re.search(r"(\d{2})-(\d{2})-(\d{4})_(\d{2})-(\d{2})-(\d{2})", nombre)
    if not m:
        return None
    dd, mm, yyyy, hh, mi, ss = map(int, m.groups())
    try:
        return datetime(yyyy, mm, dd, hh, mi, ss)
    except Exception:
        return None

CAND_FECHA_ATEN   = ["fecha atencion", "fecha de atencion", "fec atencion", "fec. atencion", "fecha", "fecha_atencion"]
CAND_EESS         = ["eess", "establecimiento", "ipress", "nombre eess", "establecimieto", "establecimiento de salud"]
CAND_PPDD         = ["ppdd", "punto digitacion", "punto de digitacion", "p.p.d.d.", "punto digitación"]
CAND_SERVICIO     = ["servicio", "prestacion", "prestación", "procedimiento", "descripcion servicio", "descrip. servicio"]
CAND_ID_SERVICIO  = ["id servicio", "id_servicio", "idservicio", "cpms", "id procedimiento"]
CAND_DIGITADOR    = ["digitador", "usuario digitador", "usuario registro", "dni digitador"]
CAND_FECHA_REG    = ["fecha registro", "fecha de registro", "fec. registro", "fecha_registro"]
CAND_HORA_REG     = ["hora registro", "hora de registro", "hora_registro"]
CAND_FUA          = ["fua", "n fua", "nro fua", "numero fua", "n° fua", "num fua"]

# ==========================
# MAIN
# ==========================
def main():
    files = [f for f in os.listdir(EXCEL_DIR) if f.lower().endswith((".xlsx", ".xls"))]
    if LIMIT_FILES > 0:
        files = files[:LIMIT_FILES]
    if not files:
        print("❌ No encontré excels en:", EXCEL_DIR)
        return

    # Orden cronológico según la fecha/hora en el propio nombre del archivo,
    # para que si una FUA se repite, se quede con la versión MÁS RECIENTE.
    files.sort(key=lambda f: (extraer_timestamp_archivo(f) or datetime.min, f))

    vistos = {}         # fua -> rec (se sobreescribe con la versión más reciente)
    sin_fua_contador = 0
    total_filas = 0
    omitidos_sin_digitador = 0
    omitidos_sin_fecha_reg = 0
    archivos_omitidos = []

    for i, fn in enumerate(files, 1):
        path = os.path.join(EXCEL_DIR, fn)
        print(f"📄 [{i}/{len(files)}] Leyendo: {fn}")
        try:
            df = read_excel_any(path)
        except Exception as e:
            print("   ⚠️ No pude leer:", fn, "->", e)
            archivos_omitidos.append(fn)
            continue
        if df is None or df.empty:
            print("   ⚠️ Vacío:", fn)
            continue

        colmap = detect_columns(df)
        digitador_col = buscar_col(colmap, CAND_DIGITADOR)
        fecha_reg_col = buscar_col(colmap, CAND_FECHA_REG)

        if digitador_col is None or fecha_reg_col is None:
            encontrado = False
            for header_row in (1, 2, 3):
                try:
                    df_retry = pd.read_excel(path, engine="openpyxl", header=header_row)
                except Exception:
                    continue
                if df_retry is None or df_retry.empty:
                    continue
                colmap_retry = detect_columns(df_retry)
                dcol = buscar_col(colmap_retry, CAND_DIGITADOR)
                fcol = buscar_col(colmap_retry, CAND_FECHA_REG)
                if dcol and fcol:
                    df, colmap = df_retry, colmap_retry
                    digitador_col, fecha_reg_col = dcol, fcol
                    encontrado = True
                    print(f"   ↻ Encontrada leyendo desde la fila {header_row+1}")
                    break
            if not encontrado:
                print("   ⚠️ No encuentro columnas 'Digitador' / 'Fecha Registro' en", fn, "-> omito archivo")
                print("      Columnas detectadas:", list(df.columns))
                archivos_omitidos.append(fn)
                continue

        fecha_col    = buscar_col(colmap, CAND_FECHA_ATEN)
        eess_col     = buscar_col(colmap, CAND_EESS)
        ppdd_col     = buscar_col(colmap, CAND_PPDD)
        servicio_col = buscar_col(colmap, CAND_SERVICIO)
        id_serv_col  = buscar_col(colmap, CAND_ID_SERVICIO)
        hora_reg_col = buscar_col(colmap, CAND_HORA_REG)
        fua_col      = buscar_col(colmap, CAND_FUA)

        if fua_col is None:
            print("   ⚠️ Este archivo no tiene columna 'FUA' identificable — no podré deduplicar sus filas por FUA")

        rows = df.to_dict(orient="records")
        total_filas += len(rows)

        for row in rows:
            dig_dni = canon_dni(row.get(digitador_col, ""))
            if not dig_dni:
                omitidos_sin_digitador += 1
                continue

            fr = parse_date_any(row.get(fecha_reg_col))
            if fr is None:
                omitidos_sin_fecha_reg += 1
                continue

            fua = str(row.get(fua_col, "")).strip() if fua_col else ""
            if not fua:
                sin_fua_contador += 1
                fua = f"__sinfua__{fn}__{sin_fua_contador}"

            fa = None if fecha_col is None else parse_date_any(row.get(fecha_col))
            eess = "" if eess_col is None else str(row.get(eess_col, "")).strip()
            ppdd = "" if ppdd_col is None else str(row.get(ppdd_col, "")).strip()
            servicio = "" if servicio_col is None else str(row.get(servicio_col, "")).strip()
            id_serv = "" if id_serv_col is None else str(row.get(id_serv_col, "")).strip()
            hora_reg = "" if hora_reg_col is None else parse_hora_any(row.get(hora_reg_col))

            vistos[fua] = {
                "fua": fua,  # se guarda para poder actualizar meses ya consolidados sin duplicar (ver actualizar_mes.py)
                "digitador_dni": dig_dni,
                "anio_registro": fr.year, "mes_registro": fr.month, "dia_registro": fr.day,
                "fecha_registro": fr.strftime("%Y-%m-%d"),
                "hora_registro": hora_reg,
                "fecha_atencion": fa.strftime("%Y-%m-%d") if fa else "",
                "establecimiento": eess,
                "punto_digitacion": ppdd,
                "servicio": servicio,
                "id_servicio": id_serv,
                "source": fn
            }

    registros_finales = list(vistos.values())
    duplicados_quitados = total_filas - len(registros_finales) - omitidos_sin_digitador - omitidos_sin_fecha_reg

    buckets = {}
    for rec in registros_finales:
        key = (rec["anio_registro"], rec["mes_registro"])
        buckets.setdefault(key, []).append(rec)

    disponibles = []
    for (anio, mes), regs in sorted(buckets.items()):
        fname = f"dig_{anio}_{mes:02d}.json"
        with open(os.path.join(OUT_DIR, fname), "w", encoding="utf-8") as f:
            json.dump(regs, f, ensure_ascii=False)
        disponibles.append({"anio": anio, "mes": mes, "file": fname})
        print(f"✅ Guardado {fname}  ({len(regs)} FUAs digitados)")

    manifest = {
        "disponibles": disponibles,
        "ultimo": disponibles[-1] if disponibles else {"anio": 0, "mes": 0},
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "origen": EXCEL_DIR
    }
    with open(os.path.join(OUT_DIR, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print("\n==========================")
    print("✅ CONSOLIDACIÓN TERMINADA (Producción del Digitador)")
    print("Total filas leídas (todos los Excel, con repetidos):", total_filas)
    print("FUAs únicas después de deduplicar:", len(registros_finales))
    print("Filas duplicadas descartadas (misma FUA en varios archivos):", max(duplicados_quitados, 0))
    print("Filas sin número de FUA identificable:", sin_fua_contador, "(se conservaron todas, no se pudieron deduplicar)")
    print("Omitidas (sin DNI de digitador):", omitidos_sin_digitador)
    print("Omitidas (sin fecha de registro):", omitidos_sin_fecha_reg)
    if archivos_omitidos:
        print("Archivos que no se pudieron leer:", archivos_omitidos)
    print("Meses generados:", len(disponibles))
    print("Salida:", OUT_DIR)
    print("==========================")

if __name__ == "__main__":
    main()
