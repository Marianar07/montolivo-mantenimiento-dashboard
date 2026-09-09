#!/usr/bin/env python3
"""
Procesador del Tablero de Mantenimiento - Inversiones Montolivo
=================================================================

Lee los 4 exports del CMMS (OT, SS, Disponibilidad, Criticidad de equipos),
los consolida y cruza, y genera:
  1. data.json   -> los datos ya calculados, en el formato que espera el tablero
  2. index.html  -> el tablero final (plantilla + datos), listo para subir a GitHub

USO:
    python process_dashboard.py

Por defecto busca los 4 archivos de entrada en la carpeta ./datos_nuevos/
usando coincidencia flexible de nombre (ver ENCONTRAR ARCHIVOS DE ENTRADA
más abajo). Si prefieres indicar los archivos exactos, edita las rutas en
la sección CONFIGURACIÓN, o pásalas como argumentos:

    python process_dashboard.py ruta_OT.xlsx ruta_SS.xlsx ruta_Disponibilidad.xlsx ruta_Criticidad.xlsx

Ver LEEME_PROCESO.md (o CLAUDE.md) en esta misma carpeta para el detalle
completo de la lógica de cruce, los supuestos de datos y las limitaciones
conocidas del CMMS.
"""

import sys
import re
import json
from pathlib import Path
from datetime import datetime, date

import pandas as pd
import numpy as np

# ============================================================
# CONFIGURACIÓN
# ============================================================

CARPETA_BASE = Path(__file__).resolve().parent
CARPETA_DATOS = CARPETA_BASE / "datos_nuevos"
PLANTILLA_HTML = CARPETA_BASE / "index_template.html"
SALIDA_JSON = CARPETA_BASE / "data.json"
SALIDA_HTML = CARPETA_BASE / "index.html"

# Umbrales de semáforo (provisionales - validar con el jefe de mantenimiento)
UMBRALES = {
    "disponibilidad": {"verde": 95, "amarillo": 90},          # % , mayor es mejor
    "cumplimiento_preventivo": {"verde": 90, "amarillo": 75},  # % , mayor es mejor
    "otif": {"verde": 90, "amarillo": 75},                     # % , mayor es mejor
    "paradas_no_programadas": {"verde": 5, "amarillo": 10},    # % , menor es mejor
}


# ============================================================
# ENCONTRAR ARCHIVOS DE ENTRADA
# ============================================================

def _columnas_hoja(path: Path, sheet_name=0):
    """Lee solo la fila de encabezados de una hoja (rápido, sin cargar todo
    el archivo) y devuelve los nombres de columna en minúsculas."""
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, nrows=0)
        return [str(c).strip().lower() for c in df.columns]
    except Exception:
        return []


def _es_ot(path: Path) -> bool:
    cols = _columnas_hoja(path)
    return any("código o.t." in c or "codigo o.t." in c for c in cols)


def _es_ss(path: Path) -> bool:
    cols = _columnas_hoja(path)
    return any("fecha de solicitud" in c for c in cols)


def encontrar_archivos(carpeta: Path):
    """Busca los 4 archivos de entrada en `carpeta`. Disponibilidad y
    Criticidad se identifican por patrones típicos de nombre (son estables
    en el export del CMMS). OT y SS se distinguen por el nombre de archivo
    cuando es inequívoco, y si no, por sus columnas propias (el nombre de
    archivo de estos dos varía: p.ej. Windows agrega " (1)" cuando el CMMS
    exporta dos archivos con el mismo nombre por defecto)."""
    archivos = list(carpeta.glob("*.xlsx"))
    if not archivos:
        raise FileNotFoundError(
            f"No encontré ningún .xlsx en {carpeta}. "
            f"Copia ahí los 4 exports del CMMS (OT, SS, Disponibilidad, Criticidad)."
        )

    def buscar(patrones, excluir=()):
        candidatos = [
            a for a in archivos
            if any(p.lower() in a.name.lower() for p in patrones)
            and not any(e.lower() in a.name.lower() for e in excluir)
        ]
        return candidatos

    disp_candidatos = buscar(["disponibilidad"])
    crit_candidatos = buscar(["eqxip", "criticidad"])

    def elegir(nombre, candidatos):
        if len(candidatos) == 0:
            raise FileNotFoundError(
                f"No pude identificar el archivo de '{nombre}' en {carpeta}. "
                f"Archivos disponibles: {[a.name for a in archivos]}"
            )
        if len(candidatos) > 1:
            print(f"AVISO: varios archivos parecen ser '{nombre}': "
                  f"{[a.name for a in candidatos]} -> uso el más reciente")
            candidatos = sorted(candidatos, key=lambda a: a.stat().st_mtime, reverse=True)
        return candidatos[0]

    disp_path = elegir("Disponibilidad", disp_candidatos)
    crit_path = elegir("Criticidad de equipos", crit_candidatos)

    # OT y SS: de los archivos restantes, identificar por columnas propias
    # de cada export (más confiable que el nombre de archivo, que el CMMS
    # no siempre exporta igual entre corridas).
    restantes = [a for a in archivos if a not in (disp_path, crit_path)]
    ot_candidatos = [a for a in restantes if _es_ot(a)]
    ss_candidatos = [a for a in restantes if _es_ss(a)]

    ot_path = elegir("OT (Órdenes de Trabajo)", ot_candidatos)
    ss_path = elegir("SS (Solicitudes de Servicio)", [a for a in ss_candidatos if a != ot_path])

    return ot_path, ss_path, disp_path, crit_path


# ============================================================
# CARGA DE DATOS
# ============================================================

def cargar_ot(path):
    ot = pd.read_excel(path, sheet_name=0, dtype={"Código O.T.": str})
    return ot


def cargar_ss(path):
    ss = pd.read_excel(path, sheet_name=0, dtype={"Código": str})
    # hay códigos de SS duplicados en el export del CMMS (mismo número,
    # dos filas idénticas) - nos quedamos con la primera ocurrencia
    ss = ss.drop_duplicates(subset=["Código"], keep="first").reset_index(drop=True)
    return ss


def cargar_disponibilidad(path):
    # La primera fila del archivo es un encabezado de periodo
    # ("Periodo: Desde ... Hasta ..."), los nombres de columna reales
    # están en la fila 2 (header=1)
    raw = pd.read_excel(path, sheet_name="Disponibilidad", header=None, nrows=1)
    periodo_texto = str(raw.iloc[0, 0])
    disp = pd.read_excel(path, sheet_name="Disponibilidad", header=1, dtype={"Código": str})
    return disp, periodo_texto


def cargar_criticidad(path):
    crit = pd.read_excel(path, sheet_name="Novedades", dtype={"Código": str})
    return crit


# ============================================================
# RESOLUCIÓN EQUIPO <-> LUGAR
# ============================================================

def construir_indices(disp, crit):
    disp_by_code = disp.set_index("Código")
    crit_by_code = crit.set_index("Código")

    # activos "AI-xxxx | ADECUACIÓN E INSTALACIÓN <lugar>" - sirven de
    # sustituto cuando la OT/SS solo referencia una ubicación (código MTV-)
    # y no un equipo específico
    ai_rows = disp[disp["Código"].astype(str).str.startswith("AI-")]
    lugar_to_ai = ai_rows.set_index("Instalación de Proceso")

    return disp_by_code, crit_by_code, lugar_to_ai


def lugar_desde_instalacion(texto):
    """'MTV-PV-AN-AK | ARKADIA' -> 'ARKADIA'"""
    if pd.isna(texto):
        return None
    partes = str(texto).split("|", 1)
    return partes[-1].strip() if len(partes) > 1 else str(texto).strip()


def resolver_equipo(equipo_cod, entidad_full, disp_by_code, lugar_to_ai):
    """Dado el código de entidad de una OT/SS ('NR-0057' o 'MTV-PV-AN-AK'),
    devuelve (codigo_resuelto, nombre_equipo, lugar, tipo_match).

    tipo_match:
      - 'equipo_especifico': el código referencia un equipo puntual que
        existe en el maestro de disponibilidad
      - 'instalacion_pdv': el código es una ubicación (MTV-...) que se
        resuelve al activo sustituto "ADECUACIÓN E INSTALACIÓN <lugar>"
      - 'solo_ubicacion': es una ubicación pero no se encontró el activo
        sustituto correspondiente
      - 'sin_dato': no se pudo resolver nada
    """
    if equipo_cod is not None and equipo_cod in disp_by_code.index:
        row = disp_by_code.loc[equipo_cod]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        lugar = lugar_desde_instalacion(row["Instalación de Proceso"])
        return equipo_cod, row["Equipo"], lugar, "equipo_especifico"

    if isinstance(equipo_cod, str) and equipo_cod.startswith("MTV"):
        if entidad_full in lugar_to_ai.index:
            r = lugar_to_ai.loc[entidad_full]
            if isinstance(r, pd.DataFrame):
                r = r.iloc[0]
            lugar = lugar_desde_instalacion(entidad_full)
            return r.name if hasattr(r, "name") else r["Código"], r["Equipo"], lugar, "instalacion_pdv"
        else:
            lugar = lugar_desde_instalacion(entidad_full)
            return None, None, lugar, "solo_ubicacion"

    return None, None, None, "sin_dato"


# ============================================================
# TRANSFORMACIÓN OT
# ============================================================

TIPO_OT = {
    "Correctiva programada": "Correctivo",
    "Correctiva de Emergencia": "Correctivo",
    "Sistemática": "Preventivo",
}

SS_REF_RE = re.compile(r"SS-0*(\d+)")


def severidad_desde_prioridad(valor):
    if pd.isna(valor):
        return None
    partes = str(valor).split("-", 1)
    return partes[-1].strip() if len(partes) > 1 else str(valor).strip()


def extraer_ss_ref(descripcion):
    if pd.isna(descripcion):
        return None
    m = SS_REF_RE.search(str(descripcion))
    if not m:
        return None
    return m.group(1).zfill(5)


def a_iso(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(valor).isoformat()
    return str(valor)


def construir_ot(ot, ss, disp_by_code, crit_by_code, lugar_to_ai):
    ss_by_code = ss.set_index("Código")
    crit_map = crit_by_code["Criticidad"]

    registros = []
    for _, r in ot.iterrows():
        entidad_full = r["Entidades"]
        if pd.isna(entidad_full) or "|" not in str(entidad_full):
            equipo_cod, equipo_nombre_raw = None, str(entidad_full) if pd.notna(entidad_full) else None
        else:
            equipo_cod, equipo_nombre_raw = [p.strip() for p in str(entidad_full).split("|", 1)]

        cod_resuelto, nombre_resuelto, lugar, tipo_match = resolver_equipo(
            equipo_cod, entidad_full, disp_by_code, lugar_to_ai
        )
        if nombre_resuelto is None:
            nombre_resuelto = equipo_nombre_raw

        criticidad = crit_map.get(cod_resuelto) if cod_resuelto else None

        ss_ref = extraer_ss_ref(r["Descripción"])
        viene_ss = ss_ref is not None and ss_ref in ss_by_code.index
        fecha_creacion_ss = None
        if viene_ss:
            fila_ss = ss_by_code.loc[ss_ref]
            if isinstance(fila_ss, pd.DataFrame):
                fila_ss = fila_ss.iloc[0]
            fecha_creacion_ss = a_iso(fila_ss["Fecha de solicitud"])

        costo_real = r["Total Real"]
        if pd.isna(costo_real):
            costo_real = 0.0

        registros.append({
            "ot": r["Código O.T."],
            "tipo": TIPO_OT.get(r["Tipo"], r["Tipo"]),
            "severidad": severidad_desde_prioridad(r["Prioridad"]),
            "viene_ss": bool(viene_ss),
            "ss_codigo": ss_ref if viene_ss else None,
            "fecha_creacion_ss": fecha_creacion_ss,
            "fecha_creacion": a_iso(r["Fecha Creación"]),
            "inicio_programado": a_iso(r["Fecha Inicio Programado"]),
            "fin_programado": a_iso(r["Fecha Fin Programado"]),
            "inicio_real": a_iso(r["Fecha Inicio Real"]),
            "fin_real": a_iso(r["Fecha Fin Real"]),
            "estado": r["Estado O.T."],
            "avance": None if pd.isna(r["Ejecución (%)"]) else float(r["Ejecución (%)"]),
            "lugar": lugar,
            "equipo": nombre_resuelto,
            "equipo_cod": cod_resuelto,
            "criticidad": criticidad,
            "tecnico": None if pd.isna(r["Ejecutores"]) else str(r["Ejecutores"]),
            "costo_real": float(costo_real),
            "_tipo_match": tipo_match,  # uso interno para construir "paros", no va al tablero
        })
    return pd.DataFrame(registros)


# ============================================================
# TRANSFORMACIÓN SS
# ============================================================

def codigo_ot_desde_id(ot_id, id_a_codigo):
    """Traduce el Id interno del CMMS (columna 'OTs' del export de SS) al
    Código O.T. visible, usando la columna 'Id' del archivo de OT. Si el Id
    no aparece en el export de OT actual (por ejemplo, quedó fuera del rango
    exportado), se muestra 'No disponible' en vez del número interno."""
    if pd.isna(ot_id):
        return None
    codigo = id_a_codigo.get(int(ot_id))
    if codigo is None or (isinstance(codigo, float) and pd.isna(codigo)):
        return "No disponible"
    return str(codigo)


def construir_ss(ss, disp_by_code, lugar_to_ai, ot_raw):
    # "OTs" en el export de SS trae el Id interno del CMMS (no el Código
    # O.T. visible) - se traduce con la columna "Id" del archivo de OT.
    id_a_codigo = ot_raw.set_index("Id")["Código O.T."]

    registros = []
    for _, r in ss.iterrows():
        entidad_full = r["Entidad"]
        if pd.isna(entidad_full) or "|" not in str(entidad_full):
            equipo_cod, equipo_nombre_raw = None, str(entidad_full) if pd.notna(entidad_full) else None
        else:
            equipo_cod, equipo_nombre_raw = [p.strip() for p in str(entidad_full).split("|", 1)]

        _, nombre_resuelto, lugar, _ = resolver_equipo(equipo_cod, entidad_full, disp_by_code, lugar_to_ai)
        if nombre_resuelto is None:
            nombre_resuelto = equipo_nombre_raw

        registros.append({
            "ss": r["Código"],
            "tipo": r["Tipo"],
            "severidad": r["Prioridad"],
            "equipo": nombre_resuelto,
            "lugar": lugar,
            "fecha_ss": a_iso(r["Fecha de solicitud"]),
            "fecha_respuesta": a_iso(r["Fecha de respuesta"]),
            "estado": r["Estado"],
            "ot_asociada": codigo_ot_desde_id(r["OTs"], id_a_codigo),
        })
    return pd.DataFrame(registros)


# ============================================================
# PAROS (a partir de OT correctivas resueltas a un equipo/lugar)
# ============================================================

def construir_paros(ot_df, ot_raw):
    ot_raw_idx = ot_raw.set_index("Código O.T.")
    correctivas = ot_df[
        (ot_df["tipo"] == "Correctivo")
        & (ot_df["_tipo_match"].isin(["equipo_especifico", "instalacion_pdv"]))
    ]

    registros = []
    for _, r in correctivas.iterrows():
        fila_raw = ot_raw_idx.loc[r["ot"]]
        if isinstance(fila_raw, pd.DataFrame):
            fila_raw = fila_raw.iloc[0]

        inicio_real = fila_raw["Fecha Inicio Real"]
        inicio_prog = fila_raw["Fecha Inicio Programado"]
        fin_real = fila_raw["Fecha Fin Real"]

        iniciado = pd.notna(inicio_real)
        inicio_efectivo = inicio_real if iniciado else inicio_prog

        dur = None
        if iniciado and pd.notna(fin_real):
            dur = round((fin_real - inicio_real).total_seconds() / 3600, 2)

        registros.append({
            "ot": r["ot"],
            "equipo": r["equipo"],
            "equipo_cod": r["equipo_cod"],
            "lugar": r["lugar"],
            "criticidad": r["criticidad"],
            "inicio": a_iso(inicio_efectivo),
            "iniciado": bool(iniciado),
            "fin": a_iso(fin_real),
            "duracion_h": dur,
            "tecnico": r["tecnico"],
            "estado": r["estado"],
        })
    return registros


# ============================================================
# EQUIPOS (directorio de disponibilidad/criticidad)
# ============================================================

def construir_equipos(disp, crit):
    crit_idx = crit.set_index("Código")
    registros = []
    total_maestro = len(disp)
    for _, r in disp.iterrows():
        instalacion = r["Instalación de Proceso"]
        if isinstance(instalacion, str) and instalacion.startswith("ANM"):
            continue  # activo no mantenible - se excluye del directorio operativo
        criticidad = None
        if r["Código"] in crit_idx.index:
            fila_crit = crit_idx.loc[r["Código"]]
            if isinstance(fila_crit, pd.DataFrame):
                fila_crit = fila_crit.iloc[0]
            criticidad = fila_crit["Criticidad"]

        disponibilidad_pct = r["Disponibilidad [%]"]
        registros.append({
            "codigo": r["Código"],
            "equipo": r["Equipo"],
            "lugar": lugar_desde_instalacion(instalacion),
            "criticidad": criticidad,
            "disponibilidad_pct": None if pd.isna(disponibilidad_pct) else float(disponibilidad_pct),
            "horas_paro_correctivo": float(r["Tiempo Paro Correctivo [Horas]"]) if pd.notna(r["Tiempo Paro Correctivo [Horas]"]) else 0.0,
        })
    return registros, total_maestro


def construir_criticidad_totales(crit):
    return crit["Criticidad"].value_counts().to_dict()


# ============================================================
# ENSAMBLAJE FINAL
# ============================================================

def construir_data_json(ot_path, ss_path, disp_path, crit_path):
    ot_raw = cargar_ot(ot_path)
    ss_raw = cargar_ss(ss_path)
    disp, periodo_texto = cargar_disponibilidad(disp_path)
    crit = cargar_criticidad(crit_path)

    disp_by_code, crit_by_code, lugar_to_ai = construir_indices(disp, crit)

    ot_df = construir_ot(ot_raw, ss_raw, disp_by_code, crit_by_code, lugar_to_ai)
    ss_df = construir_ss(ss_raw, disp_by_code, lugar_to_ai, ot_raw)
    paros = construir_paros(ot_df, ot_raw)
    equipos, total_maestro = construir_equipos(disp, crit)
    criticidad_totales = construir_criticidad_totales(crit)

    ot_records = ot_df.drop(columns=["_tipo_match"]).to_dict(orient="records")

    data = {
        "generado": datetime.now().isoformat(timespec="seconds"),
        "periodo_datos": periodo_texto,
        "ot": ot_records,
        "ss": ss_df.to_dict(orient="records"),
        "paros": paros,
        "equipos": equipos,
        "criticidad_totales": criticidad_totales,
        "total_equipos_maestro": total_maestro,
    }

    # reporte rápido en consola para poder revisar antes de subir
    resueltas = ot_df["_tipo_match"].isin(["equipo_especifico", "instalacion_pdv"]).sum()
    print(f"OT procesadas: {len(ot_df)} (de las cuales {resueltas} con equipo/lugar resuelto)")
    print(f"SS procesadas: {len(ss_df)}")
    print(f"Paros identificados: {len(paros)} "
          f"({sum(1 for p in paros if p['iniciado'])} iniciados, "
          f"{sum(1 for p in paros if not p['iniciado'])} pendientes de iniciar)")
    print(f"Equipos en directorio (mantenibles): {len(equipos)} de {total_maestro} en el maestro")
    print(f"Criticidad: {criticidad_totales}")

    return data


def limpiar_nan(obj):
    """Reemplaza recursivamente cualquier NaN/NaT flotante por None,
    para que el resultado sea JSON válido (json.dumps no lo hace solo)."""
    if isinstance(obj, dict):
        return {k: limpiar_nan(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [limpiar_nan(v) for v in obj]
    if isinstance(obj, float) and np.isnan(obj):
        return None
    try:
        if pd.isna(obj):
            return None
    except (TypeError, ValueError):
        pass
    return obj


def inyectar_en_plantilla(data, plantilla_path, salida_path):
    plantilla = plantilla_path.read_text(encoding="utf-8")
    data_json_str = json.dumps(data, ensure_ascii=False)
    if "__DATA_JSON__" not in plantilla:
        raise ValueError(
            f"La plantilla {plantilla_path} no contiene el marcador __DATA_JSON__ "
            f"- no se puede inyectar los datos."
        )
    final = plantilla.replace("__DATA_JSON__", data_json_str)
    salida_path.write_text(final, encoding="utf-8")
    print(f"Tablero generado: {salida_path} ({len(final)/1024:.0f} KB)")


def main():
    if len(sys.argv) == 5:
        ot_path, ss_path, disp_path, crit_path = (Path(p) for p in sys.argv[1:5])
    else:
        carpeta = CARPETA_DATOS if any(CARPETA_DATOS.glob("*.xlsx")) else CARPETA_BASE
        ot_path, ss_path, disp_path, crit_path = encontrar_archivos(carpeta)

    print(f"OT:           {ot_path.name}")
    print(f"SS:           {ss_path.name}")
    print(f"Disponibilidad: {disp_path.name}")
    print(f"Criticidad:   {crit_path.name}")
    print()

    data = construir_data_json(ot_path, ss_path, disp_path, crit_path)
    data = limpiar_nan(data)

    SALIDA_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\ndata.json guardado en {SALIDA_JSON}")

    if PLANTILLA_HTML.exists():
        inyectar_en_plantilla(data, PLANTILLA_HTML, SALIDA_HTML)
    else:
        print(f"\nAVISO: no encontré la plantilla {PLANTILLA_HTML} - "
              f"solo se generó data.json, no el index.html final.")


if __name__ == "__main__":
    main()
