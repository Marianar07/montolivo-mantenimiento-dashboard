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
    # excluye archivos temporales de bloqueo de Excel (p.ej. "~$TECNICOS.xlsx"
    # que Excel crea mientras el archivo real está abierto)
    archivos = [a for a in carpeta.glob("*.xlsx") if not a.name.startswith("~$")]
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
    crit_candidatos = buscar(["datos generales de equipos"])

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
    tecnicos_candidatos = buscar(["tecnicos"])
    tecnicos_path = elegir("TECNICOS", tecnicos_candidatos) if tecnicos_candidatos else None

    # OT y SS: de los archivos restantes, identificar por columnas propias
    # de cada export (más confiable que el nombre de archivo, que el CMMS
    # no siempre exporta igual entre corridas).
    excluidos = {disp_path, crit_path} | ({tecnicos_path} if tecnicos_path else set())
    restantes = [a for a in archivos if a not in excluidos]
    ot_candidatos = [a for a in restantes if _es_ot(a)]
    ss_candidatos = [a for a in restantes if _es_ss(a)]

    ot_path = elegir("OT (Órdenes de Trabajo)", ot_candidatos)
    ss_path = elegir("SS (Solicitudes de Servicio)", [a for a in ss_candidatos if a != ot_path])

    return ot_path, ss_path, disp_path, crit_path, tecnicos_path


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
    """'Datos Generales de Equipos.xlsx' - maestro consolidado de activos,
    hoja única 'Sheet'. Trae, entre otras, las columnas 'Código',
    'Criticidad' (Alta/Media/Baja) y 'Provoca Paro?' (Sí/No) - esta última
    es la clasificación oficial del CMMS que usa construir_paros() para
    decidir qué correctivos cuentan como paro (ver nota ahí)."""
    crit = pd.read_excel(path, sheet_name="Sheet", dtype={"Código": str})
    return crit


# El campo Ejecutores de las OT a veces trae el nombre completo del técnico
# (con más apellidos/nombres) mientras que TECNICOS.xlsx (columna NOMBRE)
# trae una versión más corta del mismo nombre. Alias conocidos hoy - si
# aparece un técnico interno nuevo cuyo nombre en OT no calza exacto con
# TECNICOS.xlsx, quedará mal clasificado como "tercero" hasta agregarlo
# aquí (revisar el aviso que imprime el script).
ALIAS_TECNICOS = {
    "ALEJANDRO CARDONA CARMONA": "ALEJANDRO DE JESUS CARDONA CARMONA",
    "JEFFERSON ANDRES OSORIO": "JEFFERSON OSORIO BUSTAMANTE",
    "OSCAR DARIO CASTAÑEDA": "OSCAR DARIO CASTAÑEDA GARAY",
}


def cargar_tecnicos(path):
    """TECNICOS.xlsx - maestro de técnicos internos (planta), hoja única,
    columnas NOMBRE/CEDULA/CIUDAD. Un técnico que aparece en el campo
    Ejecutores de una OT pero NO está en este maestro (ni en ALIAS_TECNICOS)
    se trata como trabajo de tercero en la pestaña "Técnicos" - ver
    renderTecnicos() en index_template.html y la nota en CLAUDE.md."""
    df = pd.read_excel(path, sheet_name=0, dtype=str)
    nombres = [str(n).strip() for n in df["NOMBRE"].dropna()]
    # normaliza al nombre "largo" (como aparece en Ejecutores) vía alias
    return sorted({ALIAS_TECNICOS.get(n, n) for n in nombres})


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
            # NOTA: usar r["Código"] (la columna), no r.name - lugar_to_ai
            # esta indexado por "Instalación de Proceso", así que r.name es
            # el texto del lugar, no el código AI-xxxx del activo sustituto
            # (bug corregido: antes esto dejaba equipo_cod con el texto
            # crudo "MTV-... | Lugar" en vez del código real, y por eso
            # tanto la criticidad como el cruce con 'Provoca Paro?'
            # fallaban en silencio para estas filas).
            return r["Código"], r["Equipo"], lugar, "instalacion_pdv"
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

def construir_ss(ss, disp_by_code, lugar_to_ai, ot_df):
    """NOTA IMPORTANTE (verificado con datos reales, no borrar este
    comentario): la columna cruda "OTs" del export de SS NO es confiable
    para saber qué OT resolvió una solicitud. Ejemplo real: la SS 00004
    trae OTs=8; la OT con Id=8 es la 000005 (un horno, equipo totalmente
    distinto). La OT que en realidad resolvió la SS 00004 es la 000008 -
    se sabe porque es sobre el mismo equipo (Sistema de Extracción) y su
    Descripción cita textualmente "SS-00004". O sea, ni el Id ni ningún
    otro campo directo de "OTs" apunta de forma confiable a la OT correcta.
    En su lugar, usamos la referencia de texto "SS-XXXXX" que ya se extrae
    de la Descripción de cada OT (columna `ss_codigo` de `ot_df`, calculada
    en `construir_ot`) y armamos el cruce en la dirección contraria
    (SS -> OT). NO reintroducir un mapeo basado en la columna "OTs" cruda.
    """
    ot_por_ss = {r["ss_codigo"]: r["ot"] for r in ot_df.to_dict(orient="records") if r["viene_ss"]}

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
            "fecha_esperada": a_iso(r["Fecha esperada"]),
            "fecha_respuesta": a_iso(r["Fecha de respuesta"]),
            "estado": r["Estado"],
            "ot_asociada": ot_por_ss.get(r["Código"]),
        })
    return pd.DataFrame(registros)


# ============================================================
# PAROS (a partir de OT correctivas resueltas a un equipo/lugar)
# ============================================================

def construir_paros(ot_df, ot_raw, crit, tipo="Correctivo"):
    """Una OT cuenta como paro (o parada programada, si tipo='Preventivo')
    solo si su equipo_cod resuelto está marcado 'Provoca Paro?' = 'Sí' en
    'Datos Generales de Equipos.xlsx' - la clasificación oficial del CMMS
    (columna propia del maestro de activos, cargado en `crit`). Ya NO se usa
    la marca ANM ni el prefijo AI- para esta decisión (esos criterios eran
    una aproximación de cuando no existía esta columna; ver CLAUDE.md).

    tipo="Correctivo" (default) -> paros reales no programados.
    tipo="Preventivo" -> paradas programadas (mismo filtro de equipo, para
    el mantenimiento planeado sobre equipos que sí generan paro)."""
    provoca_paro_set = set(crit[crit["Provoca Paro?"] == "Sí"]["Código"])

    ot_raw_idx = ot_raw.set_index("Código O.T.")
    filtradas = ot_df[
        (ot_df["tipo"] == tipo)
        & (ot_df["equipo_cod"].isin(provoca_paro_set))
    ]

    registros = []
    for _, r in filtradas.iterrows():
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
            "ss_codigo": r["ss_codigo"] if r["viene_ss"] else None,
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
        mantenible = not (isinstance(instalacion, str) and instalacion.startswith("ANM"))
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
            "mantenible": mantenible,
            "disponibilidad_pct": None if pd.isna(disponibilidad_pct) else float(disponibilidad_pct),
            "horas_paro_correctivo": float(r["Tiempo Paro Correctivo [Horas]"]) if pd.notna(r["Tiempo Paro Correctivo [Horas]"]) else 0.0,
        })
    return registros, total_maestro


def construir_criticidad_totales(crit):
    return crit["Criticidad"].value_counts().to_dict()


# ============================================================
# ENSAMBLAJE FINAL
# ============================================================

def construir_data_json(ot_path, ss_path, disp_path, crit_path, tecnicos_path=None):
    ot_raw = cargar_ot(ot_path)
    ss_raw = cargar_ss(ss_path)
    disp, periodo_texto = cargar_disponibilidad(disp_path)
    crit = cargar_criticidad(crit_path)
    tecnicos_internos = cargar_tecnicos(tecnicos_path) if tecnicos_path else None

    disp_by_code, crit_by_code, lugar_to_ai = construir_indices(disp, crit)

    ot_df = construir_ot(ot_raw, ss_raw, disp_by_code, crit_by_code, lugar_to_ai)
    ss_df = construir_ss(ss_raw, disp_by_code, lugar_to_ai, ot_df)
    paros = construir_paros(ot_df, ot_raw, crit, tipo="Correctivo")
    paros_programados = construir_paros(ot_df, ot_raw, crit, tipo="Preventivo")
    equipos, total_maestro = construir_equipos(disp, crit)
    criticidad_totales = construir_criticidad_totales(crit)

    ot_records = ot_df.drop(columns=["_tipo_match"]).to_dict(orient="records")

    data = {
        "generado": datetime.now().isoformat(timespec="seconds"),
        "periodo_datos": periodo_texto,
        "ot": ot_records,
        "ss": ss_df.to_dict(orient="records"),
        "paros": paros,
        "paros_programados": paros_programados,
        "equipos": equipos,
        "criticidad_totales": criticidad_totales,
        "total_equipos_maestro": total_maestro,
        # lista de nombres de técnicos internos (planta), normalizados a como
        # aparecen en el campo Ejecutores de las OT (ver ALIAS_TECNICOS). Si
        # es None, no se encontró TECNICOS.xlsx - el tablero no separa
        # terceros en ese caso (trata a todos como internos).
        "tecnicos_internos": tecnicos_internos,
    }

    # reporte rápido en consola para poder revisar antes de subir
    resueltas = ot_df["_tipo_match"].isin(["equipo_especifico", "instalacion_pdv"]).sum()
    print(f"OT procesadas: {len(ot_df)} (de las cuales {resueltas} con equipo/lugar resuelto)")
    print(f"SS procesadas: {len(ss_df)}")
    provoca_paro_n = int((crit["Provoca Paro?"] == "Sí").sum())
    print(f"Activos que 'Provoca Paro?' = Sí: {provoca_paro_n} de {len(crit)} en el maestro")
    print(f"Paros identificados: {len(paros)} "
          f"({sum(1 for p in paros if p['iniciado'])} iniciados, "
          f"{sum(1 for p in paros if not p['iniciado'])} pendientes de iniciar)")
    horas_prog = sum(p['duracion_h'] or 0 for p in paros_programados)
    print(f"Paradas programadas (preventivo, equipos que generan paro): {len(paros_programados)} ({horas_prog:.1f} h)")
    mantenibles = sum(1 for e in equipos if e["mantenible"])
    print(f"Equipos en directorio: {len(equipos)} de {total_maestro} en el maestro ({mantenibles} mantenibles, {len(equipos)-mantenibles} no mantenibles)")
    print(f"Criticidad: {criticidad_totales}")
    if tecnicos_internos is not None:
        nombres_en_ot = {
            n.strip() for r in ot_records if isinstance(r.get("tecnico"), str)
            for n in r["tecnico"].split(",")
        }
        terceros = sorted(nombres_en_ot - set(tecnicos_internos))
        print(f"Técnicos internos (TECNICOS.xlsx): {len(tecnicos_internos)}")
        print(f"Nombres en OT no reconocidos como internos (van como tercero): {terceros}")

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
        tecnicos_path = None
        for carpeta in (CARPETA_DATOS, CARPETA_BASE):
            candidatos = [a for a in carpeta.glob("*.xlsx") if "tecnicos" in a.name.lower()]
            if candidatos:
                tecnicos_path = sorted(candidatos, key=lambda a: a.stat().st_mtime, reverse=True)[0]
                break
    else:
        carpeta = CARPETA_DATOS if any(CARPETA_DATOS.glob("*.xlsx")) else CARPETA_BASE
        ot_path, ss_path, disp_path, crit_path, tecnicos_path = encontrar_archivos(carpeta)

    print(f"OT:           {ot_path.name}")
    print(f"SS:           {ss_path.name}")
    print(f"Disponibilidad: {disp_path.name}")
    print(f"Datos Generales de Equipos: {crit_path.name}")
    if tecnicos_path:
        print(f"Técnicos:     {tecnicos_path.name}")
    else:
        print("Técnicos:     no encontrado (TECNICOS.xlsx) - la pestaña Técnicos no separará terceros")
    print()

    data = construir_data_json(ot_path, ss_path, disp_path, crit_path, tecnicos_path)
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
