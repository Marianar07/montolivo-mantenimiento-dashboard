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

from collections import Counter
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
    proveedores_candidatos = buscar(["proveedores"])
    proveedores_path = elegir("Datos Generales de Proveedores", proveedores_candidatos) if proveedores_candidatos else None
    activos_candidatos = buscar(["datos generales activos"])
    activos_path = elegir("Datos Generales Activos", activos_candidatos) if activos_candidatos else None
    paros_candidatos = buscar(["paros"])
    paros_path = elegir("Informe de paros", paros_candidatos) if paros_candidatos else None

    # OT y SS: de los archivos restantes, identificar por columnas propias
    # de cada export (más confiable que el nombre de archivo, que el CMMS
    # no siempre exporta igual entre corridas).
    excluidos = {disp_path, crit_path}
    if activos_path:
        excluidos.add(activos_path)
    if tecnicos_path:
        excluidos.add(tecnicos_path)
    if proveedores_path:
        excluidos.add(proveedores_path)
    if paros_path:
        excluidos.add(paros_path)  # trae "Código O.T.": no confundirlo con el export de OT
    restantes = [a for a in archivos if a not in excluidos]
    ot_candidatos = [a for a in restantes if _es_ot(a)]
    ss_candidatos = [a for a in restantes if _es_ss(a)]

    ot_path = elegir("OT (Órdenes de Trabajo)", ot_candidatos)
    ss_path = elegir("SS (Solicitudes de Servicio)", [a for a in ss_candidatos if a != ot_path])

    return ot_path, ss_path, disp_path, crit_path, tecnicos_path, proveedores_path, activos_path, paros_path


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
    se usa en el reporte de consola (los paros salen del informe de paros
    de Mantum, ver cargar_paros)."""
    crit = pd.read_excel(path, sheet_name="Sheet", dtype={"Código": str})
    return crit


def integrar_activos(path, disp, crit):
    """'Datos Generales Activos.xlsx' (opcional, agregado por Mariana el
    1-oct-2026 para tener los equipos nuevos) - export de ACTIVOS del CMMS,
    más actualizado que Disponibilidad y que Datos Generales de Equipos, pero
    con otro formato: columnas 'Codigo' (sin tilde), 'Nombre', 'Código IP.',
    'Instalación Proceso', 'Provoca Paro' (sin '?') y SIN 'Criticidad'.

    Se usa como maestro de activos:
      - agrega a `disp` los activos que no están en Disponibilidad (sin datos
        de disponibilidad - ese módulo igual reporta 100% para todos), para
        que cuenten en el total de equipos, salgan en el directorio y las OT/SS
        sobre ellos se resuelvan a equipo/lugar;
      - actualiza la 'Instalación de Proceso' de los que ya estaban (equipos
        trasladados de punto de venta) y su nombre;
      - en `crit`, toma 'Provoca Paro?' de este archivo y conserva la
        'Criticidad' de Datos Generales de Equipos; los activos nuevos quedan
        sin criticidad (None) hasta que llegue un export que la traiga.
    Devuelve (disp, crit, resumen) con `resumen` para el reporte de consola."""
    if path is None:
        # Sin 'Datos Generales Activos.xlsx': el export de Datos Generales de
        # Equipos (5-oct-2026) ya trae los mismos activos y su instalación de
        # proceso, así que hace de maestro con las columnas renombradas.
        act = crit.rename(columns={
            "Código": "Codigo",
            "Código Instalación de Proceso": "Código IP.",
            "Nombre Instalación de Proceso": "Instalación Proceso",
            "Provoca Paro?": "Provoca Paro",
        })[["Codigo", "Nombre", "Código IP.", "Instalación Proceso", "Provoca Paro"]].copy()
        act["Codigo"] = act["Codigo"].astype("string")
        act["Código IP."] = act["Código IP."].astype("string")
    else:
        act = pd.read_excel(path, sheet_name=0, dtype={"Codigo": str, "Código IP.": str})
    act = act[act["Codigo"].notna()].drop_duplicates("Codigo")
    act["Codigo"] = act["Codigo"].str.strip()
    instalacion = act["Código IP."].fillna("").str.strip() + " | " + act["Instalación Proceso"].fillna("").str.strip()
    act_inst = dict(zip(act["Codigo"], instalacion))
    act_nombre = dict(zip(act["Codigo"], act["Nombre"]))

    # Disponibilidad: actualizar ubicación/nombre y agregar los nuevos
    disp = disp.copy()
    antes = disp["Instalación de Proceso"].copy()
    en_act = disp["Código"].isin(act_inst)
    disp.loc[en_act, "Instalación de Proceso"] = disp.loc[en_act, "Código"].map(act_inst)
    disp.loc[en_act, "Equipo"] = disp.loc[en_act, "Código"].map(act_nombre)
    trasladados = int((antes[en_act] != disp.loc[en_act, "Instalación de Proceso"]).sum())
    nuevos = act[~act["Codigo"].isin(set(disp["Código"]))]
    if len(nuevos):
        filas = pd.DataFrame({
            "Código": nuevos["Codigo"].values,
            "Equipo": nuevos["Nombre"].values,
            "Instalación de Proceso": [act_inst[c] for c in nuevos["Codigo"]],
        })
        for col in disp.columns:
            if col not in filas.columns:
                filas[col] = 0.0 if "Paro" in col else pd.NA
        disp = pd.concat([disp, filas[disp.columns]], ignore_index=True)

    # Maestro de criticidad / provoca paro
    crit_old = crit.drop_duplicates("Código").set_index("Código")
    crit_new = pd.DataFrame({
        "Código": act["Codigo"].values,
        "Nombre": act["Nombre"].values,
        "Provoca Paro?": act["Provoca Paro"].astype(str).str.strip().values,
    })
    crit_new["Criticidad"] = crit_new["Código"].map(crit_old["Criticidad"]) if "Criticidad" in crit_old else None
    # activos que solo estén en el maestro viejo (no debería pasar) se conservan
    solo_viejo = crit[~crit["Código"].isin(set(crit_new["Código"]))][["Código", "Nombre", "Provoca Paro?", "Criticidad"]]
    crit = pd.concat([crit_new, solo_viejo], ignore_index=True)

    resumen = {
        "activos": len(act),
        "nuevos": nuevos[["Codigo", "Nombre", "Instalación Proceso", "Provoca Paro"]].values.tolist(),
        "trasladados": trasladados,
        "sin_criticidad": int(crit["Criticidad"].isna().sum()),
    }
    return disp, crit, resumen


# El campo Ejecutores de las OT a veces trae el nombre completo del técnico
# (con más apellidos/nombres) mientras que TECNICOS.xlsx (columna NOMBRE)
# trae una versión más corta del mismo nombre. Alias conocidos hoy - si
# aparece un técnico interno nuevo cuyo nombre en OT no calza exacto con
# TECNICOS.xlsx, quedará mal clasificado como "tercero" hasta agregarlo
# aquí (revisar el aviso que imprime el script).
ALIAS_TECNICOS = {
    "ALEJANDRO CARDONA CARMONA": "ALEJANDRO DE JESUS CARDONA CARMONA",
    # 6-oct-2026: las OT ya traen el nombre completo (como en TECNICOS.xlsx)
    "JEFFERSON ANDRES OSORIO": "JEFFERSON ANDRES OSORIO BUSTAMANTE",
    "JEFFERSON OSORIO BUSTAMANTE": "JEFFERSON ANDRES OSORIO BUSTAMANTE",
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


def cargar_tecnicos_inactivos(path):
    """Técnicos de TECNICOS.xlsx con ESTADO = INACTIVO (ya no trabajan en la
    compañía). Siguen siendo internos - su trabajo pasado se muestra en
    "Capacidad y ocupación" - pero el tablero los saca de "Ubicación del
    técnico" y del conteo de técnicos en sitio. Si el archivo no trae la
    columna ESTADO, nadie se considera inactivo."""
    df = pd.read_excel(path, sheet_name=0, dtype=str)
    if "ESTADO" not in df.columns:
        return []
    inactivos = df[df["ESTADO"].astype(str).str.strip().str.upper() == "INACTIVO"]
    nombres = [str(n).strip() for n in inactivos["NOMBRE"].dropna()]
    return sorted({ALIAS_TECNICOS.get(n, n) for n in nombres})


# Personas que ejecutan OT a nombre de un proveedor sin estar dadas de alta
# individualmente en "Datos Generales de Proveedores" - decisión de Mariana
# (29-sep-2026): se cuentan como proveedor, no como "otro". Si aparece otra
# persona así en el aviso de consola, agregarla aquí.
PROVEEDORES_EXTRA = [
    "SEBASTIAN BUITRAGO GRACIANO",
]


def cargar_proveedores(path):
    """Datos Generales de Proveedores.xlsx - maestro de proveedores/terceros
    del CMMS, hoja única 'Sheet', columna 'Nombre'. El tablero clasifica un
    ejecutor no interno como "proveedor" si está en este listado (más
    PROVEEDORES_EXTRA, ver arriba) y como "otro" si no - ver
    esOTProveedor/esOTOtro en index_template.html."""
    df = pd.read_excel(path, sheet_name=0, dtype=str)
    return sorted({str(n).strip() for n in df["Nombre"].dropna()})


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

# La Realimentación de una OT repite, por cada visita, un bloque
# "Ejecutores:\n<NOMBRE>\n..." - ese nombre ya está en la columna Técnico,
# así que se recorta del texto de comentarios para no duplicarlo.
EJECUTORES_RE = re.compile(r"Ejecutores:\n[^\n]*\n")


def limpiar_comentarios(texto):
    if pd.isna(texto):
        return None
    limpio = EJECUTORES_RE.sub("", str(texto))
    limpio = re.sub(r"\n{3,}", "\n\n", limpio).strip()
    return limpio or None


# La columna "Comentarios" del export de SS trae el historial de estados
# ("Estado / Persona / fecha / texto"), más reciente primero, separados por
# el texto literal "SEPARADOR-COMENTARIOS" -> uno por línea.
def limpiar_comentarios_ss(texto):
    if pd.isna(texto):
        return None
    partes = [p.strip() for p in str(texto).split("SEPARADOR-COMENTARIOS") if p.strip()]
    return "\n".join(partes) or None


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
            "comentarios": limpiar_comentarios(r["Realimentación"]),
            "_tipo_match": tipo_match,  # uso interno para construir "paros", no va al tablero
            "_actividad": r.get("Actividades"),  # uso interno: consolidar el costo por actividad
        })
    df = pd.DataFrame(registros)
    # El export de OT trae una fila por cada ACTIVIDAD de la OT (columna
    # "Actividades"), cada una con su propio "Total Real" - el resto de campos
    # se repite. Se consolida a una fila por OT sumando el costo de todas sus
    # actividades (confirmado con Mariana contra Mantum: OT 000022 = 128.242,85
    # y OT 000182 = 243.749,90, ambas = suma de sus filas).
    # Además trae una fila por cada RECURSO asignado (columnas "Código
    # recurso"/"Cantidad real"), repitiendo el mismo "Total Real" de la
    # actividad en cada una (5-oct-2026: OT 000040 = 39 filas de 24.318,47).
    # Por eso se toma un solo Total Real por (OT, actividad) antes de sumar;
    # sumar todas las filas multiplicaba el costo por el número de recursos.
    costo_por_ot = (df.drop_duplicates(subset=["ot", "_actividad"], keep="first")
                      .groupby("ot", sort=False)["costo_real"].sum())
    df = df.drop_duplicates(subset=["ot"], keep="first").copy()
    df["costo_real"] = df["ot"].map(costo_por_ot)
    df = df.drop(columns=["_actividad"])
    return df.reset_index(drop=True)


# ============================================================
# TRANSFORMACIÓN SS
# ============================================================

def construir_ss(ss, disp_by_code, lugar_to_ai, ot_df):
    """`ot_asociada` se arma con la referencia de texto "SS-XXXXX" que se
    extrae de la Descripción de cada OT (columna `ss_codigo` de `ot_df`).

    Nota (re-verificado el 28-sep-2026): la columna cruda "OTs" del export
    de SS SÍ trae el CÓDIGO de la OT (número sin ceros a la izquierda, p.ej.
    8 -> OT 000008), no su Id interno: en las 92 SS que la traen, el equipo
    de la SS coincide con el de la OT. Un comentario anterior decía lo
    contrario (confundía ese número con la columna Id de la OT).
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
            "descripcion": r["Descripción"],
            "comentarios": limpiar_comentarios_ss(r.get("Comentarios")),
        })
    return pd.DataFrame(registros)


# ============================================================
# PAROS (informe de paros de Mantum)
# ============================================================

def cargar_paros(path):
    """'Datos generales de paros de equipos discriminados por O.T..xlsx' -
    informe de paros de Mantum (hoja única 'Sheet'). Es la fuente oficial
    de paros (Mariana, 8-oct-2026): solo cuentan los paros que aparecen
    aquí; ya no se derivan de SS ni de OT."""
    return pd.read_excel(path, sheet_name=0, dtype={"Código O.T.": str})


def construir_paros(paros_raw, ot_df, crit, disp_by_code, lugar_to_ai):
    """Un registro por fila del informe de paros. Devuelve (correctivos,
    preventivos) según el 'Tipo O.T.' de la OT del paro (Preventivo ->
    preventivos; cualquier otro tipo -> correctivos).
    Estado: 'Finalizado' si trae 'Fecha fin paro'; si no, 'En reparación'
    cuando la OT ya tiene Fecha Inicio Real, o 'Fuera de servicio'."""
    if paros_raw is None:
        return [], []
    crit_map = crit.drop_duplicates("Código").set_index("Código")["Criticidad"]
    ot_idx = {r["ot"]: r for r in ot_df.to_dict(orient="records")}

    def ts(v):
        return pd.Timestamp(v) if v is not None and pd.notna(v) else None

    correctivos, preventivos = [], []
    for _, r in paros_raw.iterrows():
        entidad = r["Equipo"]
        if pd.isna(entidad):
            continue
        if "|" in str(entidad):
            cod, nombre = [p.strip() for p in str(entidad).split("|", 1)]
        else:
            cod, nombre = str(entidad).strip(), str(entidad).strip()
        _, equipo, lugar, _ = resolver_equipo(cod, entidad, disp_by_code, lugar_to_ai)

        ot_cod = r.get("Código O.T.")
        ot_cod = str(ot_cod).strip().zfill(6) if pd.notna(ot_cod) and str(ot_cod).strip() else None
        o = ot_idx.get(ot_cod)
        if lugar is None and o is not None:
            lugar = o["lugar"]

        inicio, fin = ts(r["Fecha inicio paro"]), ts(r["Fecha fin paro"])
        inicio_trabajo = ts(o["inicio_real"]) if o is not None else ts(r.get("Fecha incio O.T."))
        if fin is not None:
            estado_paro = "Finalizado"
        elif inicio_trabajo is not None:
            estado_paro = "En reparación"
        else:
            estado_paro = "Fuera de servicio"

        motivo = " · ".join(str(x) for x in (r.get("Parado por"), r.get("Tipo paro"), r.get("Descripción"))
                            if pd.notna(x) and str(x).strip())
        registro = {
            "id_paro": int(r["Id"]) if pd.notna(r.get("Id")) else None,
            "ot": ot_cod,
            "equipo": equipo or nombre,
            "equipo_cod": cod,
            "lugar": lugar,
            "criticidad": crit_map.get(cod),
            "inicio": a_iso(inicio),
            "inicio_trabajo": a_iso(inicio_trabajo),
            "iniciado": inicio_trabajo is not None,
            "fin": a_iso(fin),
            "duracion_h": round((fin - inicio).total_seconds() / 3600, 2) if fin is not None and inicio is not None else None,
            "estado_paro": estado_paro,
            "tecnico": o["tecnico"] if o is not None else None,
            "estado": o["estado"] if o is not None else None,
            "ss_codigo": o["ss_codigo"] if o is not None and o["viene_ss"] else None,
            "tipo_ot": r.get("Tipo O.T.") if pd.notna(r.get("Tipo O.T.")) else None,
            # motivo registrado en Mantum (Parado por / Tipo paro / Descripción)
            "evidencia": motivo or None,
        }
        (preventivos if registro["tipo_ot"] == "Preventivo" else correctivos).append(registro)
    correctivos.sort(key=lambda x: x["inicio"] or "")
    preventivos.sort(key=lambda x: x["inicio"] or "")
    return correctivos, preventivos


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

def construir_data_json(ot_path, ss_path, disp_path, crit_path, tecnicos_path=None, proveedores_path=None, activos_path=None, paros_path=None):
    ot_raw = cargar_ot(ot_path)
    ss_raw = cargar_ss(ss_path)
    disp, periodo_texto = cargar_disponibilidad(disp_path)
    crit = cargar_criticidad(crit_path)
    resumen_activos = None
    if activos_path or {"Código Instalación de Proceso", "Nombre Instalación de Proceso"} <= set(crit.columns):
        disp, crit, resumen_activos = integrar_activos(activos_path, disp, crit)
    tecnicos_internos = cargar_tecnicos(tecnicos_path) if tecnicos_path else None
    tecnicos_inactivos = cargar_tecnicos_inactivos(tecnicos_path) if tecnicos_path else []
    proveedores = cargar_proveedores(proveedores_path) if proveedores_path else None
    if proveedores is not None or PROVEEDORES_EXTRA:
        proveedores = sorted(set(proveedores or []) | set(PROVEEDORES_EXTRA))

    disp_by_code, crit_by_code, lugar_to_ai = construir_indices(disp, crit)

    ot_df = construir_ot(ot_raw, ss_raw, disp_by_code, crit_by_code, lugar_to_ai)
    ss_df = construir_ss(ss_raw, disp_by_code, lugar_to_ai, ot_df)
    paros_raw = cargar_paros(paros_path) if paros_path else None
    paros, paros_programados = construir_paros(paros_raw, ot_df, crit, disp_by_code, lugar_to_ai)
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
        # subconjunto de tecnicos_internos con ESTADO = INACTIVO en
        # TECNICOS.xlsx (ya no trabajan en la compañía): no aparecen en
        # "Ubicación del técnico" ni en el conteo de técnicos en sitio.
        "tecnicos_inactivos": tecnicos_inactivos,
        # maestro de proveedores (Datos Generales de Proveedores.xlsx), solo
        # de referencia - no se usa para clasificar terceros (ver
        # cargar_proveedores). None si el archivo no se encontró.
        "proveedores": proveedores,
    }

    # reporte rápido en consola para poder revisar antes de subir
    resueltas = ot_df["_tipo_match"].isin(["equipo_especifico", "instalacion_pdv"]).sum()
    print(f"OT procesadas: {len(ot_df)} (de las cuales {resueltas} con equipo/lugar resuelto)")
    print(f"SS procesadas: {len(ss_df)}")
    provoca_paro_n = int((crit["Provoca Paro?"] == "Sí").sum())
    print(f"Activos que 'Provoca Paro?' = Sí: {provoca_paro_n} de {len(crit)} en el maestro")
    estados_paro = Counter(p["estado_paro"] for p in paros + paros_programados)
    print(f"Paros (informe de paros de Mantum): {len(paros)} correctivos, {len(paros_programados)} preventivos - " + ", ".join(
        f"{k}: {estados_paro.get(k, 0)}" for k in ("Fuera de servicio", "En reparación", "Finalizado")))
    for p in paros + paros_programados:
        print(f"   paro: OT {p['ot']} | {p['equipo_cod']} | {p['equipo']} | {p['lugar']} | {p['inicio']} -> {p['fin'] or 'sin fin'}")
    mantenibles = sum(1 for e in equipos if e["mantenible"])
    print(f"Equipos en directorio: {len(equipos)} de {total_maestro} en el maestro ({mantenibles} mantenibles, {len(equipos)-mantenibles} no mantenibles)")
    print(f"Criticidad: {criticidad_totales}")
    if resumen_activos:
        ra = resumen_activos
        print(f"Datos Generales Activos: {ra['activos']} activos - {len(ra['nuevos'])} nuevos frente a Disponibilidad, "
              f"{ra['trasladados']} con ubicación actualizada, {ra['sin_criticidad']} sin criticidad")
        for cod, nom, lugar, paro in ra["nuevos"]:
            print(f"   nuevo: {cod} | {nom} | {lugar} | provoca paro: {paro}")
    if tecnicos_internos is not None:
        nombres_en_ot = {
            n.strip() for r in ot_records if isinstance(r.get("tecnico"), str)
            for n in r["tecnico"].split(",")
        }
        terceros = sorted(nombres_en_ot - set(tecnicos_internos))
        print(f"Técnicos internos (TECNICOS.xlsx): {len(tecnicos_internos)} "
              f"(inactivos: {tecnicos_inactivos or 'ninguno'})")
        print(f"Nombres en OT no reconocidos como internos (van como tercero): {terceros}")
        if proveedores is not None:
            no_registrados = sorted(n for n in terceros if n not in proveedores)
            if no_registrados:
                print(f"  (de esos, no aparecen en Datos Generales de Proveedores ni en PROVEEDORES_EXTRA: {no_registrados} "
                      f"- van como 'Otros'; si ejecutan a nombre de un proveedor, agregarlos a PROVEEDORES_EXTRA)")

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
        proveedores_path = None
        activos_path = None
        paros_path = None
        for carpeta in (CARPETA_DATOS, CARPETA_BASE):
            candidatos_act = [a for a in carpeta.glob("*.xlsx") if "datos generales activos" in a.name.lower()]
            if candidatos_act and activos_path is None:
                activos_path = sorted(candidatos_act, key=lambda a: a.stat().st_mtime, reverse=True)[0]
            candidatos = [a for a in carpeta.glob("*.xlsx") if "tecnicos" in a.name.lower()]
            if candidatos and tecnicos_path is None:
                tecnicos_path = sorted(candidatos, key=lambda a: a.stat().st_mtime, reverse=True)[0]
            candidatos_prov = [a for a in carpeta.glob("*.xlsx") if "proveedores" in a.name.lower()]
            if candidatos_prov and proveedores_path is None:
                proveedores_path = sorted(candidatos_prov, key=lambda a: a.stat().st_mtime, reverse=True)[0]
            candidatos_paros = [a for a in carpeta.glob("*.xlsx") if "paros" in a.name.lower()]
            if candidatos_paros and paros_path is None:
                paros_path = sorted(candidatos_paros, key=lambda a: a.stat().st_mtime, reverse=True)[0]
    else:
        carpeta = CARPETA_DATOS if any(CARPETA_DATOS.glob("*.xlsx")) else CARPETA_BASE
        ot_path, ss_path, disp_path, crit_path, tecnicos_path, proveedores_path, activos_path, paros_path = encontrar_archivos(carpeta)

    print(f"OT:           {ot_path.name}")
    print(f"SS:           {ss_path.name}")
    print(f"Disponibilidad: {disp_path.name}")
    print(f"Datos Generales de Equipos: {crit_path.name}")
    print(f"Datos Generales Activos: {activos_path.name if activos_path else 'no encontrado - se usa Datos Generales de Equipos como maestro'}")
    if tecnicos_path:
        print(f"Técnicos:     {tecnicos_path.name}")
    else:
        print("Técnicos:     no encontrado (TECNICOS.xlsx) - la pestaña Técnicos no separará terceros")
    if proveedores_path:
        print(f"Proveedores:  {proveedores_path.name}")
    else:
        print("Proveedores:  no encontrado (Datos Generales de Proveedores.xlsx) - se omite el aviso de cruce")
    if paros_path:
        print(f"Paros:        {paros_path.name}")
    else:
        print("Paros:        no encontrado (informe de paros de Mantum) - el tablero no mostrará paros")
    print()

    data = construir_data_json(ot_path, ss_path, disp_path, crit_path, tecnicos_path, proveedores_path, activos_path, paros_path)
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
