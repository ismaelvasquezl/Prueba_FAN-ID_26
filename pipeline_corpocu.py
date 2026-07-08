# -*- coding: utf-8 -*-
"""
====================================================================
PIPELINE DE NORMALIZACION, SCORING Y GEORREFERENCIACION
Caso piloto: Coquimbo Unido  (Corporacion / Base de socios)
Archivo fuente: BD_CAPTA_CORPOCU.xlsx  ->  hoja 'BD_CORPO'
--------------------------------------------------------------------
Que hace este script, en simple:
  1. Lee el Excel tal cual viene del formulario.
  2. Renombra las columnas a nombres tecnicos limpios.
  3. Repara telefonos danados por Excel (notacion cientifica) y
     los deja en formato chileno estandar.
  4. Limpia RUN (quita puntos/guion), valida digito verificador.
  5. Normaliza emails (minusculas, sin espacios).
  6. Detecta duplicados por RUN, email, movil y nombre+ciudad.
  7. Estima un TRAMO ETARIO aproximado desde el RUN (con honestidad
     metodologica: NO es edad exacta, se entrega nivel de confianza).
  8. Normaliza ciudad/domicilio y prepara geocodificacion (geopy).
  9. Calcula un SCORING semipredictivo y asigna un segmento.
 10. Exporta CSV y JSON listos para el dashboard.

Como ejecutar:
  pip install pandas numpy openpyxl geopy
  python pipeline_corpocu.py            (sin geocodificar)
  python pipeline_corpocu.py --geo      (con geocodificacion real)

Salidas (carpeta ./salidas_corpocu):
  - bd_corpo_limpia.csv        (dataset limpio, 1 fila por registro)
  - diccionario_datos.csv      (diccionario de datos)
  - resumen_dashboard.json     (todos los agregados para el HTML)
  - geo_input.csv              (direcciones listas para geocodificar)
  - geo_resultado.csv          (se genera solo si corres --geo)
====================================================================
"""

import re
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd

# --------------------------------------------------------------------
# CONFIGURACION
# --------------------------------------------------------------------
ARCHIVO = "BD_CAPTA_CORPOCU.xlsx"
HOJA = "BD_CORPO"
SALIDA = Path("salidas_corpocu")
SALIDA.mkdir(exist_ok=True)

# Corre geocodificacion real (geopy/Nominatim) solo si pasas --geo
HACER_GEO = "--geo" in sys.argv

# Columnas originales -> nombres tecnicos (por POSICION, es lo mas robusto
# porque el ultimo encabezado del formulario es larguisimo).
NOMBRES_TECNICOS = [
    "marca_temporal",     # 0
    "email_formulario",   # 1  correo con que se envio el formulario
    "nombre",             # 2
    "id_run",             # 3
    "edad_original",      # 4  (VACIA en el 100% de los casos)
    "domicilio",          # 5
    "ciudad",             # 6
    "profesion_oficio",   # 7
    "telefono_fijo",      # 8
    "telefono_movil",     # 9
    "email_contacto",     # 10
    "acepta_estatutos",   # 11 (columna larga de consentimiento)
]


# ====================================================================
# 1. LECTURA
# ====================================================================
def leer_base():
    df = pd.read_excel(ARCHIVO, sheet_name=HOJA)
    if df.shape[1] != len(NOMBRES_TECNICOS):
        raise ValueError(
            f"Se esperaban {len(NOMBRES_TECNICOS)} columnas y llegaron {df.shape[1]}."
        )
    df.columns = NOMBRES_TECNICOS
    df["id_registro"] = range(1, len(df) + 1)
    return df


# ====================================================================
# 2. UTILIDADES DE LIMPIEZA
# ====================================================================
def limpiar_texto(x):
    """Quita espacios extra, colapsa espacios internos, deja None si vacio."""
    if pd.isna(x):
        return None
    s = str(x).strip()
    s = re.sub(r"\s+", " ", s)
    return s if s and s.lower() != "nan" else None


def normalizar_email(x):
    if pd.isna(x):
        return None
    s = str(x).strip().lower().replace(" ", "")
    return s if "@" in s and "." in s else None


def dv_run(cuerpo: str):
    """Calcula digito verificador chileno para el cuerpo numerico."""
    s = 0
    m = 2
    for d in reversed(cuerpo):
        s += int(d) * m
        m = 2 if m == 7 else m + 1
    r = 11 - (s % 11)
    if r == 11:
        return "0"
    if r == 10:
        return "K"
    return str(r)


def normalizar_run(x):
    """
    Devuelve (run_limpio, cuerpo_int, dv, run_valido).
    Quita puntos, guiones y espacios. Valida el DV.
    """
    if pd.isna(x):
        return None, None, None, False
    s = re.sub(r"[^0-9kK]", "", str(x)).upper()
    if len(s) < 2:
        return None, None, None, False
    cuerpo, dv = s[:-1], s[-1]
    if not cuerpo.isdigit():
        return None, None, None, False
    cuerpo_int = int(cuerpo)
    valido = dv_run(cuerpo) == dv
    run_limpio = f"{cuerpo}-{dv}"
    return run_limpio, cuerpo_int, dv, valido


def normalizar_telefono(x):
    """
    Repara telefonos danados por Excel y estandariza a formato chileno.
    Maneja notacion cientifica (5.6992e+10), 9 digitos, 8 digitos, etc.
    Devuelve (telefono_e164, tipo) donde tipo in {movil, fijo, invalido, vacio}.
    """
    if pd.isna(x):
        return None, "vacio"
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return None, "vacio"
    # Rescatar notacion cientifica -> entero -> string
    if re.search(r"[eE]", s) or ("." in s and "e" not in s.lower()):
        try:
            s = f"{float(s):.0f}"
        except ValueError:
            pass
    d = re.sub(r"\D", "", s)  # solo digitos
    if d == "":
        return None, "vacio"
    # Normalizaciones
    if d.startswith("56") and len(d) >= 10:
        d = d[2:]  # quita prefijo pais
    if len(d) == 9 and d.startswith("9"):
        return "+569" + d[1:], "movil"
    if len(d) == 8 and d[0] in "2345678":
        return "+56" + d, "fijo"
    if len(d) == 9 and d[0] in "2345678":
        return "+56" + d, "fijo"
    return d, "invalido"


# ====================================================================
# 3. ESTIMACION DE TRAMO ETARIO DESDE EL RUN  (con honestidad)
# ====================================================================
# ADVERTENCIA METODOLOGICA:
# El RUN chileno se asigna de forma correlativa segun orden de inscripcion
# en el Registro Civil. Existe una correlacion ESTADISTICA entre el numero
# de RUN y el ano de nacimiento, pero NO es exacta: hay inscripciones
# tardias, extranjeros nacionalizados, y variacion natural. Por eso NO
# entregamos edad exacta: entregamos un TRAMO y un NIVEL DE CONFIANZA.
ANCLAS_RUN = [
    (3_000_000, 1943),
    (5_000_000, 1953),
    (7_000_000, 1961),
    (9_000_000, 1969),
    (10_000_000, 1972),
    (11_000_000, 1975),
    (12_000_000, 1977),
    (13_000_000, 1980),
    (14_000_000, 1983),
    (15_000_000, 1985),
    (16_000_000, 1987),
    (17_000_000, 1990),
    (18_000_000, 1992),
    (19_000_000, 1994),
    (20_000_000, 1996),
    (21_000_000, 1998),
    (22_000_000, 2000),
    (23_000_000, 2001),
    (24_000_000, 2003),
    (25_000_000, 2005),
    (26_000_000, 2007),
    (27_000_000, 2009),
]
ANIO_ACTUAL = 2026


def estimar_desde_run(cuerpo_int, run_valido):
    """Devuelve (anio_nac_aprox, edad_aprox, tramo, confianza)."""
    if cuerpo_int is None or pd.isna(cuerpo_int):
        return None, None, "sin_dato", "nula"
    cuerpo_int = int(cuerpo_int)
    if cuerpo_int < 2_000_000 or cuerpo_int > 27_500_000:
        return None, None, "fuera_de_rango", "nula"

    xs = [a[0] for a in ANCLAS_RUN]
    ys = [a[1] for a in ANCLAS_RUN]
    anio = int(round(float(np.interp(cuerpo_int, xs, ys))))
    edad = ANIO_ACTUAL - anio
    if edad < 15:
        edad = 15

    if edad <= 17:
        tramo = "menor_18"
    elif edad <= 25:
        tramo = "18_25"
    elif edad <= 34:
        tramo = "26_34"
    elif edad <= 44:
        tramo = "35_44"
    elif edad <= 54:
        tramo = "45_54"
    elif edad <= 64:
        tramo = "55_64"
    else:
        tramo = "65_mas"

    if not run_valido:
        conf = "baja"
    elif 12_000_000 <= cuerpo_int <= 26_000_000:
        conf = "media"
    else:
        conf = "baja"
    return anio, edad, tramo, conf


# ====================================================================
# 4. NORMALIZACION DE CIUDAD (typos frecuentes detectados)
# ====================================================================
MAPA_CIUDAD = {
    "COQUIMBO": "Coquimbo",
    "COQUIMBI": "Coquimbo",
    "COQUIMNO": "Coquimbo",
    "COQUIMB": "Coquimbo",
    "LA SERENA": "La Serena",
    "SERENA": "La Serena",
    "GUANAQUEROS": "Guanaqueros",
    "TONGOY": "Tongoy",
    "SANTIAGO": "Santiago",
    "CALAMA": "Calama",
    "COPIAPO": "Copiapo",
    "COPIAPÓ": "Copiapo",
    "RANCAGUA": "Rancagua",
    "SAN BERNARDO": "San Bernardo",
    "PUERTO MONTT": "Puerto Montt",
    "CANBERRA": "Canberra",
}


def normalizar_ciudad(x):
    s = limpiar_texto(x)
    if s is None:
        return None
    key = s.upper().strip()
    if key in MAPA_CIUDAD:
        return MAPA_CIUDAD[key]
    return s.title()


# ====================================================================
# 5. PROCESO PRINCIPAL
# ====================================================================
def procesar(df):
    df["nombre"] = df["nombre"].apply(limpiar_texto)
    df["domicilio"] = df["domicilio"].apply(limpiar_texto)
    df["profesion_oficio"] = df["profesion_oficio"].apply(limpiar_texto)
    df["ciudad_norm"] = df["ciudad"].apply(normalizar_ciudad)

    df["email_norm"] = df["email_contacto"].apply(normalizar_email)
    df["email_form_norm"] = df["email_formulario"].apply(normalizar_email)
    df["email_final"] = df["email_norm"].fillna(df["email_form_norm"])

    run_res = df["id_run"].apply(normalizar_run)
    df["run_limpio"] = run_res.apply(lambda t: t[0])
    df["run_cuerpo"] = run_res.apply(lambda t: t[1])
    df["run_dv"] = run_res.apply(lambda t: t[2])
    df["run_valido"] = run_res.apply(lambda t: t[3])

    mov = df["telefono_movil"].apply(normalizar_telefono)
    df["movil_norm"] = mov.apply(lambda t: t[0])
    df["movil_tipo"] = mov.apply(lambda t: t[1])
    fij = df["telefono_fijo"].apply(normalizar_telefono)
    df["fijo_norm"] = fij.apply(lambda t: t[0])
    df["fijo_tipo"] = fij.apply(lambda t: t[1])
    df["movil_valido"] = df["movil_tipo"] == "movil"

    est = df.apply(lambda r: estimar_desde_run(r["run_cuerpo"], r["run_valido"]), axis=1)
    df["anio_nac_aprox"] = est.apply(lambda t: t[0])
    df["edad_estimada_aprox"] = est.apply(lambda t: t[1])
    df["tramo_etario_estimado"] = est.apply(lambda t: t[2])
    df["nivel_confianza_edad"] = est.apply(lambda t: t[3])

    df["acepta_estatutos_bool"] = (
        df["acepta_estatutos"].astype(str).str.strip().str.lower()
        .isin(["si", "sí", "ok", "okey", "okey.", "si."])
    )

    df["dup_run"] = df["run_limpio"].notna() & df["run_limpio"].duplicated(keep="first")
    df["dup_email"] = df["email_final"].notna() & df["email_final"].duplicated(keep="first")
    df["dup_movil"] = df["movil_valido"] & df["movil_norm"].duplicated(keep="first")
    clave_nc = (df["nombre"].astype(str).str.lower() + "|" + df["ciudad_norm"].astype(str).str.lower())
    df["dup_nombre_ciudad"] = clave_nc.duplicated(keep="first") & df["nombre"].notna()
    df["es_duplicado"] = df[["dup_run", "dup_email", "dup_movil"]].any(axis=1)

    campos_clave = ["nombre", "run_limpio", "movil_norm", "email_final", "domicilio", "ciudad_norm"]
    df["campos_completos"] = df[campos_clave].notna().sum(axis=1)
    df["completitud_pct"] = (df["campos_completos"] / len(campos_clave) * 100).round(0)

    def nivel_calidad(r):
        if r["es_duplicado"]:
            return "requiere_validacion"
        if r["run_valido"] and r["movil_valido"] and r["email_final"] and r["domicilio"]:
            return "alta"
        if r["movil_valido"] and (r["email_final"] or r["run_valido"]):
            return "media"
        return "baja"

    df["nivel_calidad"] = df.apply(nivel_calidad, axis=1)
    return df


# ====================================================================
# 6. SCORING SEMIPREDICTIVO  (0-100)
# ====================================================================
PESOS = {
    "contactabilidad": 35,
    "identidad": 20,
    "completitud": 15,
    "territorio": 20,
    "unicidad": 10,
}


def calcular_scoring(df):
    ciudad_local = {"Coquimbo": 1.0, "La Serena": 0.8, "Guanaqueros": 0.9,
                    "Tongoy": 0.85, "Ovalle": 0.6}

    def score_row(r):
        s = 0.0
        c = 0.0
        if r["movil_valido"]:
            c += 0.6
        if r["email_final"]:
            c += 0.4
        s += PESOS["contactabilidad"] * c
        s += PESOS["identidad"] * (1.0 if r["run_valido"] else 0.0)
        s += PESOS["completitud"] * (r["campos_completos"] / 6.0)
        t = ciudad_local.get(r["ciudad_norm"], 0.3)
        s += PESOS["territorio"] * t
        s += PESOS["unicidad"] * (0.0 if r["es_duplicado"] else 1.0)
        return round(s, 1)

    df["score"] = df.apply(score_row, axis=1)

    def segmento(r):
        core = r["ciudad_norm"] in ("Coquimbo", "Guanaqueros", "Tongoy")
        contactable = r["movil_valido"] and bool(r["email_final"])
        if r["es_duplicado"]:
            return "requiere_validacion"          # duplica RUN/email/movil -> depurar
        if not r["movil_valido"] or r["completitud_pct"] < 80:
            return "registro_incompleto_recuperable"  # falta contacto/dato recuperable
        if not core:
            return "contacto_utilizable"           # contactable pero fuera del territorio nucleo
        if contactable and r["run_valido"] and r["score"] >= 95:
            return "alta_prioridad_contacto"       # nucleo, contactable, identidad ok
        return "alto_potencial_territorial"        # nucleo, activable, requiere leve enriquecimiento

    df["segmento"] = df.apply(segmento, axis=1)
    return df


# ====================================================================
# 7. GEOCODIFICACION (opcional, con geopy/Nominatim)
# ====================================================================
def preparar_geo(df):
    def direccion_completa(r):
        partes = [r["domicilio"], r["ciudad_norm"], "Region de Coquimbo", "Chile"]
        partes = [p for p in partes if p]
        return ", ".join(partes)

    geo = df[["id_registro", "domicilio", "ciudad_norm"]].copy()
    geo["direccion_geo"] = df.apply(direccion_completa, axis=1)

    def precision(r):
        if r["domicilio"] and re.search(r"\d", str(r["domicilio"])):
            return "direccion_exacta"
        if r["domicilio"]:
            return "direccion_parcial"
        if r["ciudad_norm"]:
            return "solo_ciudad"
        return "sin_dato"

    geo["nivel_precision"] = df.apply(precision, axis=1)
    geo.to_csv(SALIDA / "geo_input.csv", index=False, encoding="utf-8-sig")
    return geo


def geocodificar(geo):
    """Solo si corres con --geo. Respeta el limite de 1 req/seg de Nominatim."""
    try:
        from geopy.geocoders import Nominatim
        from geopy.extra.rate_limiter import RateLimiter
    except ImportError:
        print("geopy no instalado; omito geocodificacion. Instala con: pip install geopy")
        return
    geoloc = Nominatim(user_agent="corpocu_piloto")
    geocode = RateLimiter(geoloc.geocode, min_delay_seconds=1.1)
    lats, lons = [], []
    for _, r in geo.iterrows():
        try:
            loc = geocode(r["direccion_geo"])
        except Exception:
            loc = None
        lats.append(loc.latitude if loc else None)
        lons.append(loc.longitude if loc else None)
    geo["lat"] = lats
    geo["lon"] = lons
    geo.to_csv(SALIDA / "geo_resultado.csv", index=False, encoding="utf-8-sig")
    print(f"Geocodificados {geo['lat'].notna().sum()}/{len(geo)} registros.")


# ====================================================================
# 8. AGREGADOS PARA EL DASHBOARD
# ====================================================================
def construir_resumen(df):
    def vc(col):
        return df[col].value_counts(dropna=False).to_dict()

    resumen = {
        "total_registros": int(len(df)),
        "duplicados": {
            "por_run": int(df["dup_run"].sum()),
            "por_email": int(df["dup_email"].sum()),
            "por_movil": int(df["dup_movil"].sum()),
            "por_nombre_ciudad": int(df["dup_nombre_ciudad"].sum()),
            "registros_marcados_duplicados": int(df["es_duplicado"].sum()),
            "registros_unicos": int((~df["es_duplicado"]).sum()),
        },
        "calidad_contacto": {
            "movil_valido": int(df["movil_valido"].sum()),
            "movil_invalido": int((df["movil_tipo"] == "invalido").sum()),
            "email_valido": int(df["email_final"].notna().sum()),
            "run_valido": int(df["run_valido"].sum()),
            "run_invalido": int((~df["run_valido"]).sum()),
            "tiene_fijo": int(df["fijo_norm"].notna().sum()),
        },
        "nivel_calidad": {str(k): int(v) for k, v in vc("nivel_calidad").items()},
        "completitud_promedio_pct": float(round(df["completitud_pct"].mean(), 1)),
        "territorio": {str(k): int(v) for k, v in df["ciudad_norm"].value_counts().head(15).to_dict().items()},
        "profesiones_top": {str(k): int(v) for k, v in df["profesion_oficio"].value_counts().head(15).to_dict().items()},
        "tramo_etario_estimado": {str(k): int(v) for k, v in vc("tramo_etario_estimado").items()},
        "confianza_edad": {str(k): int(v) for k, v in vc("nivel_confianza_edad").items()},
        "segmentos": {str(k): int(v) for k, v in vc("segmento").items()},
        "score": {
            "promedio": float(round(df["score"].mean(), 1)),
            "mediana": float(df["score"].median()),
            "min": float(df["score"].min()),
            "max": float(df["score"].max()),
        },
        "edad_vacia": bool(df["edad_original"].isna().all()),
        "consentimiento_explicito": int(df["acepta_estatutos_bool"].sum()),
    }
    return resumen


def construir_diccionario():
    filas = [
        ("marca_temporal", "datetime", "Fecha/hora de envio del formulario", "Original"),
        ("email_formulario", "texto", "Correo con que se completo el formulario (Google)", "Original"),
        ("nombre", "texto", "Nombre del socio/postulante", "Limpiado"),
        ("id_run", "texto", "RUN original tal como fue ingresado", "Original"),
        ("edad_original", "vacio", "Columna de edad: VACIA en 100% -> no utilizable", "Original"),
        ("domicilio", "texto", "Direccion declarada", "Limpiado"),
        ("ciudad", "texto", "Ciudad declarada (con typos)", "Original"),
        ("profesion_oficio", "texto", "Profesion u oficio declarado", "Limpiado"),
        ("telefono_fijo", "texto", "Telefono fijo (opcional, casi siempre vacio)", "Original"),
        ("telefono_movil", "texto", "Telefono movil (obligatorio; danado por Excel)", "Original"),
        ("email_contacto", "texto", "Correo de contacto (obligatorio)", "Original"),
        ("acepta_estatutos", "texto", "Consentimiento/compromiso estatutario", "Original"),
        ("run_limpio", "texto", "RUN normalizado cuerpo-DV, sin puntos", "Derivado"),
        ("run_valido", "bool", "TRUE si el digito verificador es correcto", "Derivado"),
        ("movil_norm", "texto", "Movil reparado en formato +569XXXXXXXX", "Derivado"),
        ("movil_valido", "bool", "TRUE si es un movil chileno valido", "Derivado"),
        ("email_final", "texto", "Email normalizado (minusculas, sin espacios)", "Derivado"),
        ("ciudad_norm", "texto", "Ciudad normalizada (typos corregidos)", "Derivado"),
        ("edad_estimada_aprox", "entero", "Edad APROXIMADA estimada desde RUN (NO exacta)", "Derivado"),
        ("tramo_etario_estimado", "categoria", "Tramo etario estimado desde RUN", "Derivado"),
        ("nivel_confianza_edad", "categoria", "Confianza de la estimacion de edad", "Derivado"),
        ("completitud_pct", "numero", "% de campos clave completos", "Derivado"),
        ("nivel_calidad", "categoria", "alta/media/baja/requiere_validacion", "Derivado"),
        ("es_duplicado", "bool", "TRUE si duplica RUN, email o movil previo", "Derivado"),
        ("score", "numero", "Score de priorizacion 0-100 (semipredictivo)", "Derivado"),
        ("segmento", "categoria", "Segmento accionable de captacion/relacion", "Derivado"),
    ]
    return pd.DataFrame(filas, columns=["campo", "tipo", "descripcion", "origen"])


# ====================================================================
# MAIN
# ====================================================================
def main():
    print(">> Leyendo base...")
    df = leer_base()
    print(f"   {len(df)} registros, {df.shape[1]} columnas.")

    print(">> Normalizando y limpiando...")
    df = procesar(df)

    print(">> Calculando scoring semipredictivo...")
    df = calcular_scoring(df)

    print(">> Preparando geocodificacion...")
    geo = preparar_geo(df)
    if HACER_GEO:
        print(">> Geocodificando (puede tardar ~1 seg por registro)...")
        geocodificar(geo)

    cols_export = [
        "id_registro", "nombre", "run_limpio", "run_valido",
        "movil_norm", "movil_valido", "fijo_norm", "email_final",
        "domicilio", "ciudad_norm", "profesion_oficio",
        "edad_estimada_aprox", "tramo_etario_estimado", "nivel_confianza_edad",
        "completitud_pct", "nivel_calidad", "es_duplicado",
        "score", "segmento", "acepta_estatutos_bool", "marca_temporal",
    ]
    df[cols_export].to_csv(SALIDA / "bd_corpo_limpia.csv", index=False, encoding="utf-8-sig")
    construir_diccionario().to_csv(SALIDA / "diccionario_datos.csv", index=False, encoding="utf-8-sig")

    resumen = construir_resumen(df)
    with open(SALIDA / "resumen_dashboard.json", "w", encoding="utf-8") as f:
        json.dump(resumen, f, ensure_ascii=False, indent=2)

    print("\n>> LISTO. Archivos en:", SALIDA.resolve())
    for p in sorted(SALIDA.glob("*")):
        print("   -", p.name)
    print("\n>> Resumen rapido:")
    print(json.dumps(resumen, ensure_ascii=False, indent=2)[:1400])


if __name__ == "__main__":
    main()
