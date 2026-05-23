from pathlib import Path
import re
import unicodedata

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Cotizador Sophos SMB | Media Commerce",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).resolve().parent
PRICE_FILE = BASE_DIR / "data" / "Lista de precios SOPHOS SMB.xlsx"
LOGO_FILE = BASE_DIR / "assets" / "logo_mc.png"

FAMILIAS_VALIDAS = [
    "Firewall",
    "Endpoint",
    "Server",
    "EDR User",
    "EDR Server",
    "XDR User",
    "XDR Server",
]

VIGENCIAS_VALIDAS = [12, 24, 36]

TIPOS_VENTA_FIREWALL = [
    "Venta one shot",
    "Servicio mensual"
]


# ============================================================
# VARIABLES FINANCIERAS FIJAS
# ============================================================

IMPUESTOS = 0.012
GASTOS = 0.045
CARGA_PRESTACIONAL_COMISION = 0.014
COMISION = 0.035

PORCENTAJE_COSTO_SOBRE_VENTA = 0.739
PORCENTAJE_COSTO_INSTALACION = 0.10


# ============================================================
# TARIFAS SERVICIO FIREWALL
# ============================================================

TARIFA_SERVICIO_FIREWALL = {
    "XGS 88": {12: 367000, 24: 314000, 36: 294000},
    "XGS 108": {12: 447000, 24: 382000, 36: 357000},
    "XGS 118": {12: 563000, 24: 493000, 36: 449000},
    "XGS 128": {12: 758000, 24: 651000, 36: 595000},
    "XGS 138": {12: 927000, 24: 846000, 36: 785000},
}

NRC_SERVICIO_FIREWALL = {
    "Instalación completa": 350000,
    "Instalación básica": 200000,
    "Sin instalación": 0,
}


# ============================================================
# ESTILO CORPORATIVO
# ============================================================

st.markdown(
    """
    <style>
        :root {
            --mc-blue: #0057A8;
            --mc-blue-dark: #003B73;
            --mc-blue-soft: #E8F2FC;
            --mc-yellow: #FFC400;
            --mc-gray-bg: #F5F7FA;
            --mc-gray-line: #D9E1EA;
            --mc-text: #1F2937;
            --mc-muted: #6B7280;
            --mc-white: #FFFFFF;
            --mc-green: #2E7D32;
        }

        .stApp {
            background: linear-gradient(180deg, #F7FAFD 0%, #FFFFFF 38%);
            color: var(--mc-text);
        }

        [data-testid="stToolbar"] {
            visibility: hidden;
            height: 0%;
            position: fixed;
        }

        [data-testid="stDecoration"] {
            background: linear-gradient(90deg, #0057A8 0%, #006FCB 55%, #FFC400 100%);
            height: 5px;
        }

        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 3rem;
            max-width: 1320px;
        }

        .mc-header {
            background: linear-gradient(135deg, #FFFFFF 0%, #EEF6FF 100%);
            border: 1px solid #D9E7F5;
            border-radius: 22px;
            padding: 28px 32px;
            margin-bottom: 24px;
            box-shadow: 0 12px 30px rgba(0, 57, 115, 0.08);
        }

        .mc-header-grid {
            display: grid;
            grid-template-columns: 320px 1fr;
            gap: 32px;
            align-items: center;
        }

        .mc-title {
            color: #003B73;
            font-size: 34px;
            line-height: 1.15;
            font-weight: 800;
            margin-bottom: 8px;
        }

        .mc-subtitle {
            color: #52616F;
            font-size: 16px;
            line-height: 1.45;
            max-width: 760px;
        }

        .mc-badge-row {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-top: 16px;
        }

        .mc-badge {
            background: #FFFFFF;
            color: #0057A8;
            border: 1px solid #BFD7EF;
            padding: 7px 12px;
            border-radius: 999px;
            font-size: 13px;
            font-weight: 700;
        }

        .mc-section-title {
            color: #003B73;
            font-size: 25px;
            font-weight: 800;
            margin: 18px 0 8px 0;
        }

        .mc-section-caption {
            color: #6B7280;
            font-size: 14px;
            margin-bottom: 18px;
        }

        .mc-panel {
            background: #FFFFFF;
            border: 1px solid #D9E1EA;
            border-radius: 20px;
            padding: 24px;
            margin-bottom: 24px;
            box-shadow: 0 10px 28px rgba(31, 41, 55, 0.06);
        }

        .mc-result-card {
            background: #FFFFFF;
            border: 1px solid #D9E1EA;
            border-left: 6px solid #0057A8;
            border-radius: 18px;
            padding: 18px 20px;
            min-height: 128px;
            box-shadow: 0 8px 24px rgba(0, 57, 115, 0.06);
        }

        .mc-result-card-yellow {
            border-left-color: #FFC400;
        }

        .mc-result-card-green {
            border-left-color: #2E7D32;
        }

        .mc-result-label {
            color: #6B7280;
            font-size: 14px;
            font-weight: 700;
            margin-bottom: 8px;
        }

        .mc-result-value {
            color: #1F2937;
            font-size: 30px;
            font-weight: 800;
            line-height: 1.2;
        }

        .mc-result-note {
            color: #6B7280;
            font-size: 12px;
            margin-top: 8px;
        }

        .mc-alert {
            background: #E8F2FC;
            border: 1px solid #BFD7EF;
            border-left: 6px solid #0057A8;
            color: #003B73;
            border-radius: 16px;
            padding: 16px 18px;
            font-weight: 650;
            margin: 14px 0 22px 0;
        }

        .mc-footer-note {
            color: #6B7280;
            font-size: 12px;
            text-align: center;
            margin-top: 28px;
            padding-top: 20px;
            border-top: 1px solid #D9E1EA;
        }

        div[data-testid="stMetric"] {
            background: #FFFFFF;
            border: 1px solid #D9E1EA;
            border-radius: 18px;
            padding: 18px 18px 14px 18px;
            box-shadow: 0 8px 24px rgba(0, 57, 115, 0.06);
        }

        div[data-testid="stMetric"] label {
            color: #52616F !important;
            font-weight: 700 !important;
        }

        div[data-testid="stMetricValue"] {
            color: #1F2937 !important;
            font-weight: 800 !important;
        }

        div[data-testid="stDataFrame"] {
            border-radius: 16px;
            overflow: hidden;
        }

        .stSelectbox label, .stNumberInput label {
            color: #374151 !important;
            font-weight: 700 !important;
        }

        button[kind="primary"] {
            background: #0057A8 !important;
        }

        @media (max-width: 900px) {
            .mc-header-grid {
                grid-template-columns: 1fr;
            }

            .mc-title {
                font-size: 28px;
            }
        }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# FUNCIONES UTILITARIAS
# ============================================================

def normalizar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^a-z0-9]", "", texto)
    return texto


def convertir_numero(valor):
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()
    texto = texto.replace("USD", "")
    texto = texto.replace("$", "")
    texto = texto.replace(" ", "")

    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")

    try:
        return float(texto)
    except ValueError:
        return 0.0


def formato_cop(valor):
    return f"${valor:,.0f}".replace(",", ".")


def formato_usd(valor):
    return f"USD {valor:,.2f}"


def formato_porcentaje(valor):
    return f"{valor * 100:.1f}%".replace(".", ",")


def obtener_columna(df, posibles_nombres):
    columnas_normalizadas = {
        normalizar_texto(col): col
        for col in df.columns
    }

    for nombre in posibles_nombres:
        nombre_norm = normalizar_texto(nombre)
        if nombre_norm in columnas_normalizadas:
            return columnas_normalizadas[nombre_norm]

    return None


def sugerir_firewall(cantidad_usuarios):
    if cantidad_usuarios <= 20:
        return "XGS 88", 20
    elif cantidad_usuarios <= 40:
        return "XGS 108", 40
    elif cantidad_usuarios <= 70:
        return "XGS 118", 70
    elif cantidad_usuarios <= 100:
        return "XGS 128", 100
    elif cantidad_usuarios <= 200:
        return "XGS 138", 200
    return None, None


def calcular_precio_venta_desde_costo(costo_cop):
    if PORCENTAJE_COSTO_SOBRE_VENTA > 0:
        return costo_cop / PORCENTAJE_COSTO_SOBRE_VENTA
    return 0


def obtener_costo_usd_por_escenario(msrp_usd, dr_usd, escenario):
    msrp_usd = convertir_numero(msrp_usd)
    dr_usd = convertir_numero(dr_usd)

    if escenario == "Con DR aprobado" and dr_usd > 0:
        return dr_usd

    return msrp_usd


def render_result_card(label, value, note="", variant="blue"):
    css_class = "mc-result-card"
    if variant == "yellow":
        css_class += " mc-result-card-yellow"
    elif variant == "green":
        css_class += " mc-result-card-green"

    st.markdown(
        f"""
        <div class="{css_class}">
            <div class="mc-result-label">{label}</div>
            <div class="mc-result-value">{value}</div>
            <div class="mc-result-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# CARGA DE LISTA DE PRECIOS
# ============================================================

@st.cache_data
def cargar_lista_precios():
    if not PRICE_FILE.exists():
        return None, None

    xls = pd.ExcelFile(PRICE_FILE)

    if "Lista_Precios" in xls.sheet_names:
        hoja = "Lista_Precios"
    else:
        hoja = xls.sheet_names[0]

    df = pd.read_excel(PRICE_FILE, sheet_name=hoja)
    df = df.dropna(how="all")

    return df, hoja


def mapear_columnas(df):
    columnas = {
        "familia": obtener_columna(df, ["Familia"]),
        "producto": obtener_columna(df, ["Producto", "Modelo", "Modelo / Producto", "Nombre Producto"]),
        "min": obtener_columna(df, ["Min Usuarios", "Min", "Rango Min", "Desde", "Cantidad Min"]),
        "max": obtener_columna(df, ["Max Usuarios", "Max", "Rango Max", "Hasta", "Cantidad Max"]),
        "vigencia": obtener_columna(df, ["Vigencia Meses", "Vigencia", "Meses"]),
        "msrp": obtener_columna(df, ["MSRP USD", "MSRP", "Precio MSRP USD"]),
        "dr": obtener_columna(df, ["Con DR USD", "DR USD", "Precio DR USD", "Con Descuento USD"]),
        "sku": obtener_columna(df, ["SKU", "Código", "Codigo"]),
    }

    obligatorias = ["familia", "producto", "vigencia", "msrp", "dr"]
    faltantes = [nombre for nombre in obligatorias if columnas[nombre] is None]

    return columnas, faltantes


# ============================================================
# BÚSQUEDA DE PRECIOS
# ============================================================

def buscar_precio_firewall(df, columnas, modelo_sugerido, vigencia):
    familia_col = columnas["familia"]
    producto_col = columnas["producto"]
    vigencia_col = columnas["vigencia"]
    msrp_col = columnas["msrp"]
    dr_col = columnas["dr"]
    sku_col = columnas["sku"]

    df_temp = df.copy()
    df_temp["_familia_norm"] = df_temp[familia_col].apply(normalizar_texto)
    df_temp["_producto_norm"] = df_temp[producto_col].apply(normalizar_texto)
    df_temp["_vigencia_num"] = df_temp[vigencia_col].apply(convertir_numero)

    modelo_norm = normalizar_texto(modelo_sugerido)

    filtro_equipo = (
        (df_temp["_familia_norm"] == normalizar_texto("Firewall")) &
        (df_temp["_producto_norm"].str.contains(modelo_norm, na=False)) &
        (df_temp["_vigencia_num"] == 0)
    )

    filtro_licencia = (
        (df_temp["_familia_norm"] == normalizar_texto("Firewall")) &
        (df_temp["_producto_norm"].str.contains(modelo_norm, na=False)) &
        (df_temp["_vigencia_num"] == vigencia)
    )

    fila_equipo = df_temp[filtro_equipo]
    fila_licencia = df_temp[filtro_licencia]

    if fila_equipo.empty and fila_licencia.empty:
        return None

    resultado = pd.concat([fila_equipo, fila_licencia], ignore_index=True)

    msrp_equipo_usd = fila_equipo[msrp_col].apply(convertir_numero).sum() if not fila_equipo.empty else 0
    dr_equipo_usd = fila_equipo[dr_col].apply(convertir_numero).sum() if not fila_equipo.empty else 0

    msrp_licencia_usd = fila_licencia[msrp_col].apply(convertir_numero).sum() if not fila_licencia.empty else 0
    dr_licencia_usd = fila_licencia[dr_col].apply(convertir_numero).sum() if not fila_licencia.empty else 0

    if dr_equipo_usd <= 0:
        dr_equipo_usd = msrp_equipo_usd

    if dr_licencia_usd <= 0:
        dr_licencia_usd = msrp_licencia_usd

    msrp_total = msrp_equipo_usd + msrp_licencia_usd
    dr_total = dr_equipo_usd + dr_licencia_usd

    skus = []
    if sku_col:
        skus = resultado[sku_col].dropna().astype(str).tolist()

    detalle_firewall = pd.DataFrame([
        {
            "Concepto": "Equipo / Hardware",
            "Modelo": modelo_sugerido,
            "Vigencia": "Pago único",
            "MSRP USD": formato_usd(msrp_equipo_usd),
            "DR USD": formato_usd(dr_equipo_usd),
        },
        {
            "Concepto": f"Licencia {vigencia} meses",
            "Modelo": modelo_sugerido,
            "Vigencia": f"{vigencia} meses",
            "MSRP USD": formato_usd(msrp_licencia_usd),
            "DR USD": formato_usd(dr_licencia_usd),
        },
        {
            "Concepto": "Total equipo + licencia",
            "Modelo": modelo_sugerido,
            "Vigencia": f"{vigencia} meses",
            "MSRP USD": formato_usd(msrp_total),
            "DR USD": formato_usd(dr_total),
        },
    ])

    return {
        "producto": modelo_sugerido,
        "msrp_usd": msrp_total,
        "dr_usd": dr_total,
        "msrp_unitario_usd": msrp_total,
        "dr_unitario_usd": dr_total,
        "msrp_equipo_usd": msrp_equipo_usd,
        "dr_equipo_usd": dr_equipo_usd,
        "msrp_licencia_usd": msrp_licencia_usd,
        "dr_licencia_usd": dr_licencia_usd,
        "skus": skus,
        "filas": resultado,
        "detalle_firewall": detalle_firewall,
    }


def buscar_precio_licencia(df, columnas, familia, cantidad, vigencia):
    familia_col = columnas["familia"]
    producto_col = columnas["producto"]
    vigencia_col = columnas["vigencia"]
    msrp_col = columnas["msrp"]
    dr_col = columnas["dr"]
    min_col = columnas["min"]
    max_col = columnas["max"]
    sku_col = columnas["sku"]

    df_temp = df.copy()
    df_temp["_familia_norm"] = df_temp[familia_col].apply(normalizar_texto)
    df_temp["_vigencia_num"] = df_temp[vigencia_col].apply(convertir_numero)

    resultado = df_temp[
        (df_temp["_familia_norm"] == normalizar_texto(familia)) &
        (df_temp["_vigencia_num"] == vigencia)
    ]

    if min_col and max_col:
        resultado = resultado[
            (resultado[min_col].apply(convertir_numero) <= cantidad) &
            (resultado[max_col].apply(convertir_numero) >= cantidad)
        ]

    if resultado.empty:
        return None

    fila = resultado.iloc[0]

    msrp_unitario = convertir_numero(fila[msrp_col])
    dr_unitario = convertir_numero(fila[dr_col])

    if dr_unitario <= 0:
        dr_unitario = msrp_unitario

    skus = []
    if sku_col and not pd.isna(fila[sku_col]):
        skus = [str(fila[sku_col])]

    return {
        "producto": fila[producto_col],
        "msrp_usd": msrp_unitario * cantidad,
        "dr_usd": dr_unitario * cantidad,
        "msrp_unitario_usd": msrp_unitario,
        "dr_unitario_usd": dr_unitario,
        "msrp_equipo_usd": 0,
        "dr_equipo_usd": 0,
        "msrp_licencia_usd": msrp_unitario * cantidad,
        "dr_licencia_usd": dr_unitario * cantidad,
        "skus": skus,
        "filas": resultado,
        "detalle_firewall": None,
    }


# ============================================================
# CÁLCULOS FINANCIEROS
# ============================================================

def calcular_financiero(msrp_usd, dr_usd, escenario, trm, tipo_instalacion):
    msrp_usd = convertir_numero(msrp_usd)
    dr_usd = convertir_numero(dr_usd)

    costo_base_usd = obtener_costo_usd_por_escenario(msrp_usd, dr_usd, escenario)

    if costo_base_usd <= 0:
        costo_base_usd = msrp_usd

    costo_licencias_cop = costo_base_usd * trm
    precio_venta_licencias_cop = calcular_precio_venta_desde_costo(costo_licencias_cop)

    if tipo_instalacion == "Instalación completa":
        costo_instalacion_cop = precio_venta_licencias_cop * PORCENTAJE_COSTO_INSTALACION
        precio_venta_instalacion_cop = calcular_precio_venta_desde_costo(costo_instalacion_cop)
    else:
        costo_instalacion_cop = 0
        precio_venta_instalacion_cop = 0

    precio_venta_total_cop = precio_venta_licencias_cop + precio_venta_instalacion_cop
    costo_total_cop = costo_licencias_cop + costo_instalacion_cop

    utilidad_bruta_cop = precio_venta_total_cop - costo_total_cop
    impuestos_cop = precio_venta_total_cop * IMPUESTOS
    gastos_cop = precio_venta_total_cop * GASTOS
    margen_contribucional_cop = utilidad_bruta_cop - impuestos_cop - gastos_cop
    carga_prestacional_cop = precio_venta_total_cop * CARGA_PRESTACIONAL_COMISION
    comision_cop = precio_venta_total_cop * COMISION
    utilidad_ebit_cop = margen_contribucional_cop - carga_prestacional_cop - comision_cop

    porcentaje_costo = costo_total_cop / precio_venta_total_cop if precio_venta_total_cop else 0
    porcentaje_utilidad_bruta = utilidad_bruta_cop / precio_venta_total_cop if precio_venta_total_cop else 0
    porcentaje_margen_contribucional = margen_contribucional_cop / precio_venta_total_cop if precio_venta_total_cop else 0
    porcentaje_ebit = utilidad_ebit_cop / precio_venta_total_cop if precio_venta_total_cop else 0
    ebit_sobre_costo = utilidad_ebit_cop / costo_total_cop if costo_total_cop else 0

    return {
        "costo_licencias_cop": costo_licencias_cop,
        "precio_venta_licencias_cop": precio_venta_licencias_cop,
        "costo_instalacion_cop": costo_instalacion_cop,
        "precio_venta_instalacion_cop": precio_venta_instalacion_cop,
        "precio_venta_total_cop": precio_venta_total_cop,
        "costo_total_cop": costo_total_cop,
        "utilidad_bruta_cop": utilidad_bruta_cop,
        "impuestos_cop": impuestos_cop,
        "gastos_cop": gastos_cop,
        "margen_contribucional_cop": margen_contribucional_cop,
        "carga_prestacional_cop": carga_prestacional_cop,
        "comision_cop": comision_cop,
        "utilidad_ebit_cop": utilidad_ebit_cop,
        "porcentaje_costo": porcentaje_costo,
        "porcentaje_utilidad_bruta": porcentaje_utilidad_bruta,
        "porcentaje_margen_contribucional": porcentaje_margen_contribucional,
        "porcentaje_ebit": porcentaje_ebit,
        "ebit_sobre_costo": ebit_sobre_costo,
    }


def calcular_detalle_venta_firewall(precio, escenario, trm):
    costo_equipo_usd = obtener_costo_usd_por_escenario(
        precio["msrp_equipo_usd"],
        precio["dr_equipo_usd"],
        escenario
    )

    costo_licencia_usd = obtener_costo_usd_por_escenario(
        precio["msrp_licencia_usd"],
        precio["dr_licencia_usd"],
        escenario
    )

    costo_equipo_cop = costo_equipo_usd * trm
    costo_licencia_cop = costo_licencia_usd * trm

    venta_equipo_cop = calcular_precio_venta_desde_costo(costo_equipo_cop)
    venta_licencia_cop = calcular_precio_venta_desde_costo(costo_licencia_cop)

    return {
        "costo_equipo_usd": costo_equipo_usd,
        "costo_licencia_usd": costo_licencia_usd,
        "costo_equipo_cop": costo_equipo_cop,
        "costo_licencia_cop": costo_licencia_cop,
        "venta_equipo_cop": venta_equipo_cop,
        "venta_licencia_cop": venta_licencia_cop,
        "venta_equipo_licencia_cop": venta_equipo_cop + venta_licencia_cop,
    }


def calcular_servicio_firewall(modelo, vigencia, tipo_instalacion, cantidad_firewalls=1):
    mrc_unitario = TARIFA_SERVICIO_FIREWALL.get(modelo, {}).get(vigencia, 0)
    nrc_unitario = NRC_SERVICIO_FIREWALL.get(tipo_instalacion, 0)

    mrc_total = mrc_unitario * cantidad_firewalls
    nrc_total = nrc_unitario * cantidad_firewalls
    valor_contrato_mrc = mrc_total * vigencia
    valor_total_contrato = valor_contrato_mrc + nrc_total

    return {
        "modelo": modelo,
        "vigencia": vigencia,
        "cantidad_firewalls": cantidad_firewalls,
        "mrc_unitario": mrc_unitario,
        "nrc_unitario": nrc_unitario,
        "mrc_total": mrc_total,
        "nrc_total": nrc_total,
        "valor_contrato_mrc": valor_contrato_mrc,
        "valor_total_contrato": valor_total_contrato,
    }


# ============================================================
# ENCABEZADO CORPORATIVO
# ============================================================

st.markdown('<div class="mc-header">', unsafe_allow_html=True)
header_col1, header_col2 = st.columns([1, 3])

with header_col1:
    if LOGO_FILE.exists():
        st.image(str(LOGO_FILE), width=300)
    else:
        st.warning("Logo no encontrado en assets/logo_mc.png")

with header_col2:
    st.markdown(
        """
        <div class="mc-title">Cotizador Sophos SMB</div>
        <div class="mc-subtitle">
            Herramienta interna de Media Commerce para estimación rápida de servicios,
            licenciamiento y soluciones de seguridad Sophos SMB.
        </div>
        <div class="mc-badge-row">
            <div class="mc-badge">Seguridad Gestionada</div>
            <div class="mc-badge">Sophos SMB</div>
            <div class="mc-badge">Firewall · Endpoint · EDR · XDR</div>
            <div class="mc-badge">Modelo financiero controlado</div>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# INTERFAZ
# ============================================================

df_precios, hoja_usada = cargar_lista_precios()

if df_precios is None:
    st.error(
        "No se encontró la lista de precios. "
        "Valida que el archivo exista en: data/Lista de precios SOPHOS SMB.xlsx"
    )
    st.stop()

columnas, faltantes = mapear_columnas(df_precios)

if faltantes:
    st.error(
        "No se pudieron identificar estas columnas obligatorias en el Excel: "
        + ", ".join(faltantes)
    )
    st.write("Columnas encontradas en el archivo:")
    st.write(list(df_precios.columns))
    st.stop()

st.markdown(
    f"""
    <div class="mc-alert">
        Lista de precios cargada correctamente desde la hoja: <strong>{hoja_usada}</strong>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown('<div class="mc-panel">', unsafe_allow_html=True)
st.markdown('<div class="mc-section-title">Parámetros de cotización</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="mc-section-caption">Seleccione la familia, vigencia, escenario comercial y condiciones del servicio.</div>',
    unsafe_allow_html=True
)

col1, col2, col3, col4 = st.columns(4)

with col1:
    familia = st.selectbox("Familia", FAMILIAS_VALIDAS)

with col2:
    label_cantidad = "Cantidad de usuarios" if familia == "Firewall" else "Cantidad"
    cantidad = st.number_input(
        label_cantidad,
        min_value=1,
        max_value=10000,
        value=20 if familia == "Firewall" else 1,
        step=1
    )

with col3:
    vigencia = st.selectbox("Vigencia", VIGENCIAS_VALIDAS, index=0)

with col4:
    trm = st.number_input("TRM", min_value=1.0, value=3800.0, step=10.0)

tipo_venta_firewall = "Venta one shot"

if familia == "Firewall":
    col_fw_1, col_fw_2, col_fw_3 = st.columns(3)

    with col_fw_1:
        tipo_venta_firewall = st.selectbox(
            "Tipo de venta Firewall",
            TIPOS_VENTA_FIREWALL,
            index=0
        )

    with col_fw_2:
        cantidad_firewalls = st.number_input(
            "Cantidad de firewalls",
            min_value=1,
            max_value=5,
            value=1,
            step=1
        )

    with col_fw_3:
        st.info("Para servicio mensual se usa la tarifa MRC/NRC del marco tarifario.")
else:
    cantidad_firewalls = 1

col5, col6 = st.columns(2)

with col5:
    escenario = st.selectbox(
        "Estado de oportunidad",
        ["Sin DR aprobado", "Con DR aprobado"],
        index=1
    )

with col6:
    tipo_instalacion = st.selectbox(
        "Tipo de instalación",
        ["Sin instalación", "Instalación básica", "Instalación completa"],
        index=1
    )

st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# CÁLCULO SEGÚN FAMILIA
# ============================================================

precio = None
modelo_sugerido = None
capacidad_sugerida = None

if familia == "Firewall":
    modelo_sugerido, capacidad_sugerida = sugerir_firewall(cantidad)

    if modelo_sugerido is None:
        st.error(
            "La cantidad ingresada supera el marco SMB para Firewall. "
            "Este caso requiere revisión técnica/comercial especial."
        )
        st.stop()

    st.markdown(
        f"""
        <div class="mc-alert">
            Equipo sugerido: <strong>Sophos {modelo_sugerido}</strong>
            para hasta <strong>{capacidad_sugerida} usuarios Full UTM</strong>.
        </div>
        """,
        unsafe_allow_html=True
    )

    precio = buscar_precio_firewall(
        df=df_precios,
        columnas=columnas,
        modelo_sugerido=modelo_sugerido,
        vigencia=vigencia
    )
else:
    precio = buscar_precio_licencia(
        df=df_precios,
        columnas=columnas,
        familia=familia,
        cantidad=cantidad,
        vigencia=vigencia
    )

if precio is None:
    st.error(
        "No se encontró precio para la combinación seleccionada. "
        "Revisa familia, cantidad, vigencia o estructura del archivo Excel."
    )
    st.stop()


# ============================================================
# MODO SERVICIO FIREWALL
# ============================================================

if familia == "Firewall" and tipo_venta_firewall == "Servicio mensual":
    servicio = calcular_servicio_firewall(
        modelo=modelo_sugerido,
        vigencia=vigencia,
        tipo_instalacion=tipo_instalacion,
        cantidad_firewalls=cantidad_firewalls
    )

    st.markdown('<div class="mc-panel">', unsafe_allow_html=True)
    st.markdown('<div class="mc-section-title">Resultado general · Servicio mensual</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="mc-section-caption">Resumen económico del servicio en modalidad MRC/NRC.</div>',
        unsafe_allow_html=True
    )

    s1, s2, s3, s4 = st.columns(4)

    with s1:
        render_result_card("MRC mensual", formato_cop(servicio["mrc_total"]), "Cargo mensual recurrente", "blue")

    with s2:
        render_result_card("NRC instalación", formato_cop(servicio["nrc_total"]), "Cargo único de instalación", "yellow")

    with s3:
        render_result_card(
            f"Valor MRC {vigencia} meses",
            formato_cop(servicio["valor_contrato_mrc"]),
            "Valor acumulado del servicio",
            "blue"
        )

    with s4:
        render_result_card("Total contrato", formato_cop(servicio["valor_total_contrato"]), "MRC contrato + NRC", "green")

    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="mc-panel">', unsafe_allow_html=True)
    st.markdown('<div class="mc-section-title">Detalle del servicio firewall</div>', unsafe_allow_html=True)

    tabla_servicio = pd.DataFrame([
        {"Concepto": "Equipo sugerido", "Valor": f"Sophos {modelo_sugerido}"},
        {"Concepto": "Cantidad de usuarios", "Valor": cantidad},
        {"Concepto": "Capacidad de referencia", "Valor": f"Hasta {capacidad_sugerida} usuarios Full UTM"},
        {"Concepto": "Cantidad de firewalls", "Valor": cantidad_firewalls},
        {"Concepto": "Vigencia", "Valor": f"{vigencia} meses"},
        {"Concepto": "MRC unitario", "Valor": formato_cop(servicio["mrc_unitario"])},
        {"Concepto": "MRC total mensual", "Valor": formato_cop(servicio["mrc_total"])},
        {"Concepto": "NRC unitario", "Valor": formato_cop(servicio["nrc_unitario"])},
        {"Concepto": "NRC total", "Valor": formato_cop(servicio["nrc_total"])},
        {"Concepto": "Valor MRC contrato", "Valor": formato_cop(servicio["valor_contrato_mrc"])},
        {"Concepto": "Total contrato", "Valor": formato_cop(servicio["valor_total_contrato"])},
    ])

    st.dataframe(tabla_servicio, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="mc-panel">', unsafe_allow_html=True)
    st.markdown('<div class="mc-section-title">Resumen comercial</div>', unsafe_allow_html=True)

    tabla_resumen_servicio = pd.DataFrame([
        {
            "Ítem": f"Servicio mensual Sophos {modelo_sugerido}",
            "Tipo": "MRC",
            "Cantidad": cantidad_firewalls,
            "Valor unitario": formato_cop(servicio["mrc_unitario"]),
            "Valor mensual": formato_cop(servicio["mrc_total"]),
            "Valor contrato": formato_cop(servicio["valor_contrato_mrc"]),
        },
        {
            "Ítem": "Instalación",
            "Tipo": "NRC",
            "Cantidad": cantidad_firewalls,
            "Valor unitario": formato_cop(servicio["nrc_unitario"]),
            "Valor mensual": "-",
            "Valor contrato": formato_cop(servicio["nrc_total"]),
        },
        {
            "Ítem": "Total a presentar al cliente",
            "Tipo": "Total",
            "Cantidad": "-",
            "Valor unitario": "-",
            "Valor mensual": formato_cop(servicio["mrc_total"]),
            "Valor contrato": formato_cop(servicio["valor_total_contrato"]),
        },
    ])

    st.dataframe(tabla_resumen_servicio, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if servicio["mrc_total"] > 0:
        st.success("Servicio mensual calculado correctamente con base en el marco tarifario.")
    else:
        st.error("No se encontró tarifa MRC para el equipo y vigencia seleccionados.")

    with st.expander("Validación técnica servicio firewall"):
        st.write("Tabla MRC usada:")
        st.write(TARIFA_SERVICIO_FIREWALL)
        st.write("Tabla NRC usada:")
        st.write(NRC_SERVICIO_FIREWALL)
        st.write("Resultado del cálculo:")
        st.write(servicio)

    st.markdown(
        '<div class="mc-footer-note">Media Commerce · Herramienta interna de estimación comercial</div>',
        unsafe_allow_html=True
    )

    st.stop()


# ============================================================
# MODO ONE SHOT / OTRAS FAMILIAS
# ============================================================

msrp_usd = precio["msrp_usd"]
dr_usd = precio["dr_usd"]

resultado = calcular_financiero(
    msrp_usd=msrp_usd,
    dr_usd=dr_usd,
    escenario=escenario,
    trm=trm,
    tipo_instalacion=tipo_instalacion
)

detalle_venta_firewall = None

if familia == "Firewall":
    detalle_venta_firewall = calcular_detalle_venta_firewall(
        precio=precio,
        escenario=escenario,
        trm=trm
    )


# ============================================================
# RESULTADO GENERAL
# ============================================================

st.markdown('<div class="mc-panel">', unsafe_allow_html=True)
st.markdown('<div class="mc-section-title">Resultado general</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="mc-section-caption">Valores estimados para presentar internamente y validar la rentabilidad del negocio.</div>',
    unsafe_allow_html=True
)

if familia == "Firewall":
    m1, m2, m3, m4 = st.columns(4)

    with m1:
        render_result_card("Venta equipo", formato_cop(detalle_venta_firewall["venta_equipo_cop"]), "Hardware Sophos", "blue")

    with m2:
        render_result_card(
            f"Venta licencia {vigencia} meses",
            formato_cop(detalle_venta_firewall["venta_licencia_cop"]),
            "Suscripción seleccionada",
            "blue"
        )

    with m3:
        render_result_card("Venta instalación", formato_cop(resultado["precio_venta_instalacion_cop"]), "Servicio profesional", "yellow")

    with m4:
        render_result_card("Total cliente", formato_cop(resultado["precio_venta_total_cop"]), "Valor total estimado", "green")

    m5, m6, m7, m8 = st.columns(4)

    with m5:
        render_result_card("Costo total", formato_cop(resultado["costo_total_cop"]), "Base de costo", "blue")

    with m6:
        render_result_card("Utilidad EBIT", formato_cop(resultado["utilidad_ebit_cop"]), "Resultado operativo estimado", "green")

    with m7:
        render_result_card("EBIT / Costo", formato_porcentaje(resultado["ebit_sobre_costo"]), "Rentabilidad sobre costo", "green")

    with m8:
        render_result_card("Margen EBIT", formato_porcentaje(resultado["porcentaje_ebit"]), "Sobre precio de venta", "blue")
else:
    m1, m2, m3, m4 = st.columns(4)

    with m1:
        render_result_card(
            "Venta licencias / producto",
            formato_cop(resultado["precio_venta_licencias_cop"]),
            "Valor comercial estimado",
            "blue"
        )

    with m2:
        render_result_card("Venta instalación", formato_cop(resultado["precio_venta_instalacion_cop"]), "Servicio profesional", "yellow")

    with m3:
        render_result_card("Total cliente", formato_cop(resultado["precio_venta_total_cop"]), "Valor total estimado", "green")

    with m4:
        render_result_card("EBIT / Costo", formato_porcentaje(resultado["ebit_sobre_costo"]), "Rentabilidad sobre costo", "green")

    m5, m6, m7, m8 = st.columns(4)

    with m5:
        render_result_card("Costo total", formato_cop(resultado["costo_total_cop"]), "Base de costo", "blue")

    with m6:
        render_result_card("Utilidad EBIT", formato_cop(resultado["utilidad_ebit_cop"]), "Resultado operativo estimado", "green")

    with m7:
        render_result_card("Margen EBIT", formato_porcentaje(resultado["porcentaje_ebit"]), "Sobre precio de venta", "blue")

    with m8:
        render_result_card(
            "Margen contribucional",
            formato_porcentaje(resultado["porcentaje_margen_contribucional"]),
            "Después de impuestos y gastos",
            "blue"
        )

st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# DETALLE FIREWALL
# ============================================================

if familia == "Firewall":
    st.markdown('<div class="mc-panel">', unsafe_allow_html=True)
    st.markdown('<div class="mc-section-title">Detalle Firewall: equipo y licencia</div>', unsafe_allow_html=True)

    detalle_fw = precio["detalle_firewall"].copy()

    detalle_fw["MSRP COP"] = [
        formato_cop(precio["msrp_equipo_usd"] * trm),
        formato_cop(precio["msrp_licencia_usd"] * trm),
        formato_cop(precio["msrp_usd"] * trm),
    ]

    detalle_fw["DR COP"] = [
        formato_cop(precio["dr_equipo_usd"] * trm),
        formato_cop(precio["dr_licencia_usd"] * trm),
        formato_cop(precio["dr_usd"] * trm),
    ]

    detalle_fw["Venta COP"] = [
        formato_cop(detalle_venta_firewall["venta_equipo_cop"]),
        formato_cop(detalle_venta_firewall["venta_licencia_cop"]),
        formato_cop(detalle_venta_firewall["venta_equipo_licencia_cop"]),
    ]

    st.dataframe(detalle_fw, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# DETALLE DE PRODUCTO
# ============================================================

st.markdown('<div class="mc-panel">', unsafe_allow_html=True)
st.markdown('<div class="mc-section-title">Detalle de producto cotizado</div>', unsafe_allow_html=True)

detalle_producto = {
    "Familia": familia,
    "Producto": precio["producto"],
    "Cantidad de usuarios" if familia == "Firewall" else "Cantidad": cantidad,
    "Vigencia": f"{vigencia} meses",
    "Escenario": escenario,
    "Tipo instalación": tipo_instalacion,
    "SKU(s)": ", ".join(precio["skus"]) if precio["skus"] else "-",
    "MSRP USD total": formato_usd(msrp_usd),
    "DR USD total": formato_usd(dr_usd),
    "MSRP USD unitario": formato_usd(precio["msrp_unitario_usd"]),
    "DR USD unitario": formato_usd(precio["dr_unitario_usd"]),
    "TRM": f"{trm:,.2f}",
    "Costo licencias/producto COP": formato_cop(resultado["costo_licencias_cop"]),
    "Venta licencias/producto COP": formato_cop(resultado["precio_venta_licencias_cop"]),
    "Costo instalación COP": formato_cop(resultado["costo_instalacion_cop"]),
    "Venta instalación COP": formato_cop(resultado["precio_venta_instalacion_cop"]),
    "Total cliente COP": formato_cop(resultado["precio_venta_total_cop"]),
}

if familia == "Firewall":
    detalle_producto.update({
        "Tipo de venta Firewall": tipo_venta_firewall,
        "Cantidad de firewalls": cantidad_firewalls,
        "Venta equipo COP": formato_cop(detalle_venta_firewall["venta_equipo_cop"]),
        f"Venta licencia {vigencia} meses COP": formato_cop(detalle_venta_firewall["venta_licencia_cop"]),
    })

st.dataframe(pd.DataFrame([detalle_producto]), use_container_width=True, hide_index=True)
st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# RESUMEN COMERCIAL
# ============================================================

st.markdown('<div class="mc-panel">', unsafe_allow_html=True)
st.markdown('<div class="mc-section-title">Resumen comercial</div>', unsafe_allow_html=True)

if familia == "Firewall":
    tabla_resumen_cliente = pd.DataFrame([
        {
            "Concepto": f"Equipo Sophos {precio['producto']}",
            "Valor COP": formato_cop(detalle_venta_firewall["venta_equipo_cop"])
        },
        {
            "Concepto": f"Licencia Sophos {precio['producto']} - {vigencia} meses",
            "Valor COP": formato_cop(detalle_venta_firewall["venta_licencia_cop"])
        },
        {
            "Concepto": "Instalación",
            "Valor COP": formato_cop(resultado["precio_venta_instalacion_cop"])
        },
        {
            "Concepto": "Total a presentar al cliente",
            "Valor COP": formato_cop(resultado["precio_venta_total_cop"])
        },
    ])
else:
    tabla_resumen_cliente = pd.DataFrame([
        {
            "Concepto": "Licencias / producto",
            "Valor COP": formato_cop(resultado["precio_venta_licencias_cop"])
        },
        {
            "Concepto": "Instalación",
            "Valor COP": formato_cop(resultado["precio_venta_instalacion_cop"])
        },
        {
            "Concepto": "Total a presentar al cliente",
            "Valor COP": formato_cop(resultado["precio_venta_total_cop"])
        },
    ])

st.dataframe(tabla_resumen_cliente, use_container_width=True, hide_index=True)
st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# RESULTADO FINANCIERO
# ============================================================

st.markdown('<div class="mc-panel">', unsafe_allow_html=True)
st.markdown('<div class="mc-section-title">Resultado financiero del negocio</div>', unsafe_allow_html=True)

tabla_financiera = pd.DataFrame([
    {
        "Concepto": "Precio de venta total",
        "%": "100,0%",
        "Valor COP": formato_cop(resultado["precio_venta_total_cop"])
    },
    {
        "Concepto": "Costo total",
        "%": formato_porcentaje(resultado["porcentaje_costo"]),
        "Valor COP": formato_cop(resultado["costo_total_cop"])
    },
    {
        "Concepto": "Utilidad bruta",
        "%": formato_porcentaje(resultado["porcentaje_utilidad_bruta"]),
        "Valor COP": formato_cop(resultado["utilidad_bruta_cop"])
    },
    {
        "Concepto": "Impuestos",
        "%": formato_porcentaje(IMPUESTOS),
        "Valor COP": formato_cop(resultado["impuestos_cop"])
    },
    {
        "Concepto": "Gastos",
        "%": formato_porcentaje(GASTOS),
        "Valor COP": formato_cop(resultado["gastos_cop"])
    },
    {
        "Concepto": "Margen contribucional",
        "%": formato_porcentaje(resultado["porcentaje_margen_contribucional"]),
        "Valor COP": formato_cop(resultado["margen_contribucional_cop"])
    },
    {
        "Concepto": "Carga prestacional de la comisión",
        "%": formato_porcentaje(CARGA_PRESTACIONAL_COMISION),
        "Valor COP": formato_cop(resultado["carga_prestacional_cop"])
    },
    {
        "Concepto": "Comisión",
        "%": formato_porcentaje(COMISION),
        "Valor COP": formato_cop(resultado["comision_cop"])
    },
    {
        "Concepto": "Utilidad EBIT",
        "%": formato_porcentaje(resultado["porcentaje_ebit"]),
        "Valor COP": formato_cop(resultado["utilidad_ebit_cop"])
    },
    {
        "Concepto": "EBIT / Costo",
        "%": formato_porcentaje(resultado["ebit_sobre_costo"]),
        "Valor COP": "-"
    },
])

st.dataframe(tabla_financiera, use_container_width=True, hide_index=True)
st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# VALIDACIÓN COMERCIAL
# ============================================================

st.markdown('<div class="mc-panel">', unsafe_allow_html=True)
st.markdown('<div class="mc-section-title">Validación comercial</div>', unsafe_allow_html=True)

ebit_costo = resultado["ebit_sobre_costo"]

if ebit_costo >= 0.15:
    st.success("Negocio viable: EBIT / Costo igual o superior al 15%.")
elif ebit_costo >= 0.08:
    st.warning("Negocio revisable: rentabilidad aceptable, pero requiere validación comercial.")
else:
    st.error("Negocio sensible: rentabilidad baja. Requiere revisión antes de presentar al cliente.")

st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# VALIDACIÓN TÉCNICA
# ============================================================

with st.expander("Validación técnica del cálculo"):
    st.write("Parámetros financieros usados:")
    st.write({
        "Porcentaje costo sobre venta": formato_porcentaje(PORCENTAJE_COSTO_SOBRE_VENTA),
        "Impuestos": formato_porcentaje(IMPUESTOS),
        "Gastos": formato_porcentaje(GASTOS),
        "Carga prestacional comisión": formato_porcentaje(CARGA_PRESTACIONAL_COMISION),
        "Comisión": formato_porcentaje(COMISION),
        "Costo instalación sobre venta licencias/producto": formato_porcentaje(PORCENTAJE_COSTO_INSTALACION),
    })

    valores_base = {
        "MSRP USD total": msrp_usd,
        "DR USD total": dr_usd,
        "TRM": trm,
        "Costo licencias/producto COP": resultado["costo_licencias_cop"],
        "Venta licencias/producto COP": resultado["precio_venta_licencias_cop"],
        "Costo instalación COP": resultado["costo_instalacion_cop"],
        "Venta instalación COP": resultado["precio_venta_instalacion_cop"],
        "Precio venta total COP": resultado["precio_venta_total_cop"],
        "Costo total COP": resultado["costo_total_cop"],
    }

    if familia == "Firewall":
        valores_base.update({
            "Tipo venta Firewall": tipo_venta_firewall,
            "Cantidad firewalls": cantidad_firewalls,
            "Venta equipo COP": detalle_venta_firewall["venta_equipo_cop"],
            f"Venta licencia {vigencia} meses COP": detalle_venta_firewall["venta_licencia_cop"],
        })

    st.write(valores_base)

with st.expander("Ver filas usadas del Excel"):
    st.dataframe(precio["filas"], use_container_width=True)

st.markdown(
    '<div class="mc-footer-note">Media Commerce · Herramienta interna de estimación comercial · Seguridad Sophos SMB</div>',
    unsafe_allow_html=True
)