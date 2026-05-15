# -*- coding: utf-8 -*-
"""
APP COSTO POR EQUIPO - V14.0.0

Streamlit app para cruzar:
1) Repuestos instalados
2) Uno o varios archivos de antigüedad / fabricación por modelo
3) Spare parts / precios Option 2

Entrega en pantalla:
- Dashboard ejecutivo estilo Excel
- KPIs
- Gráfica de TODOS los instrumentos, no solo Top 10
- Donut por modelo
- Tabla resumen y detalle

Entrega en Excel:
- Dashboard con TODOS los instrumentos
- Resumen por equipo
- Detalle repuestos
- Validaciones
- Antigüedad consolidada
- Diagnóstico de archivos

Ejecución:
    python -m pip install -r requirements_costo_por_equipo_v9.txt
    python -m streamlit run app_costo_por_equipo_v9_dashboard_profesional.py
"""

from __future__ import annotations

import io
import re
import zipfile
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# =============================================================================
# CONFIGURACIÓN GENERAL
# =============================================================================

APP_VERSION = "13.0.0"
APP_TITLE = "Costo de repuestos por equipo"
APP_SUBTITLE = "Dashboard ejecutivo en una sola página · Excel final intacto · multiarchivo antigüedad · Option 2"

NEON_CYAN = "#00D9FF"
NEON_GREEN = "#2CFF9B"
NEON_YELLOW = "#FFD166"
NEON_PURPLE = "#A855F7"
NEON_PINK = "#FF4D8D"
NEON_ORANGE = "#FF9F1C"
DARK_BG = "#07111F"
PANEL = "rgba(255,255,255,0.09)"
PANEL_BORDER = "rgba(255,255,255,0.22)"
TEXT = "#F4F7FB"
MUTED = "#A8B3C7"

CHART_COLORS = [NEON_CYAN, NEON_GREEN, NEON_YELLOW, NEON_PURPLE, NEON_PINK, NEON_ORANGE]

st.set_page_config(
    page_title=f"{APP_TITLE} · v{APP_VERSION}",
    page_icon="💠",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =============================================================================
# CSS / UI
# =============================================================================

def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            --excel-navy: #07111F;
            --excel-blue: #10324E;
            --excel-cyan: #00D9FF;
            --excel-green: #2CFF9B;
            --excel-yellow: #FFD166;
            --excel-red: #FF4D8D;
            --excel-bg: #F4F7FB;
            --excel-text: #172033;
            --excel-muted: #65748B;
            --excel-border: #D8E2EF;
        }
        .stApp {
            background: #F4F7FB;
            color: var(--excel-text);
            font-family: "Segoe UI", Arial, sans-serif;
        }
        .main .block-container {
            max-width: 1500px;
            padding-top: 1.0rem;
            padding-bottom: 1.5rem;
        }
        section[data-testid="stSidebar"] {
            background: #07111F;
            border-right: 1px solid #10324E;
        }
        section[data-testid="stSidebar"] * { color: #EAF3FF; }
        section[data-testid="stSidebar"] div[data-baseweb="select"] * { color: #172033 !important; }
        section[data-testid="stSidebar"] input { color: #172033 !important; }
        .app-hero {
            background: #07111F;
            color: white;
            border-radius: 6px;
            padding: 12px 16px;
            margin-bottom: 10px;
            border: 1px solid #10324E;
        }
        .app-hero h1 {
            margin: 0;
            font-size: 24px;
            font-weight: 850;
            letter-spacing: .2px;
        }
        .app-hero p {
            margin: 3px 0 0 0;
            color: #A8B3C7;
            font-size: 13px;
        }
        .version-pill {
            display: inline-block;
            margin-top: 7px;
            padding: 4px 8px;
            background: #00D9FF;
            color: #00111D;
            border-radius: 4px;
            font-weight: 850;
            font-size: 11px;
        }
        .excel-title {
            background: #07111F;
            color: #FFFFFF;
            padding: 10px 12px;
            border-radius: 5px 5px 0 0;
            font-weight: 900;
            font-size: 15px;
            border: 1px solid #07111F;
            margin-top: 8px;
        }
        .excel-body {
            background: #FFFFFF;
            border: 1px solid #D8E2EF;
            border-top: 0;
            border-radius: 0 0 5px 5px;
            padding: 10px 12px;
            margin-bottom: 10px;
            color: #172033;
        }
        .upload-box {
            background: #FFFFFF;
            border: 1px solid #D8E2EF;
            border-radius: 6px;
            padding: 10px 12px;
            min-height: 96px;
            box-shadow: 0 2px 8px rgba(7,17,31,0.04);
        }
        .upload-title {
            color: #10324E;
            font-weight: 900;
            font-size: 14px;
            margin-bottom: 2px;
        }
        .upload-note {
            color: #65748B;
            font-size: 12px;
            line-height: 1.25;
            margin-bottom: 6px;
        }
        .kpi-card {
            background: #FFFFFF;
            border: 1px solid #D8E2EF;
            border-radius: 5px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(7,17,31,0.05);
            margin-bottom: 6px;
        }
        .kpi-label {
            background: #00D9FF;
            color: #00111D;
            font-weight: 900;
            text-align: center;
            padding: 6px 4px;
            font-size: 12px;
            border-bottom: 1px solid #D8E2EF;
        }
        .kpi-value {
            background: #10324E;
            color: #FFFFFF;
            text-align: center;
            padding: 11px 4px;
            font-size: 20px;
            font-weight: 900;
        }
        .sheet-note {
            background: #FFFFFF;
            border: 1px solid #D8E2EF;
            border-left: 5px solid #00D9FF;
            border-radius: 5px;
            padding: 9px 12px;
            color: #374151;
            margin: 8px 0 10px 0;
            font-size: 13px;
        }
        .status-warning {
            background: #FFF7E6;
            border: 1px solid #FFD166;
            color: #6B4600;
            border-radius: 5px;
            padding: 9px 12px;
            font-weight: 800;
            margin-top: 8px;
        }
        .status-ok {
            background: #E8FFF3;
            border: 1px solid #2CFF9B;
            color: #0A4B2A;
            border-radius: 5px;
            padding: 9px 12px;
            font-weight: 800;
            margin-top: 8px;
        }
        .stTabs [data-baseweb="tab-list"] { gap: 4px; }
        .stTabs [data-baseweb="tab"] {
            background: white;
            border: 1px solid #D8E2EF;
            border-radius: 5px 5px 0 0;
            color: #10324E;
            font-weight: 750;
            padding: 7px 12px;
        }
        .stTabs [aria-selected="true"] {
            background: #10324E !important;
            color: white !important;
            border-color: #10324E !important;
        }
        .stDownloadButton button, .stButton button {
            border-radius: 5px !important;
            border: 1px solid #00A7C7 !important;
            background: #00D9FF !important;
            color: #00111D !important;
            font-weight: 900 !important;
            box-shadow: none !important;
        }
        div[data-testid="stFileUploader"] section {
            border: 1px dashed #00A7C7 !important;
            background: #FFFFFF !important;
            border-radius: 5px !important;
            padding: 8px !important;
        }
        div[data-testid="stDataFrame"] {
            border: 1px solid #D8E2EF;
            border-radius: 5px;
            overflow: hidden;
        }
        div[data-testid="stVerticalBlock"] { gap: 0.55rem; }
        h1, h2, h3 { color: #10324E; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def hero() -> None:
    st.markdown(
        f"""
        <div class="app-hero">
            <h1>Reporte de repuestos consumidos por equipo</h1>
            <p>{APP_SUBTITLE}</p>
            <span class="version-pill">Versión {APP_VERSION} · pantalla compacta basada en el Dashboard del Excel</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def glass_section(title: str, body: Optional[str] = None) -> None:
    st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)
    if body:
        st.markdown(f"<div class='small-note'>{body}</div>", unsafe_allow_html=True)


# =============================================================================
# NORMALIZACIÓN Y LECTURA DE ARCHIVOS
# =============================================================================

def normalize_text(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def first_str(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    return "" if text.lower() in ["nan", "nat", "none"] else text


def clean_serial(value: object) -> str:
    text = first_str(value).replace("\u00a0", " ")
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\s+", "", text).upper()


def clean_part(value: object) -> str:
    text = first_str(value).replace("\u00a0", " ")
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    text = text.replace("*", "").replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", "", text).upper()


def parse_number(value: object) -> float:
    if value is None:
        return np.nan
    try:
        if pd.isna(value):
            return np.nan
    except Exception:
        pass
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    text = str(value).strip()
    text = text.replace("€", "").replace("$", "").replace("EUR", "").replace("USD", "")
    text = re.sub(r"[^0-9,\.\-]", "", text)
    if not text or text in ["-", ".", ","]:
        return np.nan
    if "," in text and "." in text:
        # Decide separador decimal por el último símbolo
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        text = text.replace(",", ".") if len(parts) == 2 and len(parts[-1]) <= 2 else text.replace(",", "")
    try:
        return float(text)
    except Exception:
        return np.nan


def parse_qty(value: object) -> float:
    n = parse_number(value)
    return 1.0 if pd.isna(n) else n


def make_unique_columns(cols: Iterable[object]) -> List[str]:
    output: List[str] = []
    seen: Dict[str, int] = {}
    for idx, col in enumerate(cols, start=1):
        name = first_str(col)
        if not name or normalize_text(name).startswith("unnamed"):
            name = f"Columna_{idx}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        output.append(name)
    return output


HEADER_KEYWORDS = [
    "serial", "serie", "instrument", "equipo", "part", "spare", "pn", "referencia", "codigo",
    "producto", "description", "descripcion", "quantity", "cantidad", "qty", "orden", "service",
    "cliente", "customer", "fecha", "date", "fabricacion", "installation", "instalacion", "precio",
    "price", "cost", "costo", "opt2", "option", "opcion", "eur", "currency", "moneda",
    "model", "modelo", "distributor", "distribuidor", "country", "pais", "status", "estado",
]


def header_row_score(row: pd.Series) -> float:
    values = [normalize_text(x) for x in row.tolist()]
    non_empty = sum(1 for x in values if x)
    keyword_hits = sum(1 for cell in values for kw in HEADER_KEYWORDS if kw in cell)
    numeric_like = sum(1 for cell in values if cell and re.fullmatch(r"[0-9.,\-/]+", cell))
    return non_empty + keyword_hits * 5 - numeric_like * 0.85


def promote_header(raw_df: pd.DataFrame) -> pd.DataFrame:
    raw = raw_df.dropna(axis=0, how="all").dropna(axis=1, how="all").copy()
    if raw.empty:
        return raw
    best_pos, best_score = 0, -9999.0
    for pos in range(min(25, len(raw))):
        score = header_row_score(raw.iloc[pos])
        if score > best_score:
            best_pos, best_score = pos, score
    header = make_unique_columns(raw.iloc[best_pos].tolist())
    df = raw.iloc[best_pos + 1 :].copy()
    df.columns = header[: len(df.columns)]
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
    return df.reset_index(drop=True)


def read_uploaded_file(uploaded_file) -> Dict[str, pd.DataFrame]:
    name = uploaded_file.name.lower()
    uploaded_file.seek(0)
    if name.endswith((".csv", ".txt")):
        content = uploaded_file.read()
        for encoding in ["utf-8-sig", "utf-8", "latin-1", "cp1252"]:
            try:
                raw = pd.read_csv(io.BytesIO(content), header=None, dtype=object, sep=None, engine="python", encoding=encoding)
                return {"CSV": promote_header(raw)}
            except Exception:
                continue
        raise ValueError(f"No fue posible leer el archivo CSV/TXT: {uploaded_file.name}")
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        raw_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None, dtype=object)
        return {str(sheet): promote_header(df) for sheet, df in raw_sheets.items()}
    raise ValueError(f"Formato no soportado: {uploaded_file.name}")


def read_local_file(path: str) -> Dict[str, pd.DataFrame]:
    name = path.lower()
    if name.endswith((".csv", ".txt")):
        with open(path, "rb") as fh:
            content = fh.read()
        for encoding in ["utf-8-sig", "utf-8", "latin-1", "cp1252"]:
            try:
                raw = pd.read_csv(io.BytesIO(content), header=None, dtype=object, sep=None, engine="python", encoding=encoding)
                return {"CSV": promote_header(raw)}
            except Exception:
                continue
        raise ValueError(f"No fue posible leer el archivo CSV/TXT: {path}")
    raw_sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=object)
    return {str(sheet): promote_header(df) for sheet, df in raw_sheets.items()}


# =============================================================================
# DETECCIÓN DE COLUMNAS
# =============================================================================

SERIAL_ALIASES = ["serial", "serial number", "serie", "numero de serie", "nro de serie", "instrument serial", "serial equipo", "sn"]
PART_ALIASES = ["part number", "part no", "pn", "spare part", "referencia", "codigo repuesto", "numero de parte", "material", "item code", "code", "part"]
PRODUCT_ALIASES = ["producto", "product", "description", "descripcion", "part description", "product description", "material description", "descripcion material"]
QTY_ALIASES = ["cantidad", "quantity", "qty", "cant", "suma de cantidad", "cantidad instalada", "installed quantity", "cantidad consumida"]
ORDER_ALIASES = ["orden de servicio", "service order", "orden", "work order", "ss0", "sso", "ticket"]
CUSTOMER_ALIASES = ["nombre cliente", "cliente", "customer", "customer name", "client", "laboratorio", "institucion"]
DATE_ALIASES = ["fecha referencia", "fecha de fabricacion", "fecha fabricacion", "manufacturing date", "manufacture date", "fecha instalacion", "installation date", "install date", "date", "fecha"]
MODEL_ALIASES = ["modelo", "model", "instrument model", "tipo instrumento", "instrument type", "platform", "plataforma"]
DISTRIBUTOR_ALIASES = ["distribuidor", "distributor", "distribuidor records", "distribuidor ventas", "dealer"]
COUNTRY_ALIASES = ["pais", "país", "country"]
STATUS_ALIASES = ["estado rutina", "routine status", "estado instrumento", "instrument status", "status", "estado"]
PRICE_ALIASES = ["tp opt2", "opt2", "opt 2", "option 2", "opcion 2", "precio opt2", "price opt2", "unit price", "precio unitario", "price", "precio", "cost", "costo", "eur"]
CURRENCY_ALIASES = ["currency", "moneda", "divisa"]
ALT_CODE_ALIASES = ["alternative codes", "alternative code", "codigo alternativo", "codigos alternativos", "alternate", "replacement", "old code", "replaced"]


def alias_score(column_name: object, aliases: List[str]) -> float:
    col = normalize_text(column_name)
    if not col:
        return 0.0
    best = 0.0
    for alias in aliases:
        a = normalize_text(alias)
        if not a:
            continue
        if col == a:
            best = max(best, 100.0)
        elif a in col:
            best = max(best, 88.0)
        else:
            tokens = a.split()
            if tokens and all(t in col for t in tokens):
                best = max(best, 75.0 + len(tokens))
            elif tokens:
                partial = sum(t in col for t in tokens) / len(tokens)
                if partial >= 0.67:
                    best = max(best, 48.0 + partial * 22)
    return best


def find_best_column(df: pd.DataFrame, aliases: List[str], min_score: float = 45.0) -> Optional[str]:
    if df is None or df.empty:
        return None
    scored = [(c, alias_score(c, aliases)) for c in df.columns]
    col, score = max(scored, key=lambda x: x[1])
    return str(col) if score >= min_score else None


def price_column_score(df: pd.DataFrame, col: str, prefer_option2: bool = False) -> float:
    score = alias_score(col, PRICE_ALIASES)
    norm = normalize_text(col)
    if any(x in norm for x in ["opt2", "opt 2", "option 2", "opcion 2"]):
        score += 75
    elif prefer_option2 and any(x in norm for x in ["price", "precio", "cost", "costo", "eur"]):
        score += 30
    sample = df[col].dropna().head(400)
    if len(sample):
        parsed = sample.map(parse_number)
        score += parsed.notna().mean() * 32 + (parsed > 0).mean() * 20
    return score


def find_best_price_column(df: pd.DataFrame, prefer_option2: bool = False) -> Optional[str]:
    if df is None or df.empty:
        return None
    col, score = max(((str(c), price_column_score(df, str(c), prefer_option2)) for c in df.columns), key=lambda x: x[1])
    return col if score >= 55 else None


def choose_default_sheet(sheets: Dict[str, pd.DataFrame], purpose: str) -> Optional[str]:
    scores: List[Tuple[str, float]] = []
    for name, df in sheets.items():
        norm_name = normalize_text(name)
        if df is None or df.empty:
            scores.append((name, -9999.0))
            continue
        score = min(len(df), 1000) / 1000
        if purpose == "repuestos":
            score += 45 if find_best_column(df, SERIAL_ALIASES) else 0
            score += 45 if find_best_column(df, PART_ALIASES) else 0
            score += 18 if find_best_column(df, QTY_ALIASES) else 0
        elif purpose == "antiguedad":
            score += 45 if find_best_column(df, SERIAL_ALIASES) else 0
            score += 45 if find_best_column(df, DATE_ALIASES) else 0
            score += 15 if any(x in norm_name for x in ["antig", "age", "rutina", "cruce", "completo", "base", "installed"] ) else 0
        elif purpose == "precios":
            prefer = any(x in norm_name for x in ["option 2", "opcion 2", "opt2", "opt 2", "spare parts price", "standard"])
            score += 45 if prefer else 0
            score += 35 if find_best_column(df, PART_ALIASES) else 0
            score += 45 if find_best_price_column(df, prefer) else 0
        scores.append((name, score))
    return max(scores, key=lambda x: x[1])[0] if scores else None


def select_options(df: pd.DataFrame, allow_none: bool = True) -> List[str]:
    values = [str(c) for c in df.columns] if df is not None else []
    return ["-- No usar --"] + values if allow_none else values


def index_for(options: List[str], detected: Optional[str]) -> int:
    return options.index(detected) if detected in options else 0


# =============================================================================
# PRECIOS Y ANTIGÜEDAD
# =============================================================================

def first_non_empty(series: pd.Series) -> str:
    for value in series:
        text = first_str(value)
        if text:
            return text
    return ""


def unique_join(series: pd.Series) -> str:
    values: List[str] = []
    for value in series:
        text = first_str(value)
        if text and text not in values:
            values.append(text)
    return ", ".join(values)


def similarity(a: str, b: str) -> float:
    a_norm, b_norm = normalize_text(a), normalize_text(b)
    if not a_norm or not b_norm:
        return 0.0
    return SequenceMatcher(None, a_norm, b_norm).ratio()


def split_alt_codes(value: object) -> List[str]:
    text = first_str(value)
    if not text:
        return []
    parts = re.split(r"[;,\n\r|/]+", text)
    output: List[str] = []
    for part in parts:
        code = clean_part(part)
        if code and code not in output:
            output.append(code)
    return output


def build_price_records(
    price_df: pd.DataFrame,
    part_col: str,
    price_col: str,
    product_col: Optional[str],
    currency_col: Optional[str],
    alt_col: Optional[str],
    source_sheet: str,
) -> Tuple[Dict[str, dict], Dict[str, dict], pd.DataFrame]:
    records: List[dict] = []
    for idx, row in price_df.reset_index(drop=True).iterrows():
        part = clean_part(row.get(part_col, ""))
        if not part:
            continue
        records.append(
            {
                "Part Number Precio": part,
                "Precio unitario OPT2": parse_number(row.get(price_col, np.nan)),
                "Producto precio": first_str(row.get(product_col, "")) if product_col and product_col != "-- No usar --" else "",
                "Moneda": first_str(row.get(currency_col, "EUR")) if currency_col and currency_col != "-- No usar --" else "EUR",
                "Fuente precio": source_sheet,
                "Fila precio": idx + 2,
            }
        )
    price_table = pd.DataFrame(records)
    if price_table.empty:
        return {}, {}, price_table

    price_table["_has_price"] = price_table["Precio unitario OPT2"].notna().astype(int)
    price_table = price_table.sort_values(["Part Number Precio", "_has_price"], ascending=[True, False])

    exact: Dict[str, dict] = {}
    for _, rec in price_table.iterrows():
        key = rec["Part Number Precio"]
        if key not in exact:
            d = rec.to_dict()
            d["Moneda"] = d["Moneda"] if d.get("Moneda") else "EUR"
            exact[key] = d

    alt: Dict[str, dict] = {}
    if alt_col and alt_col != "-- No usar --" and alt_col in price_df.columns:
        for _, row in price_df.reset_index(drop=True).iterrows():
            part = clean_part(row.get(part_col, ""))
            if not part or part not in exact:
                continue
            for alt_code in split_alt_codes(row.get(alt_col, "")):
                if alt_code not in alt:
                    alt[alt_code] = exact[part]

    duplicates = price_table.groupby("Part Number Precio", as_index=False).size().rename(columns={"size": "Registros en lista precios"})
    duplicates = duplicates[duplicates["Registros en lista precios"] > 1].copy()
    return exact, alt, duplicates


def lookup_price(ref: str, product: str, exact: Dict[str, dict], alt: Dict[str, dict]) -> dict:
    key = clean_part(ref)
    if not key:
        return {"Referencia usada para precio": np.nan, "Precio unitario OPT2": np.nan, "Moneda": "EUR", "Estado precio": "Referencia vacía", "Fuente precio": np.nan}

    if key in exact:
        rec = exact[key]
        return {
            "Referencia usada para precio": rec["Part Number Precio"],
            "Precio unitario OPT2": rec["Precio unitario OPT2"],
            "Moneda": rec.get("Moneda", "EUR") or "EUR",
            "Estado precio": "Exacto",
            "Fuente precio": f"{rec['Fuente precio']} fila {rec['Fila precio']}",
        }

    if key in alt:
        rec = alt[key]
        return {
            "Referencia usada para precio": rec["Part Number Precio"],
            "Precio unitario OPT2": rec["Precio unitario OPT2"],
            "Moneda": rec.get("Moneda", "EUR") or "EUR",
            "Estado precio": f"Código alternativo ({key} → {rec['Part Number Precio']})",
            "Fuente precio": f"{rec['Fuente precio']} fila {rec['Fila precio']}",
        }

    # Inferencia controlada para códigos sin sufijo.
    for suffix in ["XL", "XS", "XLR", "XSR", "R", "G", "ETI", "ETIR"]:
        candidate = f"{key}{suffix}"
        if candidate in exact:
            rec = exact[candidate]
            state = f"Inferido por descripción ({key} → {candidate})" if similarity(product, rec.get("Producto precio", "")) >= 0.50 else f"Inferido por sufijo ({key} → {candidate})"
            return {
                "Referencia usada para precio": rec["Part Number Precio"],
                "Precio unitario OPT2": rec["Precio unitario OPT2"],
                "Moneda": rec.get("Moneda", "EUR") or "EUR",
                "Estado precio": state,
                "Fuente precio": f"{rec['Fuente precio']} fila {rec['Fila precio']}",
            }

    return {"Referencia usada para precio": np.nan, "Precio unitario OPT2": np.nan, "Moneda": "EUR", "Estado precio": "Precio no encontrado", "Fuente precio": np.nan}


def consolidate_age_files(age_files) -> Tuple[pd.DataFrame, pd.DataFrame]:
    all_rows: List[pd.DataFrame] = []
    diagnostics: List[dict] = []

    for uploaded in age_files:
        try:
            sheets = read_uploaded_file(uploaded)
            sheet_name = choose_default_sheet(sheets, "antiguedad")
            if sheet_name is None:
                diagnostics.append({"Archivo": uploaded.name, "Hoja usada": "", "Estado": "Sin hojas detectadas"})
                continue
            df = sheets[sheet_name]
            serial_col = find_best_column(df, SERIAL_ALIASES)
            date_col = find_best_column(df, DATE_ALIASES)
            model_col = find_best_column(df, MODEL_ALIASES)
            dist_col = find_best_column(df, DISTRIBUTOR_ALIASES)
            country_col = find_best_column(df, COUNTRY_ALIASES)
            customer_col = find_best_column(df, CUSTOMER_ALIASES)
            status_col = find_best_column(df, STATUS_ALIASES)

            diagnostics.append(
                {
                    "Archivo": uploaded.name,
                    "Hoja usada": sheet_name,
                    "Filas leídas": len(df),
                    "Serial detectado": serial_col or "NO DETECTADO",
                    "Fecha detectada": date_col or "NO DETECTADA",
                    "Modelo detectado": model_col or "",
                    "Distribuidor detectado": dist_col or "",
                    "País detectado": country_col or "",
                    "Cliente detectado": customer_col or "",
                    "Estado": "OK" if serial_col and date_col else "REVISAR: falta serial o fecha",
                }
            )
            if not serial_col:
                continue

            out = pd.DataFrame()
            out["Serial equipo"] = df[serial_col].map(clean_serial)
            out["Fecha referencia instrumento"] = pd.to_datetime(df[date_col], errors="coerce") if date_col else pd.NaT
            out["Modelo"] = df[model_col].map(first_str) if model_col else ""
            out["Distribuidor"] = df[dist_col].map(first_str) if dist_col else ""
            out["País"] = df[country_col].map(first_str) if country_col else ""
            out["Cliente según antigüedad"] = df[customer_col].map(first_str) if customer_col else ""
            out["Estado rutina"] = df[status_col].map(first_str) if status_col else ""
            out["Fuente antigüedad"] = f"{uploaded.name} · {sheet_name}"
            out = out[out["Serial equipo"].astype(bool)].copy()
            all_rows.append(out)
        except Exception as exc:
            diagnostics.append({"Archivo": getattr(uploaded, "name", "archivo"), "Hoja usada": "", "Estado": f"ERROR: {exc}"})

    diag_df = pd.DataFrame(diagnostics)
    if not all_rows:
        return pd.DataFrame(columns=["Serial equipo", "Fecha referencia instrumento", "Modelo", "Distribuidor", "País", "Cliente según antigüedad", "Estado rutina", "Fuente antigüedad"]), diag_df

    age = pd.concat(all_rows, ignore_index=True)
    # Se prioriza la fila con fecha válida más antigua, útil como proxy de fabricación/instalación.
    age["_fecha_valida"] = age["Fecha referencia instrumento"].notna().astype(int)
    age = age.sort_values(["Serial equipo", "_fecha_valida", "Fecha referencia instrumento"], ascending=[True, False, True])
    age = age.drop_duplicates("Serial equipo", keep="first").drop(columns=["_fecha_valida"])
    return age.reset_index(drop=True), diag_df


def semaforo_from_date(value: object) -> str:
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return "SIN FECHA"
    years = (pd.Timestamp.today().normalize() - dt).days / 365.25
    if years >= 9:
        return "ROJO"
    if years >= 5:
        return "NARANJA"
    return "VERDE"


def calculate_age_years(value: object) -> float:
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return np.nan
    return round((pd.Timestamp.today().normalize() - dt).days / 365.25, 2)


# =============================================================================
# PROCESAMIENTO CENTRAL
# =============================================================================

def process_data(
    rep_df: pd.DataFrame,
    age_map: pd.DataFrame,
    price_df: pd.DataFrame,
    price_sheet: str,
    c: Dict[str, Optional[str]],
    age_diag: pd.DataFrame,
) -> Dict[str, object]:
    required = [c.get("serial_rep"), c.get("part_rep"), c.get("part_price"), c.get("unit_price")]
    if any(not x or x == "-- No usar --" for x in required):
        raise ValueError("Faltan columnas obligatorias en repuestos o precios. Revisa la configuración.")
    if age_map.empty:
        raise ValueError("No se pudo consolidar ningún archivo de antigüedad/fabricación con serial válido.")

    r = rep_df.copy()
    r["Serial equipo"] = r[c["serial_rep"]].map(clean_serial)
    r["Referencia consumida"] = r[c["part_rep"]].map(clean_part)
    r["Cantidad consumida"] = r[c["qty_rep"]].map(parse_qty) if c.get("qty_rep") and c["qty_rep"] != "-- No usar --" else 1.0
    r["Producto consumido"] = r[c["product_rep"]].map(first_str) if c.get("product_rep") and c["product_rep"] != "-- No usar --" else ""
    r["Órdenes de servicio"] = r[c["order_rep"]].map(first_str) if c.get("order_rep") and c["order_rep"] != "-- No usar --" else ""
    r["Clientes reportados en repuestos"] = r[c["customer_rep"]].map(first_str) if c.get("customer_rep") and c["customer_rep"] != "-- No usar --" else ""

    input_rows = len(r)
    excluded_blank = r[(~r["Serial equipo"].astype(bool)) | (~r["Referencia consumida"].astype(bool))].copy()
    r = r[(r["Serial equipo"].astype(bool)) & (r["Referencia consumida"].astype(bool))].copy()
    r["Cantidad consumida"] = r["Cantidad consumida"].fillna(1.0)

    base = (
        r.groupby(["Serial equipo", "Referencia consumida"], as_index=False)
        .agg(
            {
                "Producto consumido": first_non_empty,
                "Cantidad consumida": "sum",
                "Órdenes de servicio": unique_join,
                "Clientes reportados en repuestos": unique_join,
            }
        )
        .copy()
    )

    exact_prices, alt_prices, duplicated_prices = build_price_records(
        price_df=price_df,
        part_col=c["part_price"],
        price_col=c["unit_price"],
        product_col=c.get("product_price"),
        currency_col=c.get("currency_price"),
        alt_col=c.get("alt_price"),
        source_sheet=price_sheet,
    )

    price_rows = [lookup_price(row["Referencia consumida"], row["Producto consumido"], exact_prices, alt_prices) for _, row in base.iterrows()]
    price_match = pd.DataFrame(price_rows)
    detail = pd.concat([base.reset_index(drop=True), price_match.reset_index(drop=True)], axis=1)
    detail["Costo línea"] = detail["Cantidad consumida"] * detail["Precio unitario OPT2"]
    detail = detail.merge(age_map, on="Serial equipo", how="left")

    detail_cols = [
        "Serial equipo", "Fecha referencia instrumento", "Antigüedad años", "Modelo", "Distribuidor", "País", "Cliente según antigüedad",
        "Referencia consumida", "Referencia usada para precio", "Producto consumido", "Cantidad consumida",
        "Precio unitario OPT2", "Moneda", "Costo línea", "Estado precio", "Fuente precio",
        "Órdenes de servicio", "Clientes reportados en repuestos", "Estado rutina", "Fuente antigüedad",
    ]
    detail["Antigüedad años"] = detail["Fecha referencia instrumento"].map(calculate_age_years)
    for col in detail_cols:
        if col not in detail.columns:
            detail[col] = np.nan
    detail = detail[detail_cols].sort_values(["Serial equipo", "Referencia consumida"]).reset_index(drop=True)

    summary = (
        detail.groupby("Serial equipo", as_index=False)
        .agg(
            {
                "Fecha referencia instrumento": "first",
                "Antigüedad años": "first",
                "Modelo": first_non_empty,
                "Distribuidor": first_non_empty,
                "País": first_non_empty,
                "Cliente según antigüedad": first_non_empty,
                "Referencia consumida": "count",
                "Cantidad consumida": "sum",
                "Precio unitario OPT2": lambda s: int(s.isna().sum()),
                "Costo línea": "sum",
                "Estado rutina": first_non_empty,
                "Fuente antigüedad": first_non_empty,
            }
        )
        .rename(
            columns={
                "Cliente según antigüedad": "Cliente",
                "Referencia consumida": "Líneas de repuesto",
                "Cantidad consumida": "Cantidad total consumida",
                "Precio unitario OPT2": "Repuestos sin precio",
                "Costo línea": "Costo total repuestos EUR",
            }
        )
    )
    summary["Semáforo antigüedad"] = summary["Fecha referencia instrumento"].map(semaforo_from_date)
    summary["Nota fecha"] = np.where(
        summary["Fecha referencia instrumento"].isna(),
        "Serial sin fecha encontrada en los archivos de antigüedad/fabricación.",
        "Fecha tomada del archivo de antigüedad/fabricación como referencia para el instrumento.",
    )
    summary = summary.sort_values("Costo total repuestos EUR", ascending=False).reset_index(drop=True)

    missing_price = (
        detail[detail["Precio unitario OPT2"].isna()]
        .groupby(["Referencia consumida", "Producto consumido"], as_index=False)
        .agg({"Cantidad consumida": "sum", "Serial equipo": lambda s: ", ".join(sorted(set(s)))})
        .rename(columns={"Referencia consumida": "Referencia", "Producto consumido": "Producto", "Cantidad consumida": "Cantidad total", "Serial equipo": "Seriales afectados"})
        .sort_values("Referencia")
    )
    if len(missing_price):
        missing_price["Observación"] = "No se encontró precio OPT2 en el archivo de spare parts."

    missing_dates = summary[summary["Fecha referencia instrumento"].isna()][["Serial equipo"]].drop_duplicates().sort_values("Serial equipo").reset_index(drop=True)

    validations = pd.DataFrame(
        [
            ["Resumen", "Filas archivo repuestos", input_rows, "Filas leídas desde repuestos instalados."],
            ["Resumen", "Filas procesadas", len(r), "Filas con serial y referencia válida."],
            ["Resumen", "Filas excluidas", len(excluded_blank), "Filas sin serial o sin referencia."],
            ["Resumen", "Equipos procesados", summary["Serial equipo"].nunique(), "Seriales únicos procesados."],
            ["Resumen", "Líneas consolidadas", len(detail), "Serial + referencia."],
            ["Resumen", "Costo total EUR", float(detail["Costo línea"].sum(skipna=True)), "Cantidad consumida × precio OPT2."],
            ["Resumen", "Referencias sin precio", missing_price["Referencia"].nunique() if len(missing_price) else 0, "Referencias distintas sin precio."],
            ["Resumen", "Seriales sin fecha", missing_dates["Serial equipo"].nunique(), "Seriales no encontrados en antigüedad."],
            ["Configuración", "Hoja precios", price_sheet, "Hoja usada para Option 2 / OPT2."],
            ["Configuración", "Columna precio", c.get("unit_price"), "Columna usada como precio unitario."],
            ["Configuración", "Archivos antigüedad cargados", len(age_diag), "Archivos recibidos en la pestaña 2."],
        ],
        columns=["Tipo", "Validación", "Resultado", "Observación"],
    )

    return {
        "summary": summary,
        "detail": detail,
        "validations": validations,
        "missing_price": missing_price,
        "missing_dates": missing_dates,
        "duplicated_prices": duplicated_prices,
        "excluded_blank": excluded_blank,
        "age_map": age_map,
        "age_diag": age_diag,
        "total_cost": float(detail["Costo línea"].sum(skipna=True)),
        "total_equipment": int(summary["Serial equipo"].nunique()),
        "total_lines": int(len(detail)),
        "total_missing_price_refs": int(missing_price["Referencia"].nunique()) if len(missing_price) else 0,
        "total_missing_price_lines": int(detail["Precio unitario OPT2"].isna().sum()),
        "total_missing_dates": int(missing_dates["Serial equipo"].nunique()),
    }


# =============================================================================
# GRÁFICAS ESTILO INFORME EXCEL
# =============================================================================

def base_fig_layout(fig: go.Figure, title: str, height: int = 520) -> go.Figure:
    fig.update_layout(
        title={"text": title, "font": {"size": 18, "color": "#10324E", "family": "Segoe UI, Arial"}},
        height=height,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font={"color": "#172033", "family": "Segoe UI, Arial"},
        margin={"l": 28, "r": 28, "t": 62, "b": 36},
        legend={"orientation": "h", "y": -0.18},
    )
    fig.update_xaxes(gridcolor="#E7EEF7", zerolinecolor="#D8E2EF", linecolor="#D8E2EF", tickfont={"color": "#172033"})
    fig.update_yaxes(gridcolor="#E7EEF7", zerolinecolor="#D8E2EF", linecolor="#D8E2EF", tickfont={"color": "#172033"})
    return fig


def chart_all_instruments(summary: pd.DataFrame) -> go.Figure:
    df = summary.copy().sort_values("Costo total repuestos EUR", ascending=True)
    height = max(520, min(1300, 120 + len(df) * 24))
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=df["Costo total repuestos EUR"],
            y=df["Serial equipo"],
            orientation="h",
            marker={"color": "#00D9FF", "line": {"color": "#10324E", "width": 0.7}},
            text=[f"€{x:,.0f}" for x in df["Costo total repuestos EUR"].fillna(0)],
            textposition="outside",
            cliponaxis=False,
            hovertemplate="<b>Serial:</b> %{y}<br><b>Costo total:</b> €%{x:,.2f}<extra></extra>",
        )
    )
    fig = base_fig_layout(fig, "Costo total por instrumento - todos los seriales", height=height)
    fig.update_xaxes(title="Costo total EUR", tickprefix="€")
    fig.update_yaxes(title="Serial equipo", automargin=True)
    return fig


def chart_cost_by_model(summary: pd.DataFrame) -> go.Figure:
    df = summary.copy()
    df["Modelo"] = df["Modelo"].replace("", np.nan).fillna("Modelo no disponible")
    data = df.groupby("Modelo", as_index=False)["Costo total repuestos EUR"].sum().sort_values("Costo total repuestos EUR", ascending=False)
    fig = go.Figure(
        data=[
            go.Pie(
                labels=data["Modelo"],
                values=data["Costo total repuestos EUR"],
                hole=0.58,
                marker={"colors": ["#00D9FF", "#2CFF9B", "#FFD166", "#A855F7", "#FF4D8D", "#FF9F1C"], "line": {"color": "#FFFFFF", "width": 2}},
                textinfo="label+percent",
                hovertemplate="<b>%{label}</b><br>€%{value:,.2f}<br>%{percent}<extra></extra>",
            )
        ]
    )
    fig = base_fig_layout(fig, "Distribución del costo por modelo", height=410)
    return fig


def chart_age_vs_cost(summary: pd.DataFrame) -> go.Figure:
    df = summary.copy()
    df = df[df["Antigüedad años"].notna()].copy()
    fig = go.Figure()
    if len(df):
        max_cost = max(float(df["Costo total repuestos EUR"].max()), 1.0)
        fig.add_trace(
            go.Scatter(
                x=df["Antigüedad años"],
                y=df["Costo total repuestos EUR"],
                mode="markers+text",
                text=df["Serial equipo"],
                textposition="top center",
                marker={
                    "size": np.clip(df["Costo total repuestos EUR"].fillna(0) / max_cost * 34 + 9, 9, 42),
                    "color": "#10324E",
                    "line": {"color": "#00D9FF", "width": 2},
                    "opacity": 0.86,
                },
                hovertemplate="<b>%{text}</b><br>Antigüedad: %{x:.2f} años<br>Costo: €%{y:,.2f}<extra></extra>",
            )
        )
    fig = base_fig_layout(fig, "Fecha / antigüedad vs inversión en repuestos", height=440)
    fig.update_xaxes(title="Antigüedad aproximada años")
    fig.update_yaxes(title="Costo total EUR", tickprefix="€")
    return fig


def chart_refs_without_price(detail: pd.DataFrame) -> go.Figure:
    df = detail.copy()
    missing = int(df["Precio unitario OPT2"].isna().sum())
    ok = int(len(df) - missing)
    data = pd.DataFrame({"Estado": ["Con precio", "Sin precio"], "Líneas": [ok, missing]}) if missing else pd.DataFrame({"Estado": ["Con precio"], "Líneas": [ok]})
    fig = go.Figure(
        data=[go.Pie(labels=data["Estado"], values=data["Líneas"], hole=0.60, marker={"colors": ["#2CFF9B", "#FF4D8D"], "line": {"color": "#FFFFFF", "width": 2}})]
    )
    fig = base_fig_layout(fig, "Cobertura de precios OPT2", height=410)
    return fig


# =============================================================================
# EXCEL FINAL
# =============================================================================

def write_df_table(writer: pd.ExcelWriter, df: pd.DataFrame, sheet: str, table_name: str, widths: Optional[Dict[str, int]] = None, tab_color: str = NEON_CYAN) -> None:
    widths = widths or {}
    safe_name = table_name[:31].replace(" ", "_").replace("-", "_")
    df.to_excel(writer, sheet_name=sheet, index=False)
    wb = writer.book
    ws = writer.sheets[sheet]
    ws.set_tab_color(tab_color)
    ws.hide_gridlines(2)
    header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#10324E", "border": 1, "align": "center", "valign": "vcenter"})
    text_fmt = wb.add_format({"border": 1, "valign": "top"})
    wrap_fmt = wb.add_format({"border": 1, "valign": "top", "text_wrap": True})
    date_fmt = wb.add_format({"num_format": "yyyy-mm-dd", "border": 1, "valign": "top"})
    money_fmt = wb.add_format({"num_format": "€#,##0.00", "border": 1, "valign": "top"})
    qty_fmt = wb.add_format({"num_format": "#,##0.00", "border": 1, "valign": "top"})

    rows, cols = df.shape
    for j, col in enumerate(df.columns):
        ws.write(0, j, col, header_fmt)
        norm = normalize_text(col)
        width = widths.get(col, min(max(len(str(col)) + 4, 12), 42))
        fmt = text_fmt
        if "fecha" in norm or "date" in norm:
            fmt = date_fmt
            width = max(width, 18)
        elif "precio" in norm or "costo" in norm or "eur" in norm:
            fmt = money_fmt
            width = max(width, 17)
        elif "cantidad" in norm or "lineas" in norm or "antiguedad" in norm:
            fmt = qty_fmt
            width = max(width, 15)
        elif any(x in norm for x in ["producto", "cliente", "nota", "observacion", "fuente", "orden"]):
            fmt = wrap_fmt
            width = max(width, 26)
        ws.set_column(j, j, min(width, 60), fmt)

    if rows > 0 and cols > 0:
        ws.add_table(0, 0, rows, cols - 1, {"name": safe_name, "style": "Table Style Medium 2", "columns": [{"header": str(c)} for c in df.columns]})
    ws.freeze_panes(1, 0)


def build_excel(result: Dict[str, object]) -> bytes:
    summary: pd.DataFrame = result["summary"]  # type: ignore
    detail: pd.DataFrame = result["detail"]  # type: ignore
    validations: pd.DataFrame = result["validations"]  # type: ignore
    missing_price: pd.DataFrame = result["missing_price"]  # type: ignore
    missing_dates: pd.DataFrame = result["missing_dates"]  # type: ignore
    duplicated_prices: pd.DataFrame = result["duplicated_prices"]  # type: ignore
    age_map: pd.DataFrame = result["age_map"]  # type: ignore
    age_diag: pd.DataFrame = result["age_diag"]  # type: ignore
    excluded_blank: pd.DataFrame = result["excluded_blank"]  # type: ignore

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
        wb = writer.book
        ws = wb.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = ws
        ws.set_tab_color("#00D9FF")
        ws.hide_gridlines(2)
        ws.set_zoom(86)
        ws.set_column("A:A", 18)
        ws.set_column("B:B", 20)
        ws.set_column("C:C", 16)
        ws.set_column("D:D", 18)
        ws.set_column("E:E", 18)
        ws.set_column("F:F", 18)
        ws.set_column("G:G", 20)
        ws.set_column("H:H", 20)
        ws.set_column("I:I", 22)
        ws.set_column("J:J", 34)

        title_fmt = wb.add_format({"bold": True, "font_size": 20, "font_color": "white", "bg_color": "#07111F", "align": "center", "valign": "vcenter"})
        subtitle_fmt = wb.add_format({"font_size": 10, "font_color": "#A8B3C7", "bg_color": "#07111F", "align": "center"})
        kpi_head = wb.add_format({"bold": True, "font_color": "#00111D", "bg_color": "#00D9FF", "align": "center", "border": 1})
        kpi_num = wb.add_format({"bold": True, "font_size": 15, "font_color": "white", "bg_color": "#10324E", "align": "center", "border": 1})
        kpi_money = wb.add_format({"bold": True, "font_size": 15, "font_color": "white", "bg_color": "#10324E", "num_format": "€#,##0.00", "align": "center", "border": 1})
        header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#10324E", "border": 1, "align": "center"})
        money_fmt = wb.add_format({"num_format": "€#,##0.00", "border": 1})
        text_fmt = wb.add_format({"border": 1})
        date_fmt = wb.add_format({"num_format": "yyyy-mm-dd", "border": 1})

        ws.merge_range("A1:J1", "Reporte de repuestos consumidos por equipo", title_fmt)
        ws.merge_range("A2:J2", f"Generado por APP v{APP_VERSION} · Dashboard con TODOS los instrumentos", subtitle_fmt)
        ws.set_row(0, 30)

        kpis = [
            ("Equipos", result["total_equipment"], kpi_num),
            ("Líneas", result["total_lines"], kpi_num),
            ("Costo total EUR", result["total_cost"], kpi_money),
            ("Refs sin precio", result["total_missing_price_refs"], kpi_num),
            ("Seriales sin fecha", result["total_missing_dates"], kpi_num),
        ]
        for idx, (label, value, fmt) in enumerate(kpis):
            col = idx * 2
            ws.write(3, col, label, kpi_head)
            ws.write(4, col, value, fmt)

        dashboard_cols = [
            "Serial equipo", "Fecha referencia instrumento", "Antigüedad años", "Modelo", "Distribuidor", "País",
            "Líneas de repuesto", "Cantidad total consumida", "Repuestos sin precio", "Costo total repuestos EUR",
        ]
        dash_data = summary[dashboard_cols].copy()
        start_row = 7
        ws.write(start_row - 1, 0, "Todos los instrumentos", header_fmt)
        for j, col in enumerate(dash_data.columns):
            ws.write(start_row, j, col, header_fmt)
        for i, row in enumerate(dash_data.itertuples(index=False), start=start_row + 1):
            for j, value in enumerate(row):
                if j == 1:
                    ws.write_datetime(i, j, pd.to_datetime(value).to_pydatetime(), date_fmt) if pd.notna(value) else ws.write(i, j, "", text_fmt)
                elif j in [2, 6, 7, 8]:
                    ws.write_number(i, j, float(value) if pd.notna(value) else 0, text_fmt)
                elif j == 9:
                    ws.write_number(i, j, float(value) if pd.notna(value) else 0, money_fmt)
                else:
                    ws.write(i, j, value, text_fmt)

        if len(dash_data):
            rows = len(dash_data)
            ws.add_table(start_row, 0, start_row + rows, len(dash_data.columns) - 1, {
                "name": "DashboardTodosInstrumentos",
                "style": "Table Style Medium 2",
                "columns": [{"header": str(c)} for c in dash_data.columns],
            })
            chart = wb.add_chart({"type": "bar"})
            chart.add_series({
                "name": "Costo total EUR",
                "categories": ["Dashboard", start_row + 1, 0, start_row + rows, 0],
                "values": ["Dashboard", start_row + 1, 9, start_row + rows, 9],
                "data_labels": {"value": True, "num_format": "€#,##0"},
                "fill": {"color": "#00D9FF"},
                "border": {"color": "#2CFF9B"},
            })
            chart.set_title({"name": "Costo total por instrumento - todos los seriales"})
            chart.set_x_axis({"num_format": "€#,##0"})
            chart.set_y_axis({"reverse": True})
            chart.set_legend({"none": True})
            chart.set_size({"width": 980, "height": max(420, min(920, 160 + len(dash_data) * 14))})
            ws.insert_chart("L4", chart)

        write_df_table(writer, summary, "Resumen por equipo", "ResumenEquipo", {
            "Serial equipo": 17, "Fecha referencia instrumento": 22, "Antigüedad años": 16, "Modelo": 18, "Distribuidor": 26,
            "País": 14, "Cliente": 38, "Líneas de repuesto": 18, "Cantidad total consumida": 22,
            "Repuestos sin precio": 20, "Costo total repuestos EUR": 24, "Estado rutina": 18,
            "Fuente antigüedad": 40, "Semáforo antigüedad": 18, "Nota fecha": 58,
        }, tab_color="#2CFF9B")
        write_df_table(writer, detail, "Detalle repuestos", "DetalleRepuestos", {
            "Serial equipo": 17, "Fecha referencia instrumento": 22, "Antigüedad años": 16, "Modelo": 18,
            "Distribuidor": 26, "País": 14, "Cliente según antigüedad": 38, "Referencia consumida": 20,
            "Referencia usada para precio": 23, "Producto consumido": 44, "Cantidad consumida": 18,
            "Precio unitario OPT2": 20, "Moneda": 10, "Costo línea": 17, "Estado precio": 38,
            "Fuente precio": 30, "Órdenes de servicio": 30, "Clientes reportados en repuestos": 42,
            "Estado rutina": 18, "Fuente antigüedad": 42,
        }, tab_color="#FFD166")
        write_df_table(writer, age_map, "Antigüedad consolidada", "AntiguedadConsolidada", tab_color="#A855F7")
        write_df_table(writer, age_diag, "Diagnóstico archivos", "DiagnosticoArchivos", tab_color="#FF9F1C")

        # Validaciones
        val_sheet = "Validaciones"
        validations.to_excel(writer, sheet_name=val_sheet, index=False, startrow=0)
        ws_val = writer.sheets[val_sheet]
        ws_val.set_tab_color("#FF4D8D")
        ws_val.hide_gridlines(2)
        ws_val.set_column("A:A", 18)
        ws_val.set_column("B:B", 34)
        ws_val.set_column("C:C", 24)
        ws_val.set_column("D:D", 70)
        for j, col in enumerate(validations.columns):
            ws_val.write(0, j, col, header_fmt)
        cursor = len(validations) + 3
        ws_val.write(cursor, 0, "Referencias sin precio OPT2", header_fmt)
        cursor += 1
        if len(missing_price):
            missing_price.to_excel(writer, sheet_name=val_sheet, index=False, startrow=cursor)
            for j, col in enumerate(missing_price.columns):
                ws_val.write(cursor, j, col, header_fmt)
            cursor += len(missing_price) + 3
        else:
            ws_val.write(cursor, 0, "Sin referencias pendientes de precio.", text_fmt)
            cursor += 3
        ws_val.write(cursor, 0, "Seriales sin fecha", header_fmt)
        cursor += 1
        if len(missing_dates):
            missing_dates.to_excel(writer, sheet_name=val_sheet, index=False, startrow=cursor)
            for j, col in enumerate(missing_dates.columns):
                ws_val.write(cursor, j, col, header_fmt)
            cursor += len(missing_dates) + 3
        else:
            ws_val.write(cursor, 0, "Sin seriales pendientes de fecha.", text_fmt)
            cursor += 3
        ws_val.write(cursor, 0, "Part numbers duplicados en lista de precios", header_fmt)
        cursor += 1
        if len(duplicated_prices):
            duplicated_prices.to_excel(writer, sheet_name=val_sheet, index=False, startrow=cursor)
        else:
            ws_val.write(cursor, 0, "Sin duplicados relevantes.", text_fmt)

        if len(excluded_blank):
            excluded_blank.head(5000).to_excel(writer, sheet_name="Filas excluidas", index=False)
            writer.sheets["Filas excluidas"].set_tab_color("#FF4D8D")

        # Formato condicional
        ws_sum = writer.sheets["Resumen por equipo"]
        if len(summary):
            cost_col = summary.columns.get_loc("Costo total repuestos EUR")
            ws_sum.conditional_format(1, cost_col, len(summary), cost_col, {"type": "data_bar", "bar_color": "#00D9FF"})
            missing_col = summary.columns.get_loc("Repuestos sin precio")
            ws_sum.conditional_format(1, missing_col, len(summary), missing_col, {"type": "cell", "criteria": ">", "value": 0, "format": wb.add_format({"bg_color": "#FFD6E5", "font_color": "#B0003A"})})
        ws_det = writer.sheets["Detalle repuestos"]
        if len(detail):
            state_col = detail.columns.get_loc("Estado precio")
            ws_det.conditional_format(1, state_col, len(detail), state_col, {"type": "text", "criteria": "containing", "value": "no encontrado", "format": wb.add_format({"bg_color": "#FFD6E5", "font_color": "#B0003A"})})

    output.seek(0)
    return output.getvalue()



# =============================================================================
# UI STREAMLIT - V14 DASHBOARD TIPO EXCEL, SIN HTML ROTO, SIN CEROS FANTASMA
# =============================================================================

import streamlit.components.v1 as components


def inject_css() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background:
                radial-gradient(circle at 12% 8%, rgba(0,217,255,0.08), transparent 22%),
                radial-gradient(circle at 88% 2%, rgba(168,85,247,0.08), transparent 24%),
                linear-gradient(180deg, #050910 0%, #08101B 100%);
            color: #EAF1FB;
            font-family: "Segoe UI", Arial, sans-serif;
        }
        [data-testid="stHeader"] { background: rgba(0,0,0,0); }
        .block-container {
            max-width: 1680px;
            padding: 0.8rem 1.1rem 1.2rem 1.1rem;
        }
        .main-title-card {
            background: linear-gradient(135deg, rgba(9,18,33,0.96), rgba(13,27,46,0.92));
            border: 1px solid rgba(0,217,255,0.20);
            border-radius: 22px;
            padding: 18px 22px;
            box-shadow: 0 18px 34px rgba(0,0,0,0.28);
            margin-bottom: 12px;
        }
        .main-title-card h1 {
            margin: 0;
            color: #FFFFFF;
            font-size: 30px;
            font-weight: 950;
            letter-spacing: .2px;
        }
        .main-title-card p {
            margin: 7px 0 0 0;
            color: #A8B3C7;
            font-size: 13px;
            line-height: 1.35;
        }
        .pill-row { display:flex; flex-wrap:wrap; gap:8px; margin-top:11px; }
        .pill {
            border: 1px solid rgba(255,255,255,0.10);
            background: rgba(255,255,255,0.055);
            color: #F5F9FF;
            border-radius: 999px;
            padding: 5px 10px;
            font-size: 11px;
            font-weight: 800;
        }
        .section-title {
            color: #FFFFFF;
            font-size: 18px;
            font-weight: 950;
            margin: 10px 0 2px 0;
        }
        .section-subtitle {
            color: #91A2BA;
            font-size: 12px;
            margin: 0 0 8px 0;
        }
        .hint-card {
            background: rgba(0,217,255,0.08);
            border: 1px solid rgba(0,217,255,0.18);
            border-radius: 14px;
            padding: 10px 12px;
            color: #CFEAFF;
            font-size: 12px;
            line-height: 1.35;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background: rgba(13,20,34,0.78);
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 18px;
            box-shadow: 0 12px 26px rgba(0,0,0,0.18);
        }
        div[data-testid="stFileUploader"] section {
            background: rgba(255,255,255,0.035) !important;
            border: 1px dashed rgba(255,255,255,0.20) !important;
            border-radius: 14px !important;
        }
        div[data-testid="stFileUploader"] label,
        div[data-testid="stFileUploader"] small,
        div[data-testid="stFileUploader"] p,
        div[data-testid="stFileUploader"] span {
            color: #DCE7F8 !important;
        }
        .stButton > button, .stDownloadButton > button {
            border-radius: 14px !important;
            border: 1px solid rgba(0,217,255,0.28) !important;
            background: linear-gradient(90deg, #073B5A, #0B4B76) !important;
            color: #FFFFFF !important;
            font-weight: 950 !important;
            min-height: 44px;
            box-shadow: 0 10px 20px rgba(0,0,0,0.22) !important;
        }
        .stButton > button:hover, .stDownloadButton > button:hover {
            border-color: rgba(44,255,155,0.44) !important;
            color: #FFFFFF !important;
        }
        .metric-card {
            background: linear-gradient(180deg, rgba(12,20,34,0.96), rgba(17,28,46,0.96));
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 18px;
            padding: 13px 14px;
            min-height: 104px;
            box-shadow: 0 12px 24px rgba(0,0,0,0.22);
        }
        .metric-line { width:46px; height:4px; border-radius:999px; margin-bottom:9px; }
        .metric-label { color:#9FB0C8; font-size:12px; font-weight:800; }
        .metric-value { color:#FFFFFF; font-size:25px; font-weight:950; line-height:1.08; margin-top:4px; }
        .metric-foot { color:#7F90A8; font-size:11px; margin-top:5px; }
        .grid-title {
            color: #FFFFFF;
            font-size: 15px;
            font-weight: 950;
            margin-bottom: 6px;
        }
        .grid-sub {
            color:#90A0B7;
            font-size:11px;
            margin-bottom:8px;
        }
        .stDataFrame { border-radius: 14px; overflow: hidden; }
        .stExpander {
            background: rgba(13,20,34,0.78);
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 14px;
        }
        div[data-baseweb="select"] > div { border-radius: 12px !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def money(value: object) -> str:
    try:
        if value is None or pd.isna(value):
            return "€0.00"
        return f"€{float(value):,.2f}"
    except Exception:
        return "€0.00"


def money0(value: object) -> str:
    try:
        if value is None or pd.isna(value):
            return "€0"
        return f"€{float(value):,.0f}"
    except Exception:
        return "€0"


def num(value: object) -> str:
    try:
        if value is None or pd.isna(value):
            return "0"
        return f"{int(float(value)):,}"
    except Exception:
        return "0"


def hero() -> None:
    st.markdown(
        """
        <div class="main-title-card">
            <h1>Costo de repuestos por equipo</h1>
            <p>Dashboard tipo Excel: KPIs, gráfico de todos los instrumentos, tabla completa y validaciones. El dashboard y el Excel se generan en el mismo ciclo para evitar renderizados intermedios.</p>
            <div class="pill-row">
                <span class="pill">Versión 14.0.0</span>
                <span class="pill">Visual tipo Excel</span>
                <span class="pill">Sin render intermedio</span>
                <span class="pill">Multiarchivo antigüedad</span>
                <span class="pill">Excel final intacto</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section(title: str, subtitle: str = "") -> None:
    st.markdown(f'<div class="section-title">{title}</div><div class="section-subtitle">{subtitle}</div>', unsafe_allow_html=True)


def metric_card(label: str, value: str, foot: str, color: str) -> None:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-line" style="background:{color}; box-shadow:0 0 16px {color};"></div>
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-foot">{foot}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def file_signature(rep_file, age_files, price_file) -> str:
    parts = []
    if rep_file is not None:
        parts.append(f"rep:{rep_file.name}:{rep_file.size}")
    for f in age_files or []:
        parts.append(f"age:{f.name}:{f.size}")
    if price_file is not None:
        parts.append(f"price:{price_file.name}:{price_file.size}")
    return "|".join(parts)


def render_uploads():
    section("Carga de archivos", "Carga los tres insumos. Antigüedad/fabricación permite varios archivos.")
    c1, c2, c3, c4 = st.columns([1, 1, 1, 0.55], gap="small")
    with c1:
        with st.container(border=True):
            st.markdown("**1 · Repuestos instalados**")
            st.caption("CSV/XLSX con serial, referencia y cantidad.")
            rep_file = st.file_uploader("Repuestos instalados", type=["csv", "txt", "xlsx", "xlsm", "xls"], key="rep_v14", label_visibility="collapsed")
    with c2:
        with st.container(border=True):
            st.markdown("**2 · Antigüedad / fabricación**")
            st.caption("Carga múltiple para diferentes modelos.")
            age_files = st.file_uploader("Antigüedad", type=["csv", "txt", "xlsx", "xlsm", "xls"], accept_multiple_files=True, key="age_v14", label_visibility="collapsed")
    with c3:
        with st.container(border=True):
            st.markdown("**3 · Spare parts Option 2**")
            st.caption("Lista de precios para costo unitario.")
            price_file = st.file_uploader("Spare parts", type=["csv", "txt", "xlsx", "xlsm", "xls"], key="price_v14", label_visibility="collapsed")
    with c4:
        with st.container(border=True):
            st.markdown("**Estado**")
            st.caption("Control de carga")
            loaded = int(rep_file is not None) + int(bool(age_files)) + int(price_file is not None)
            st.markdown(f"### {loaded}/3")
            if loaded == 3:
                st.success("Listo para procesar")
            else:
                st.info("Faltan archivos")
    return rep_file, age_files, price_file


def build_config(rep_file, age_files, price_file):
    if not (rep_file and age_files and price_file):
        return None

    try:
        rep_sheets = read_uploaded_file(rep_file)
        price_sheets = read_uploaded_file(price_file)
        age_map, age_diag = consolidate_age_files(age_files)
    except Exception as exc:
        st.error(f"No fue posible leer los archivos: {exc}")
        return None

    rep_default = choose_default_sheet(rep_sheets, "repuestos") or list(rep_sheets.keys())[0]
    price_default = choose_default_sheet(price_sheets, "precios") or list(price_sheets.keys())[0]

    # Configuración base automática. Si el usuario no abre el expander, se usa esto.
    rep_sheet = rep_default
    price_sheet = price_default
    rep_df = rep_sheets[rep_sheet]
    price_df = price_sheets[price_sheet]
    prefer_opt2 = any(x in normalize_text(price_sheet) for x in ["option 2", "opcion 2", "opt2", "opt 2", "total update"])

    detected = {
        "serial_rep": find_best_column(rep_df, SERIAL_ALIASES),
        "part_rep": find_best_column(rep_df, PART_ALIASES),
        "qty_rep": find_best_column(rep_df, QTY_ALIASES),
        "product_rep": find_best_column(rep_df, PRODUCT_ALIASES),
        "order_rep": find_best_column(rep_df, ORDER_ALIASES),
        "customer_rep": find_best_column(rep_df, CUSTOMER_ALIASES),
        "part_price": find_best_column(price_df, PART_ALIASES),
        "unit_price": find_best_price_column(price_df, prefer_opt2),
        "product_price": find_best_column(price_df, PRODUCT_ALIASES),
        "currency_price": find_best_column(price_df, CURRENCY_ALIASES),
        "alt_price": find_best_column(price_df, ALT_CODE_ALIASES),
    }
    cols = {
        "serial_rep": detected["serial_rep"],
        "part_rep": detected["part_rep"],
        "qty_rep": detected["qty_rep"] or "-- No usar --",
        "product_rep": detected["product_rep"] or "-- No usar --",
        "order_rep": detected["order_rep"] or "-- No usar --",
        "customer_rep": detected["customer_rep"] or "-- No usar --",
        "part_price": detected["part_price"],
        "unit_price": detected["unit_price"],
        "product_price": detected["product_price"] or "-- No usar --",
        "currency_price": detected["currency_price"] or "-- No usar --",
        "alt_price": detected["alt_price"] or "-- No usar --",
    }

    with st.expander("Configuración avanzada (opcional)", expanded=False):
        st.caption("Abre esto solo si el resultado no toma las columnas correctas. La detección automática queda aplicada por defecto.")
        h1, h2 = st.columns(2, gap="small")
        with h1:
            rep_sheet = st.selectbox("Hoja de repuestos", list(rep_sheets.keys()), index=list(rep_sheets.keys()).index(rep_default), key="rep_sheet_v14")
        with h2:
            price_sheet = st.selectbox("Hoja de precios", list(price_sheets.keys()), index=list(price_sheets.keys()).index(price_default), key="price_sheet_v14")
        rep_df = rep_sheets[rep_sheet]
        price_df = price_sheets[price_sheet]
        prefer_opt2 = any(x in normalize_text(price_sheet) for x in ["option 2", "opcion 2", "opt2", "opt 2", "total update"])
        detected = {
            "serial_rep": find_best_column(rep_df, SERIAL_ALIASES),
            "part_rep": find_best_column(rep_df, PART_ALIASES),
            "qty_rep": find_best_column(rep_df, QTY_ALIASES),
            "product_rep": find_best_column(rep_df, PRODUCT_ALIASES),
            "order_rep": find_best_column(rep_df, ORDER_ALIASES),
            "customer_rep": find_best_column(rep_df, CUSTOMER_ALIASES),
            "part_price": find_best_column(price_df, PART_ALIASES),
            "unit_price": find_best_price_column(price_df, prefer_opt2),
            "product_price": find_best_column(price_df, PRODUCT_ALIASES),
            "currency_price": find_best_column(price_df, CURRENCY_ALIASES),
            "alt_price": find_best_column(price_df, ALT_CODE_ALIASES),
        }
        a, b, c, d = st.columns(4, gap="small")
        with a:
            cols["serial_rep"] = st.selectbox("Serial en repuestos", select_options(rep_df, False), index=index_for(select_options(rep_df, False), detected["serial_rep"]), key="serial_rep_v14")
            cols["qty_rep"] = st.selectbox("Cantidad", select_options(rep_df, True), index=index_for(select_options(rep_df, True), detected["qty_rep"]), key="qty_rep_v14")
        with b:
            cols["part_rep"] = st.selectbox("Referencia en repuestos", select_options(rep_df, False), index=index_for(select_options(rep_df, False), detected["part_rep"]), key="part_rep_v14")
            cols["product_rep"] = st.selectbox("Producto consumido", select_options(rep_df, True), index=index_for(select_options(rep_df, True), detected["product_rep"]), key="prod_rep_v14")
        with c:
            cols["part_price"] = st.selectbox("Referencia en precios", select_options(price_df, False), index=index_for(select_options(price_df, False), detected["part_price"]), key="part_price_v14")
            cols["unit_price"] = st.selectbox("Precio unitario OPT2", select_options(price_df, False), index=index_for(select_options(price_df, False), detected["unit_price"]), key="unit_price_v14")
        with d:
            cols["product_price"] = st.selectbox("Descripción en precios", select_options(price_df, True), index=index_for(select_options(price_df, True), detected["product_price"]), key="prod_price_v14")
            cols["currency_price"] = st.selectbox("Moneda", select_options(price_df, True), index=index_for(select_options(price_df, True), detected["currency_price"]), key="currency_v14")
            cols["alt_price"] = st.selectbox("Códigos alternativos", select_options(price_df, True), index=index_for(select_options(price_df, True), detected["alt_price"]), key="alt_price_v14")
            cols["order_rep"] = st.selectbox("Orden de servicio", select_options(rep_df, True), index=index_for(select_options(rep_df, True), detected["order_rep"]), key="order_v14")
            cols["customer_rep"] = st.selectbox("Cliente en repuestos", select_options(rep_df, True), index=index_for(select_options(rep_df, True), detected["customer_rep"]), key="customer_v14")
        st.markdown("**Diagnóstico de archivos de antigüedad**")
        st.dataframe(age_diag, use_container_width=True, hide_index=True, height=min(240, 80 + 35 * max(1, len(age_diag))))

    return {
        "rep_df": rep_df,
        "price_df": price_df,
        "price_sheet": price_sheet,
        "age_map": age_map,
        "age_diag": age_diag,
        "columns": cols,
    }


def excel_like_chart_html(summary: pd.DataFrame) -> str:
    df = summary.copy().sort_values("Costo total repuestos EUR", ascending=False)
    values = pd.to_numeric(df["Costo total repuestos EUR"], errors="coerce").fillna(0)
    max_val = float(values.max()) if len(values) else 0

    css = """
    <style>
    body { margin:0; background:transparent; font-family: Segoe UI, Arial, sans-serif; color:#EAF1FB; }
    .chart-card { background:rgba(13,20,34,0.96); border:1px solid rgba(255,255,255,.08); border-radius:18px; padding:14px 16px; }
    .chart-title { font-size:16px; font-weight:900; color:#FFFFFF; margin-bottom:3px; }
    .chart-sub { font-size:12px; color:#90A0B7; margin-bottom:12px; }
    .row { display:grid; grid-template-columns: 150px 1fr 118px; gap:10px; align-items:center; margin:7px 0; }
    .label { font-size:12px; color:#DCE7F8; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .track { height:20px; border-radius:999px; background:rgba(255,255,255,.055); border:1px solid rgba(255,255,255,.05); overflow:hidden; }
    .fill { height:100%; border-radius:999px; min-width:2px; }
    .value { font-size:12px; color:#FFFFFF; font-weight:800; text-align:right; }
    .empty { border:1px solid rgba(255,209,102,.30); background:rgba(255,209,102,.08); color:#FFE9A6; border-radius:14px; padding:14px; font-size:13px; }
    </style>
    """
    if max_val <= 0:
        return css + """
        <div class='chart-card'>
          <div class='chart-title'>Costo total por instrumento</div>
          <div class='empty'>No hay costos mayores a cero para graficar. Revisa que la columna <b>Precio unitario OPT2</b> esté seleccionada correctamente en Configuración avanzada.</div>
        </div>
        """
    colors = ["#00D9FF", "#2CFF9B", "#FFD166", "#A855F7", "#FF4D8D", "#FF9F1C"]
    rows = []
    for idx, (_, row) in enumerate(df.iterrows()):
        value = float(row.get("Costo total repuestos EUR", 0) or 0)
        if value <= 0:
            continue
        pct = min(100, max(2, value / max_val * 100))
        color = colors[idx % len(colors)]
        serial = str(row.get("Serial equipo", ""))
        rows.append(
            f"""
            <div class='row'>
              <div class='label' title='{serial}'>{serial}</div>
              <div class='track'><div class='fill' style='width:{pct:.2f}%; background:linear-gradient(90deg,{color},rgba(255,255,255,.78)); box-shadow:0 0 14px {color};'></div></div>
              <div class='value'>{money0(value)}</div>
            </div>
            """
        )
    return css + "<div class='chart-card'><div class='chart-title'>Costo total por instrumento · todos los seriales</div><div class='chart-sub'>Vista equivalente al gráfico horizontal del Dashboard de Excel.</div>" + "".join(rows) + "</div>"


def render_dashboard(result: Dict[str, object], excel_bytes: bytes) -> None:
    summary: pd.DataFrame = result["summary"]  # type: ignore
    validations: pd.DataFrame = result["validations"]  # type: ignore
    missing_price: pd.DataFrame = result["missing_price"]  # type: ignore
    missing_dates: pd.DataFrame = result["missing_dates"]  # type: ignore

    section("Dashboard", "Misma salida del Excel final organizada en pantalla: KPIs, gráfico horizontal, tabla completa y validaciones.")
    top1, top2 = st.columns([1.35, 0.65], gap="small")
    with top1:
        st.markdown("<div class='hint-card'>El gráfico principal replica el Dashboard del Excel: barras horizontales por serial, ordenadas por costo total. No se dibujan valores cero.</div>", unsafe_allow_html=True)
    with top2:
        st.download_button(
            "⬇️ Descargar informe Excel final",
            data=excel_bytes,
            file_name=f"reporte_repuestos_consumidos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="download_top_v14",
        )

    m1, m2, m3, m4, m5 = st.columns(5, gap="small")
    with m1:
        metric_card("Equipos", num(result["total_equipment"]), "Seriales procesados", "#00D9FF")
    with m2:
        metric_card("Líneas", num(result["total_lines"]), "Serial + referencia", "#2CFF9B")
    with m3:
        metric_card("Costo total", money(result["total_cost"]), "Cantidad × precio OPT2", "#FFD166")
    with m4:
        metric_card("Refs sin precio", num(result["total_missing_price_refs"]), "Pendientes de precio", "#FF4D8D")
    with m5:
        metric_card("Seriales sin fecha", num(result["total_missing_dates"]), "Pendientes antigüedad", "#FF9F1C")

    left, right = st.columns([0.58, 0.42], gap="small")
    with left:
        chart_height = max(460, min(980, 120 + 32 * max(6, len(summary))))
        components.html(excel_like_chart_html(summary), height=chart_height, scrolling=True)
    with right:
        with st.container(border=True):
            st.markdown("<div class='grid-title'>Todos los instrumentos</div><div class='grid-sub'>Tabla compacta del resumen por equipo.</div>", unsafe_allow_html=True)
            cols_to_show = [
                "Serial equipo", "Fecha referencia instrumento", "Antigüedad años", "Modelo", "Distribuidor", "País",
                "Líneas de repuesto", "Cantidad total consumida", "Repuestos sin precio", "Costo total repuestos EUR",
                "Semáforo antigüedad",
            ]
            existing = [c for c in cols_to_show if c in summary.columns]
            st.dataframe(
                summary[existing],
                use_container_width=True,
                hide_index=True,
                height=chart_height - 60,
                column_config={
                    "Fecha referencia instrumento": st.column_config.DateColumn("Fecha referencia"),
                    "Antigüedad años": st.column_config.NumberColumn("Antigüedad", format="%.2f"),
                    "Costo total repuestos EUR": st.column_config.NumberColumn("Costo EUR", format="€ %.2f"),
                    "Cantidad total consumida": st.column_config.NumberColumn("Cantidad", format="%.2f"),
                },
            )

    v1, v2 = st.columns([0.55, 0.45], gap="small")
    with v1:
        with st.container(border=True):
            st.markdown("<div class='grid-title'>Validaciones</div><div class='grid-sub'>Controles del procesamiento.</div>", unsafe_allow_html=True)
            st.dataframe(validations, use_container_width=True, hide_index=True, height=min(360, 120 + 35 * max(4, len(validations))))
    with v2:
        with st.container(border=True):
            st.markdown("<div class='grid-title'>Pendientes</div><div class='grid-sub'>Referencias sin precio y seriales sin fecha.</div>", unsafe_allow_html=True)
            st.markdown("**Referencias sin precio OPT2**")
            st.dataframe(missing_price, use_container_width=True, hide_index=True, height=150)
            st.markdown("**Seriales sin fecha**")
            st.dataframe(missing_dates, use_container_width=True, hide_index=True, height=150)


def main() -> None:
    inject_css()
    hero()

    rep_file, age_files, price_file = render_uploads()
    current_sig = file_signature(rep_file, age_files, price_file)
    if st.session_state.get("v14_signature") and st.session_state.get("v14_signature") != current_sig:
        st.session_state.pop("v14_result", None)
        st.session_state.pop("v14_excel", None)
        st.session_state.pop("v14_signature", None)
        st.session_state.pop("v14_ready", None)

    cfg = build_config(rep_file, age_files, price_file)
    if cfg is None:
        st.info("Carga los tres archivos para generar el dashboard.")
        return

    b1, b2 = st.columns([0.38, 1.62], gap="small")
    with b1:
        run = st.button("⚡ GENERAR DASHBOARD", use_container_width=True, key="run_v14")
    with b2:
        st.markdown("<div class='hint-card'>Al generar, la app calcula primero el Excel y el dashboard, luego recarga la vista limpia. Así no se muestran valores intermedios ni ceros antes del resultado final.</div>", unsafe_allow_html=True)

    if run:
        with st.spinner("Procesando archivos y generando informe..."):
            try:
                result = process_data(
                    rep_df=cfg["rep_df"],
                    age_map=cfg["age_map"],
                    price_df=cfg["price_df"],
                    price_sheet=cfg["price_sheet"],
                    c=cfg["columns"],
                    age_diag=cfg["age_diag"],
                )
                excel_bytes = build_excel(result)
                st.session_state["v14_result"] = result
                st.session_state["v14_excel"] = excel_bytes
                st.session_state["v14_signature"] = current_sig
                st.session_state["v14_ready"] = True

                # Reparación clave:
                # No renderizar el dashboard en el mismo ciclo del clic.
                # Streamlit vuelve a ejecutar la app y muestra el resultado ya consolidado,
                # evitando que aparezcan valores/etiquetas intermedias como los "0" verdes.
                try:
                    st.rerun()
                except AttributeError:
                    st.experimental_rerun()
            except Exception as exc:
                st.error(f"No fue posible procesar los archivos: {exc}")
                return

    if st.session_state.get("v14_ready") and st.session_state.get("v14_signature") == current_sig and "v14_result" in st.session_state and "v14_excel" in st.session_state:
        render_dashboard(st.session_state["v14_result"], st.session_state["v14_excel"])


if __name__ == "__main__":
    main()
