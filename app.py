"""
Proc-Bulls — Ferramenta de PROCV Inteligente
Desenvolvido por João · Aure Digital
"""

import io
import re
import traceback
from datetime import datetime, timedelta
from typing import Optional, List

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Proc-Bulls",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown(
    "<style>"
    "@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');"
    "html,body,.stApp,.stMarkdown,p,h1,h2,h3,h4,h5,h6,li,label,button,input,textarea,select,.stButton,.stTextInput,.stSelectbox,.stMultiSelect,.stMetric{font-family:'Inter',-apple-system,BlinkMacSystemFont,'SF Pro Display',sans-serif!important;-webkit-font-smoothing:antialiased!important}"
    ".stApp{background:#000!important}"
    ".main .block-container{padding-top:1.5rem!important;padding-bottom:6rem!important;max-width:1100px!important}"
    ".hero-wrap{text-align:center;padding:2rem 0 0.5rem}"
    ".hero-badge{display:inline-block;font-size:.68rem;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:rgba(191,90,242,.8);border:1px solid rgba(191,90,242,.25);border-radius:999px;padding:.3rem .9rem;margin-bottom:1rem;background:rgba(191,90,242,.07)}"
    ".hero-title{font-size:3.8rem;font-weight:800;letter-spacing:-.04em;line-height:1;margin-bottom:.75rem;background:linear-gradient(135deg,#E879F9 0%,#BF5AF2 50%,#9D4EDD 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}"
    ".hero-subtitle{color:rgba(235,235,245,.45);font-size:.92rem;font-weight:400;letter-spacing:.01em}"
    ".step-wrap{display:flex;align-items:center;gap:12px;margin:2rem 0 1.2rem}"
    ".step-num{width:28px;height:28px;border-radius:50%;flex-shrink:0;background:linear-gradient(135deg,#BF5AF2,#7B2FBE);color:#fff;font-size:.75rem;font-weight:700;display:flex;align-items:center;justify-content:center;box-shadow:0 0 16px rgba(191,90,242,.45)}"
    ".step-text{font-size:1.05rem;font-weight:600;color:#F5F5F7;letter-spacing:-.01em}"
    "[data-testid='stFileUploader']{border:1.5px dashed rgba(191,90,242,.35)!important;border-radius:16px!important;background:rgba(191,90,242,.04)!important;transition:all .25s ease!important}"
    "[data-testid='stFileUploader']:hover{border-color:rgba(191,90,242,.65)!important;background:rgba(191,90,242,.08)!important}"
    "[data-testid='stFileUploadDropzone']{border:none!important;background:transparent!important}"
    "[data-testid='stButton']>button{background:linear-gradient(135deg,#BF5AF2 0%,#9D4EDD 100%)!important;color:#fff!important;border:none!important;border-radius:14px!important;font-weight:600!important;font-size:1rem!important;letter-spacing:.01em!important;padding:.75rem 1.5rem!important;box-shadow:0 4px 24px rgba(191,90,242,.3)!important;transition:all .22s ease!important}"
    "[data-testid='stButton']>button:hover{transform:translateY(-2px)!important;box-shadow:0 10px 36px rgba(191,90,242,.5)!important}"
    "[data-testid='stButton']>button:active{transform:translateY(0px)!important}"
    "[data-testid='stDownloadButton']>button{background:linear-gradient(135deg,#30D158 0%,#1DA844 100%)!important;color:#fff!important;border:none!important;border-radius:14px!important;font-weight:600!important;font-size:1rem!important;width:100%!important;padding:.8rem!important;box-shadow:0 4px 20px rgba(48,209,88,.25)!important;transition:all .22s ease!important}"
    "[data-testid='stDownloadButton']>button:hover{transform:translateY(-2px)!important;box-shadow:0 10px 32px rgba(48,209,88,.4)!important}"
    "[data-testid='stMetric']{background:rgba(255,255,255,.04)!important;border:1px solid rgba(255,255,255,.08)!important;border-radius:16px!important;padding:1.2rem 1.4rem!important;transition:all .2s ease!important}"
    "[data-testid='stMetric']:hover{background:rgba(255,255,255,.06)!important;border-color:rgba(191,90,242,.3)!important}"
    "[data-testid='stMetricLabel']{font-size:.72rem!important;font-weight:600!important;text-transform:uppercase!important;letter-spacing:.08em!important;color:rgba(235,235,245,.45)!important}"
    "[data-testid='stMetricValue']{font-size:2rem!important;font-weight:700!important;letter-spacing:-.025em!important;color:#F5F5F7!important}"
    "[data-testid='stMetricDelta']{font-size:.82rem!important;font-weight:500!important}"
    "[data-testid='stProgressBar']>div{border-radius:6px!important;background:rgba(255,255,255,.08)!important;height:6px!important}"
    "[data-testid='stProgressBar']>div>div{background:linear-gradient(90deg,#E879F9,#BF5AF2,#9D4EDD)!important;border-radius:6px!important}"
    ".stSpinner>div{border-color:rgba(191,90,242,.15)!important;border-top-color:#BF5AF2!important}"
    "[data-baseweb='select']>div{border-radius:10px!important;border-color:rgba(255,255,255,.12)!important;background:rgba(255,255,255,.05)!important;transition:border-color .2s!important}"
    "[data-baseweb='select']>div:hover{border-color:rgba(191,90,242,.45)!important}"
    "[data-baseweb='input']{border-radius:10px!important;background:rgba(255,255,255,.05)!important;border-color:rgba(255,255,255,.12)!important}"
    "[data-baseweb='tag']{background:rgba(191,90,242,.18)!important;border:1px solid rgba(191,90,242,.3)!important;border-radius:6px!important}"
    "[data-testid='stExpander']{border:1px solid rgba(255,255,255,.08)!important;border-radius:14px!important;overflow:hidden!important;background:rgba(255,255,255,.025)!important}"
    "[data-testid='stNotification'],.stAlert{border-radius:12px!important;border-left:none!important}"
    "[data-testid='stDataFrame']{border-radius:14px!important;overflow:hidden!important;border:1px solid rgba(255,255,255,.08)!important}"
    "hr{border:none!important;border-top:1px solid rgba(255,255,255,.07)!important;margin:1.8rem 0!important}"
    ".stCaption,[data-testid='stCaptionContainer']{color:rgba(235,235,245,.38)!important;font-size:.78rem!important}"
    ".footer{text-align:center;color:rgba(235,235,245,.22);font-size:.75rem;letter-spacing:.04em;margin-top:4rem;padding-top:1.5rem;border-top:1px solid rgba(255,255,255,.06)}"
    "</style>",
    unsafe_allow_html=True,
)


# ── Funções de limpeza ─────────────────────────────────────────────────────────

def clean_phone(raw) -> str:
    """
    Trata número com maldade:
    - Remove espaços, traços, pontos, parênteses, barras
    - Remove código do país +55 / 55 quando número tem > 11 dígitos
    - Remove 0 inicial antigo
    - Retorna só dígitos
    """
    if raw is None:
        return ""
    if isinstance(raw, float):
        if np.isnan(raw):
            return ""
        raw = str(int(raw))  # evita "11987654321.0"
    raw = str(raw).strip()
    if raw.lower() in ("", "nan", "none", "-", "n/a", "#n/a"):
        return ""

    # Mantém só dígitos e +
    digits = re.sub(r"[^\d+]", "", raw).lstrip("+")

    # Remove DDI 55 se número ficar muito longo
    if digits.startswith("55") and len(digits) > 11:
        digits = digits[2:]

    # Remove 0 de discagem longa (ex: 011...)
    if digits.startswith("0") and len(digits) >= 11:
        digits = digits[1:]

    return digits


def right8(phone_clean: str) -> str:
    """Últimos 8 dígitos — padroniza números com/sem 9 e com/sem DDD."""
    if not phone_clean:
        return ""
    digits = re.sub(r"\D", "", str(phone_clean))
    return digits[-8:] if len(digits) >= 8 else digits


# ── Detecção automática de colunas ─────────────────────────────────────────────

def _match_keywords(col_name: str, keywords: list[str]) -> bool:
    c = col_name.lower().strip()
    return any(kw in c for kw in keywords)


def detect_phone_col(df: pd.DataFrame) -> Optional[str]:
    kws = [
        "telefone", "celular", "whatsapp", "wpp", "fone", "phone",
        "mobile", "cel", "numero", "número", "tel", "contato",
        "phone_number", "nr_tel", "nro", "n_tel",
    ]
    for col in df.columns:
        if _match_keywords(col, kws):
            return col
    # Detecção por conteúdo
    for col in df.columns:
        sample = df[col].dropna().astype(str).head(20)
        hits = sample.apply(
            lambda x: bool(re.search(r"\d{7,}", re.sub(r"[\s\-\(\)\.+]", "", x)))
        )
        if hits.sum() >= max(2, len(sample) * 0.4):
            return col
    return None


def detect_tag_col(df: pd.DataFrame) -> Optional[str]:
    kws = ["tag", "etiqueta", "label", "categoria", "tipo", "fonte", "origem"]
    for col in df.columns:
        if _match_keywords(col, kws):
            return col
    return None


# ── Carregamento de arquivo ────────────────────────────────────────────────────

def get_excel_sheets(uploaded) -> list:
    """Retorna lista de abas de um Excel. Lista vazia indica CSV."""
    name = uploaded.name.lower()
    if not name.endswith((".xlsx", ".xls", ".xlsm")):
        return []
    try:
        uploaded.seek(0)
        return pd.ExcelFile(uploaded).sheet_names
    except Exception:
        return []


def detect_header_row(df_raw: pd.DataFrame) -> int:
    """
    Varre as primeiras linhas e retorna o índice da linha que parece
    ser o cabeçalho real (ignora títulos e linhas em branco acima).
    """
    n_cols = len(df_raw.columns)
    min_filled = max(2, int(n_cols * 0.3))

    for i in range(min(10, len(df_raw))):
        row = df_raw.iloc[i]
        non_null = [v for v in row if pd.notna(v) and str(v).strip() not in ("", "nan")]
        if len(non_null) < min_filled:
            continue
        label_count = sum(
            1 for v in non_null
            if isinstance(v, str)
            and not re.fullmatch(r"[\d\s\.\,\-\+\/\%]+", v.strip())
        )
        if label_count / len(non_null) >= 0.4:
            return i
    return 0


def _load_single_sheet(xls, sheet_name: str) -> tuple[pd.DataFrame, int]:
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str)
    hrow = detect_header_row(df_raw)
    df = pd.read_excel(xls, sheet_name=sheet_name, header=hrow, dtype=str)
    df = df.dropna(how="all").reset_index(drop=True)
    return df, hrow


def _load_csv_smart(uploaded) -> tuple[pd.DataFrame, int]:
    for enc in ("utf-8", "latin-1", "cp1252", "iso-8859-1"):
        try:
            uploaded.seek(0)
            df_raw = pd.read_csv(uploaded, encoding=enc, header=None, dtype=str)
            hrow = detect_header_row(df_raw)
            uploaded.seek(0)
            df = pd.read_csv(uploaded, encoding=enc, header=hrow, dtype=str)
            df = df.dropna(how="all").reset_index(drop=True)
            return df, hrow
        except UnicodeDecodeError:
            continue
        except Exception:
            break
    return pd.DataFrame(), 0


def load_file_multisheet(
    uploaded, selected_sheets: list
) -> tuple[Optional[pd.DataFrame], dict]:
    """
    Carrega uma ou mais abas. Retorna (df_combinado, {aba: linha_cabecalho}).
    Para CSV, selected_sheets é ignorado.
    """
    name = uploaded.name.lower()
    sheet_info: dict = {}
    dfs = []

    try:
        if name.endswith(".csv"):
            df, hrow = _load_csv_smart(uploaded)
            if len(df) > 0:
                sheet_info["CSV"] = hrow
                dfs.append(df)
        elif name.endswith((".xlsx", ".xls", ".xlsm")):
            uploaded.seek(0)
            xls = pd.ExcelFile(uploaded)
            for sheet in selected_sheets:
                df, hrow = _load_single_sheet(xls, sheet)
                if len(df) > 0:
                    if len(selected_sheets) > 1:
                        df.insert(0, "_Planilha", sheet)
                    sheet_info[sheet] = hrow
                    dfs.append(df)
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {e}")
        return None, {}

    if not dfs:
        return None, {}

    return pd.concat(dfs, ignore_index=True), sheet_info


# ── Lógica principal do PROCV ──────────────────────────────────────────────────

def run_procv(
    df_sales: pd.DataFrame,
    sales_phone_col: str,
    df_kommo: pd.DataFrame,
    kommo_phone_col: str,
    kommo_tag_col: str,
    traffic_keyword: str,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Retorna: (vendas_tratada, kommo_tratada, resultado_trafego, resultado_completo)
    """

    # ── Trata planilha de vendas ──────────────────────────────────────────────
    ds = df_sales.copy()
    pos = ds.columns.get_loc(sales_phone_col) + 1
    ds.insert(pos, "Tel_Limpo_Vendas", ds[sales_phone_col].apply(clean_phone))
    ds.insert(pos + 1, "Tel_8dig_Vendas", ds["Tel_Limpo_Vendas"].apply(right8))

    # ── Trata planilha do Kommo ───────────────────────────────────────────────
    dk = df_kommo.copy()
    pos_k = dk.columns.get_loc(kommo_phone_col) + 1
    dk.insert(pos_k, "Tel_Limpo_Kommo", dk[kommo_phone_col].apply(clean_phone))
    dk.insert(pos_k + 1, "Tel_8dig_Kommo", dk["Tel_Limpo_Kommo"].apply(right8))

    # ── Monta dicionário de lookup: 8dig → linha de vendas ───────────────────
    lookup: dict[str, dict] = {}
    for _, row in ds.iterrows():
        key = row["Tel_8dig_Vendas"]
        if key and key not in lookup:
            lookup[key] = row.to_dict()

    # ── PROCV: cruza Kommo × Vendas ───────────────────────────────────────────
    result_rows = []
    for _, kr in dk.iterrows():
        k8 = kr["Tel_8dig_Kommo"]
        tag_raw = "" if pd.isna(kr.get(kommo_tag_col, np.nan)) else str(kr[kommo_tag_col])
        is_traffic = traffic_keyword.lower() in tag_raw.lower()
        sales_match = lookup.get(k8)

        row_out = {
            "Tag_Kommo": tag_raw,
            "Telefone_Kommo": kr.get(kommo_phone_col, ""),
            "Tel_8dig": k8,
            "É_Tráfego": "SIM" if is_traffic else "NÃO",
            "Venda_Confirmada": "SIM" if sales_match else "NÃO",
        }
        for col in df_sales.columns:
            row_out[f"[Venda] {col}"] = sales_match.get(col, "") if sales_match else ""

        result_rows.append(row_out)

    df_full = pd.DataFrame(result_rows)
    df_trafego = df_full[
        (df_full["É_Tráfego"] == "SIM") & (df_full["Venda_Confirmada"] == "SIM")
    ].copy()

    return ds, dk, df_trafego, df_full


# ── Busca automática de melhor combinação de colunas ──────────────────────────

def find_best_column_combo(
    df_sales: pd.DataFrame,
    df_kommo: pd.DataFrame,
    traffic_keyword: str,
) -> Optional[tuple[str, str, str, int]]:
    """
    Testa todas as combinações (col_tel_vendas × col_tel_kommo × col_tag_kommo)
    e retorna a que gera mais conversões de tráfego, ou None se nenhuma encontrar.
    """
    # Pré-computa chaves 8-dig para cada coluna de vendas
    sales_sets: dict = {}
    for col in df_sales.columns:
        cleaned = df_sales[col].apply(
            lambda v: right8(clean_phone(str(v))) if pd.notna(v) else ""
        )
        valid = set(v for v in cleaned if v)
        if len(valid) >= 2:
            sales_sets[col] = valid

    if not sales_sets:
        return None

    # Pré-computa chaves 8-dig para cada coluna do kommo
    kommo_cleaned: dict = {}
    for col in df_kommo.columns:
        cleaned = df_kommo[col].apply(
            lambda v: right8(clean_phone(str(v))) if pd.notna(v) else ""
        )
        if (cleaned != "").sum() >= 2:
            kommo_cleaned[col] = cleaned

    if not kommo_cleaned:
        return None

    best: tuple = (None, None, None, 0)

    for t_col in df_kommo.columns:
        traffic_mask = (
            df_kommo[t_col].fillna("").astype(str)
            .str.lower()
            .str.contains(traffic_keyword.lower(), regex=False)
        )
        if traffic_mask.sum() == 0:
            continue

        for k_col, k_series in kommo_cleaned.items():
            traffic_keys = set(k_series[traffic_mask])
            traffic_keys.discard("")
            if not traffic_keys:
                continue

            for s_col, s_set in sales_sets.items():
                conv = len(traffic_keys & s_set)
                if conv > best[3]:
                    best = (s_col, k_col, t_col, conv)

    return best if best[3] > 0 else None


# ── Utilitários de data ────────────────────────────────────────────────────────

_MONTH_MAP = {
    "janeiro": 1, "jan": 1, "fevereiro": 2, "fev": 2,
    "março": 3, "marco": 3, "mar": 3, "abril": 4, "abr": 4,
    "maio": 5, "mai": 5, "junho": 6, "jun": 6,
    "julho": 7, "jul": 7, "agosto": 8, "ago": 8,
    "setembro": 9, "set": 9, "outubro": 10, "out": 10,
    "novembro": 11, "nov": 11, "dezembro": 12, "dez": 12,
    "january": 1, "february": 2, "feb": 2, "march": 3, "april": 4, "apr": 4,
    "may": 5, "june": 6, "july": 7, "august": 8, "aug": 8,
    "september": 9, "sep": 9, "october": 10, "oct": 10,
    "november": 11, "december": 12, "dec": 12,
}

def _normalize(s: str) -> str:
    return (s.lower()
            .replace("ç", "c").replace("ã", "a").replace("á", "a")
            .replace("â", "a").replace("é", "e").replace("ê", "e")
            .replace("í", "i").replace("ó", "o").replace("ô", "o")
            .replace("ú", "u").replace("õ", "o"))


def parse_date(value) -> Optional[datetime]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    s = str(value).strip()
    if not s or s.lower() in ("nan", "none", "-", "n/a", "#n/a", ""):
        return None

    # Excel serial number como string (ex: "45678")
    try:
        serial = int(float(s))
        if 35000 < serial < 60000:
            return datetime(1899, 12, 30) + timedelta(days=serial)
    except (ValueError, OverflowError):
        pass

    # Formatos padrão via pandas (DD/MM/YYYY, ISO, etc.)
    try:
        return pd.to_datetime(s, dayfirst=True).to_pydatetime()
    except Exception:
        pass

    # "month_name/YY", "month_name/YYYY", "month_name YYYY" (ex: "abril/26", "março 2024")
    sn = _normalize(s)
    m = re.match(r"^([a-z]+)[/\s\-](\d{2,4})$", sn)
    if m:
        month = _MONTH_MAP.get(m.group(1))
        if month:
            year = int(m.group(2))
            if year < 100:
                year += 2000
            try:
                return datetime(year, month, 1)
            except ValueError:
                pass

    # "DD de month_name de YYYY" / "DD month_name YYYY"
    m = re.match(r"^(\d{1,2})\s+(?:de\s+)?([a-z]+)(?:\s+(?:de\s+)?(\d{2,4}))?$", sn)
    if m:
        month = _MONTH_MAP.get(m.group(2))
        if month:
            year = int(m.group(3)) if m.group(3) else datetime.now().year
            if year < 100:
                year += 2000
            try:
                return datetime(year, month, int(m.group(1)))
            except ValueError:
                pass

    return None


def detect_date_col(df: pd.DataFrame) -> Optional[str]:
    kws = ["data", "date", "dt", "dia", "quando", "periodo", "período",
           "competencia", "competência", "venda_em", "criado", "created", "hora"]
    for col in df.columns:
        if _match_keywords(col, kws):
            return col
    for col in df.columns:
        sample = df[col].dropna().head(15)
        hits = sum(1 for v in sample if parse_date(v) is not None)
        if hits >= max(2, len(sample) * 0.5):
            return col
    return None


# ── Análise de Disparo ─────────────────────────────────────────────────────────

def run_disparo(
    df_sales: pd.DataFrame,
    sales_phone_col: str,
    sales_date_col: Optional[str],
    df_kommo: pd.DataFrame,
    kommo_phone_col: str,
    kommo_tag_col: str,
    disparo_keyword: str,
    kommo_date_col: Optional[str],
    window_days: int = 30,
) -> pd.DataFrame:
    """
    Filtra leads de disparo do Kommo pela tag e cruza com vendas.
    Janela de tempo: venda entre 0 e window_days após a data do disparo no Kommo.
    Sem datas: inclui todos os matches para filtro manual.
    """
    # Filtra leads de disparo no Kommo
    disparo_mask = (
        df_kommo[kommo_tag_col].fillna("").astype(str)
        .str.lower()
        .str.contains(disparo_keyword.lower(), regex=False)
    )
    df_disp_leads = df_kommo[disparo_mask].copy()

    if len(df_disp_leads) == 0:
        return pd.DataFrame()

    # Pré-processa vendas
    ds = df_sales.copy()
    ds["_tel8"] = ds[sales_phone_col].apply(
        lambda v: right8(clean_phone(str(v))) if pd.notna(v) else ""
    )
    if sales_date_col:
        ds["_dt_venda"] = ds[sales_date_col].apply(parse_date)

    sales_lookup: dict = {}
    for _, row in ds.iterrows():
        key = row["_tel8"]
        if not key:
            continue
        sales_lookup.setdefault(key, []).append(row.to_dict())

    result_rows = []
    for _, kr in df_disp_leads.iterrows():
        tel_raw = kr.get(kommo_phone_col, "")
        tel8 = right8(clean_phone(str(tel_raw))) if pd.notna(tel_raw) else ""
        if not tel8:
            continue

        disp_date = parse_date(kr.get(kommo_date_col)) if kommo_date_col else None
        candidates = sales_lookup.get(tel8, [])

        matched_sale = None
        for sale in candidates:
            if disp_date and sales_date_col:
                sale_dt = sale.get("_dt_venda")
                if sale_dt:
                    delta = (sale_dt - disp_date).days
                    if 0 <= delta <= window_days:
                        matched_sale = sale
                        break
                else:
                    matched_sale = sale
                    break
            else:
                matched_sale = sale
                break

        tag_raw = str(kr.get(kommo_tag_col, ""))
        row_out: dict = {
            "Tag_Kommo": tag_raw,
            "Telefone_Disparo": tel_raw,
            "Tel_8dig": tel8,
        }

        if kommo_date_col:
            row_out["Data_Disparo"] = kr.get(kommo_date_col, "")

        if matched_sale:
            row_out["Venda_Confirmada"] = "SIM"
            if sales_date_col:
                sale_dt = matched_sale.get("_dt_venda")
                row_out["Data_Venda"] = (
                    sale_dt.strftime("%d/%m/%Y") if sale_dt
                    else str(matched_sale.get(sales_date_col, ""))
                )
                if disp_date and sale_dt:
                    row_out["Dias_Após_Disparo"] = (sale_dt - disp_date).days
            for col in df_sales.columns:
                row_out[f"[Venda] {col}"] = matched_sale.get(col, "")
        else:
            row_out["Venda_Confirmada"] = "NÃO"
            if sales_date_col:
                row_out["Data_Venda"] = ""
                if disp_date:
                    row_out["Dias_Após_Disparo"] = ""
            for col in df_sales.columns:
                row_out[f"[Venda] {col}"] = ""

        result_rows.append(row_out)

    return pd.DataFrame(result_rows) if result_rows else pd.DataFrame()


# ── Geração do Excel formatado ─────────────────────────────────────────────────

def _col_width(df: pd.DataFrame, col: str) -> float:
    max_data = df[col].dropna().astype(str).apply(len).max() if len(df) > 0 else 10
    return min(float(max(len(str(col)), max_data)) + 4, 45)


def _write_sheet(
    ws,
    df: pd.DataFrame,
    title: str,
    header_hex: str,
    highlight_cols: Optional[List[str]] = None,
):
    THIN = Border(
        left=Side(style="thin", color="DDDDDD"),
        right=Side(style="thin", color="DDDDDD"),
        top=Side(style="thin", color="DDDDDD"),
        bottom=Side(style="thin", color="DDDDDD"),
    )
    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center")

    # Título
    last_col = get_column_letter(max(len(df.columns), 1))
    ws.merge_cells(f"A1:{last_col}1")
    tc = ws.cell(1, 1, title)
    tc.font = Font(color="FFFFFF", bold=True, size=13)
    tc.fill = PatternFill("solid", fgColor=header_hex)
    tc.alignment = CENTER
    ws.row_dimensions[1].height = 32

    # Cabeçalhos
    for c, col in enumerate(df.columns, 1):
        cell = ws.cell(2, c, str(col))
        cell.font = Font(color="FFFFFF", bold=True, size=10)
        cell.fill = PatternFill("solid", fgColor="222222")
        cell.alignment = CENTER
        cell.border = THIN
    ws.row_dimensions[2].height = 24

    # Dados
    highlight_set = set(highlight_cols or [])
    for r, row_data in enumerate(df.itertuples(index=False), 3):
        zebra = (r - 3) % 2 == 1
        for c, value in enumerate(row_data, 1):
            col_name = df.columns[c - 1]
            cell = ws.cell(r, c, value if value != "" else None)
            cell.font = Font(size=10)
            cell.alignment = LEFT
            cell.border = THIN

            if col_name == "Venda_Confirmada":
                if value == "SIM":
                    cell.fill = PatternFill("solid", fgColor="27AE60")
                    cell.font = Font(color="FFFFFF", bold=True, size=10)
                elif value == "NÃO":
                    cell.fill = PatternFill("solid", fgColor="E74C3C")
                    cell.font = Font(color="FFFFFF", bold=True, size=10)
            elif col_name == "É_Tráfego" and value == "SIM":
                cell.fill = PatternFill("solid", fgColor="2980B9")
                cell.font = Font(color="FFFFFF", bold=True, size=10)
            elif col_name in highlight_set:
                cell.fill = PatternFill("solid", fgColor="FFE5D9" if not zebra else "FFD5B8")
            elif zebra:
                cell.fill = PatternFill("solid", fgColor="F5F5F5")
        ws.row_dimensions[r].height = 17

    # Largura das colunas
    for c, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(c)].width = _col_width(df, col)

    ws.freeze_panes = "A3"
    if len(df) > 0:
        ws.auto_filter.ref = f"A2:{last_col}{len(df)+2}"


def build_excel(
    ds: pd.DataFrame,
    dk: pd.DataFrame,
    df_result: pd.DataFrame,
    df_full: pd.DataFrame,
    df_disparo_result: Optional[pd.DataFrame] = None,
) -> bytes:
    wb = Workbook()

    # Sheet 1 — Vendas Tratada
    ws1 = wb.active
    ws1.title = "Vendas Tratada"
    _write_sheet(ws1, ds, "PLANILHA DE VENDAS — TRATADA", "FF6B35",
                 ["Tel_Limpo_Vendas", "Tel_8dig_Vendas"])

    # Sheet 2 — Kommo Tratada
    ws2 = wb.create_sheet("Kommo Tratada")
    _write_sheet(ws2, dk, "PLANILHA KOMMO — TRATADA", "2980B9",
                 ["Tel_Limpo_Kommo", "Tel_8dig_Kommo"])

    # Sheet 3 — Resultado PROCV (só tráfego com venda)
    ws3 = wb.create_sheet("Resultado PROCV")
    if len(df_result) == 0:
        ws3.cell(1, 1, "Nenhum lead de tráfego com venda confirmada encontrado.")
        ws3.cell(1, 1).font = Font(italic=True, color="888888")
    else:
        _write_sheet(ws3, df_result,
                     "LEADS DE TRÁFEGO COM VENDA CONFIRMADA ✅", "27AE60")

    # Sheet 4 — Kommo Completo com PROCV (todas as linhas, filtrável)
    ws4 = wb.create_sheet("Kommo Completo + PROCV")
    _write_sheet(ws4, df_full,
                 "KOMMO COMPLETO — TODAS AS LINHAS (USE O FILTRO)", "6C3483")

    # Sheet 5 — Resultado Disparo (opcional)
    all_sheets = [ws1, ws2, ws3, ws4]
    if df_disparo_result is not None and len(df_disparo_result) > 0:
        ws_disp = wb.create_sheet("Resultado Disparo")
        disp_confirmed = int((df_disparo_result["Venda_Confirmada"] == "SIM").sum())
        _write_sheet(
            ws_disp, df_disparo_result,
            f"DISPARO — {disp_confirmed} CONVERSÕES CONFIRMADAS",
            "8E44AD",
            ["Data_Disparo", "Data_Venda", "Dias_Após_Disparo"],
        )
        all_sheets.append(ws_disp)

    # Sheet Resumo
    ws_res = wb.create_sheet("Resumo")
    all_sheets.append(ws_res)
    total_traffic = int((df_full["É_Tráfego"] == "SIM").sum())
    total_sales = len(ds)
    confirmed = len(df_result)
    conv_rate = f"{confirmed/total_traffic*100:.1f}%" if total_traffic > 0 else "—"

    ws_res.merge_cells("A1:C1")
    t = ws_res.cell(1, 1, "RESUMO DO PROC-BULLS")
    t.font = Font(color="FFFFFF", bold=True, size=14)
    t.fill = PatternFill("solid", fgColor="FF6B35")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws_res.row_dimensions[1].height = 36

    summary_rows = [
        ("Total de vendas carregadas", total_sales),
        ("Total de leads no Kommo", len(dk)),
        ("Leads com tag de tráfego", total_traffic),
        ("Conversões confirmadas (tráfego → venda)", confirmed),
        ("Taxa de conversão do tráfego", conv_rate),
    ]
    if df_disparo_result is not None and len(df_disparo_result) > 0:
        disp_total = len(df_disparo_result)
        disp_conv = int((df_disparo_result["Venda_Confirmada"] == "SIM").sum())
        disp_rate = f"{disp_conv/disp_total*100:.1f}%" if disp_total > 0 else "—"
        summary_rows += [
            ("— DISPARO —", ""),
            ("Total de disparos analisados", disp_total),
            ("Conversões confirmadas (disparo → venda)", disp_conv),
            ("Taxa de conversão do disparo", disp_rate),
        ]

    for i, (label, val) in enumerate(summary_rows, 2):
        ws_res.cell(i, 1, label).font = Font(bold=True, size=11)
        c = ws_res.cell(i, 2, val)
        c.font = Font(bold=True, size=13, color="FF6B35")
        c.alignment = Alignment(horizontal="center")
        ws_res.row_dimensions[i].height = 26

    ws_res.column_dimensions["A"].width = 46
    ws_res.column_dimensions["B"].width = 22

    for ws in all_sheets:
        r = ws.max_row + 2
        ws.cell(r, 1, "Desenvolvido por João  ·  Proc-Bulls  ·  Aure Digital").font = Font(
            italic=True, color="AAAAAA", size=9
        )

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Interface ──────────────────────────────────────────────────────────────────

st.markdown("""
<div class="hero-wrap">
  <div class="hero-badge">Aure Digital</div>
  <h1 class="hero-title" translate="no">PROC-BULLS</h1>
  <p class="hero-subtitle">Análise inteligente de conversão · Tráfego & Disparo</p>
</div>
""", unsafe_allow_html=True)
st.divider()

# ── Passo 1: Upload ────────────────────────────────────────────────────────────
st.markdown('<div class="step-wrap"><div class="step-num">1</div><div class="step-text">Carregue as planilhas</div></div>', unsafe_allow_html=True)

col_left, col_right = st.columns(2)

with col_left:
    st.markdown("#### 📊 Planilha de Vendas")
    sales_file = st.file_uploader(
        "Arraste ou clique (Excel ou CSV)",
        type=["xlsx", "xls", "csv", "xlsm"],
        key="sales_upload",
        help="Planilha do cliente com nome, telefone e valor de compra",
    )
    df_sales_raw: Optional[pd.DataFrame] = None
    if sales_file:
        sales_sheets = get_excel_sheets(sales_file)
        selected_sales_sheets = sales_sheets

        if len(sales_sheets) > 1:
            st.caption(f"📂 {len(sales_sheets)} planilhas encontradas neste arquivo")
            selected_sales_sheets = st.multiselect(
                "Selecione as planilhas a usar:",
                options=sales_sheets,
                default=sales_sheets,
                key="sales_sheets_select",
            )
            if not selected_sales_sheets:
                st.warning("Selecione ao menos uma planilha.")

        if selected_sales_sheets or not sales_sheets:
            df_sales_raw, sales_info = load_file_multisheet(sales_file, selected_sales_sheets)
            if df_sales_raw is not None:
                for sheet, hrow in sales_info.items():
                    if hrow > 0:
                        st.info(f"📋 '{sheet}': cabeçalho detectado na linha {hrow + 1} — título(s) anteriores ignorados.")
                st.success(f"✅ {len(df_sales_raw):,} linhas · {len(df_sales_raw.columns)} colunas")
                with st.expander("Prévia"):
                    st.dataframe(df_sales_raw.head(6), use_container_width=True)

with col_right:
    st.markdown("#### 🗂️ Planilha do Kommo")
    kommo_file = st.file_uploader(
        "Arraste ou clique (Excel ou CSV)",
        type=["xlsx", "xls", "csv", "xlsm"],
        key="kommo_upload",
        help="Export do Kommo CRM com leads, telefones e tags",
    )
    df_kommo_raw: Optional[pd.DataFrame] = None
    if kommo_file:
        kommo_sheets = get_excel_sheets(kommo_file)
        selected_kommo_sheets = kommo_sheets

        if len(kommo_sheets) > 1:
            st.caption(f"📂 {len(kommo_sheets)} planilhas encontradas neste arquivo")
            selected_kommo_sheets = st.multiselect(
                "Selecione as planilhas a usar:",
                options=kommo_sheets,
                default=kommo_sheets,
                key="kommo_sheets_select",
            )
            if not selected_kommo_sheets:
                st.warning("Selecione ao menos uma planilha.")

        if selected_kommo_sheets or not kommo_sheets:
            df_kommo_raw, kommo_info = load_file_multisheet(kommo_file, selected_kommo_sheets)
            if df_kommo_raw is not None:
                for sheet, hrow in kommo_info.items():
                    if hrow > 0:
                        st.info(f"📋 '{sheet}': cabeçalho detectado na linha {hrow + 1} — título(s) anteriores ignorados.")
                st.success(f"✅ {len(df_kommo_raw):,} linhas · {len(df_kommo_raw.columns)} colunas")
                with st.expander("Prévia"):
                    st.dataframe(df_kommo_raw.head(6), use_container_width=True)

st.divider()

# ── Passo 2: Configuração ──────────────────────────────────────────────────────
if df_sales_raw is not None and df_kommo_raw is not None:
    st.markdown('<div class="step-wrap"><div class="step-num">2</div><div class="step-text">Configure as colunas</div></div>', unsafe_allow_html=True)

    auto_sp = detect_phone_col(df_sales_raw)
    auto_kp = detect_phone_col(df_kommo_raw)
    auto_kt = detect_tag_col(df_kommo_raw)

    def _idx(df, col):
        cols = list(df.columns)
        return cols.index(col) if col in cols else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        sales_phone_col = st.selectbox(
            "Coluna de telefone — Vendas",
            list(df_sales_raw.columns),
            index=_idx(df_sales_raw, auto_sp),
        )
        if auto_sp == sales_phone_col:
            st.caption("✨ Auto-detectado")
    with c2:
        kommo_phone_col = st.selectbox(
            "Coluna de telefone — Kommo",
            list(df_kommo_raw.columns),
            index=_idx(df_kommo_raw, auto_kp),
        )
        if auto_kp == kommo_phone_col:
            st.caption("✨ Auto-detectado")
    with c3:
        kommo_tag_col = st.selectbox(
            "Coluna de tags — Kommo",
            list(df_kommo_raw.columns),
            index=_idx(df_kommo_raw, auto_kt) if auto_kt else 0,
        )
        if auto_kt == kommo_tag_col:
            st.caption("✨ Auto-detectado")
    with c4:
        traffic_keyword = st.text_input(
            "Palavra-chave da tag de tráfego",
            value="trafego",
            help="A busca é case-insensitive e parcial. Ex: 'trafego' encontra 'Lead Tráfego Pago'.",
        )

    # ── Configuração de Disparo ────────────────────────────────────────────────
    st.markdown("##### 📣 Disparo")
    none_opt = "(não usar)"
    auto_kd = detect_date_col(df_kommo_raw)
    auto_sd = detect_date_col(df_sales_raw)

    e1, e2, e3 = st.columns(3)
    with e1:
        disparo_keyword = st.text_input(
            "Palavra-chave da tag de disparo",
            value="disparo",
            help="Leads do Kommo com essa tag serão analisados como disparo.",
            key="disparo_kw",
        )
    with e2:
        kommo_date_opts = [none_opt] + list(df_kommo_raw.columns)
        kommo_date_default = (kommo_date_opts.index(auto_kd)
                              if auto_kd and auto_kd in kommo_date_opts else 0)
        kommo_date_sel = st.selectbox(
            "Data do disparo no Kommo *(opcional)*",
            kommo_date_opts,
            index=kommo_date_default,
            key="kommo_date",
        )
        kommo_date_col = None if kommo_date_sel == none_opt else kommo_date_sel
        if auto_kd == kommo_date_col:
            st.caption("✨ Auto-detectado")
    with e3:
        sales_date_opts = [none_opt] + list(df_sales_raw.columns)
        sales_date_default = (sales_date_opts.index(auto_sd)
                              if auto_sd and auto_sd in sales_date_opts else 0)
        sales_date_sel = st.selectbox(
            "Data da venda *(opcional)*",
            sales_date_opts,
            index=sales_date_default,
            key="sales_date",
        )
        sales_date_col = None if sales_date_sel == none_opt else sales_date_sel
        if auto_sd == sales_date_col:
            st.caption("✨ Auto-detectado")

    if not kommo_date_col:
        st.caption("⚠️ Sem data do disparo: todos os matches de disparo serão incluídos — filtre pelo campo Data_Venda no Excel.")

    considerar_disparo = st.checkbox(
        "📣 Considerar disparo também",
        value=True,
        help="Quando marcado, analisa conversões via disparo mesmo que já existam vendas pelo tráfego.",
    )

    st.divider()

    # ── Passo 3: Processar ─────────────────────────────────────────────────────
    st.markdown('<div class="step-wrap"><div class="step-num">3</div><div class="step-text">Processar</div></div>', unsafe_allow_html=True)

    if st.button("🎯  RODAR PROC-BULLS", use_container_width=True, type="primary"):
        progress = st.progress(0, text="Iniciando...")
        try:
            progress.progress(10, text="Tratando planilha de vendas...")
            ds_t, dk_t, df_result, df_full = run_procv(
                df_sales_raw, sales_phone_col,
                df_kommo_raw, kommo_phone_col, kommo_tag_col,
                traffic_keyword,
            )

            confirmed = len(df_result)
            auto_combo = None

            if confirmed == 0:
                progress.progress(40, text="Buscando combinações alternativas de colunas...")
                auto_combo = find_best_column_combo(df_sales_raw, df_kommo_raw, traffic_keyword)
                if auto_combo:
                    progress.progress(55, text="Reprocessando com colunas otimizadas...")
                    s_col, k_col, t_col, _ = auto_combo
                    ds_t, dk_t, df_result, df_full = run_procv(
                        df_sales_raw, s_col,
                        df_kommo_raw, k_col, t_col,
                        traffic_keyword,
                    )
                    confirmed = len(df_result)

            progress.progress(70, text="Analisando disparo...")
            df_disparo_result = None
            if considerar_disparo and disparo_keyword.strip():
                df_disparo_result = run_disparo(
                    df_sales_raw, sales_phone_col, sales_date_col,
                    df_kommo_raw, kommo_phone_col, kommo_tag_col,
                    disparo_keyword, kommo_date_col,
                )

            progress.progress(90, text="Gerando relatório Excel...")
            excel_bytes = build_excel(ds_t, dk_t, df_result, df_full, df_disparo_result)

            progress.progress(100, text="Concluído!")
            progress.empty()

            # ── Métricas ───────────────────────────────────────────────────────
            st.markdown('<div class="step-wrap"><div class="step-num">✓</div><div class="step-text">Resultados</div></div>', unsafe_allow_html=True)
            total_traffic = int((df_full["É_Tráfego"] == "SIM").sum())

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Vendas carregadas", f"{len(ds_t):,}")
            m2.metric("Leads no Kommo", f"{len(dk_t):,}")
            m3.metric("Leads de tráfego", f"{total_traffic:,}")
            m4.metric(
                "Conversões confirmadas",
                f"{confirmed:,}",
                delta=f"{confirmed/total_traffic*100:.1f}% de conv." if total_traffic > 0 else None,
            )

            st.divider()

            if auto_combo:
                s_col, k_col, t_col, n = auto_combo
                st.info(
                    f"💡 Colunas originais não geraram matches. "
                    f"A ferramenta encontrou **{confirmed}** conversão(ões) usando:\n\n"
                    f"- Telefone Vendas → **{s_col}**\n"
                    f"- Telefone Kommo → **{k_col}**\n"
                    f"- Tag Kommo → **{t_col}**"
                )

            if confirmed > 0:
                st.success(f"🎉 {confirmed} leads de tráfego com venda confirmada!")
                st.dataframe(df_result, use_container_width=True, height=280)
            elif auto_combo is None:
                st.warning(
                    "Nenhuma conversão de tráfego encontrada mesmo após testar todas as combinações de colunas."
                )

            # ── Resultado Disparo ───────────────────────────────────────────────
            st.divider()
            st.markdown('<div class="step-wrap"><div class="step-num">📣</div><div class="step-text">Resultado do Disparo</div></div>', unsafe_allow_html=True)

            if not considerar_disparo:
                st.caption("Análise de disparo desativada.")
            elif df_disparo_result is None or len(df_disparo_result) == 0:
                st.warning(f"Nenhum lead encontrado com a tag **\"{disparo_keyword}\"** no Kommo. Verifique a palavra-chave ou a coluna de tags selecionada.")
            else:
                disp_conv = int((df_disparo_result["Venda_Confirmada"] == "SIM").sum())
                disp_total = len(df_disparo_result)
                disp_rate = f"{disp_conv/disp_total*100:.1f}%" if disp_total > 0 else "—"

            if df_disparo_result is not None and len(df_disparo_result) > 0:
                dm1, dm2, dm3 = st.columns(3)
                dm1.metric("Leads de disparo", f"{disp_total:,}")
                dm2.metric("Conversões confirmadas", f"{disp_conv:,}")
                dm3.metric("Taxa de conversão", disp_rate)

                if kommo_date_col and sales_date_col:
                    st.success(f"🎯 {disp_conv} conversões dentro da janela de 30 dias após o disparo.")
                else:
                    st.info("📅 Datas parciais ou ausentes — filtre pela coluna Data_Venda no Excel para analisar por competência.")

                if disp_conv > 0:
                    st.dataframe(
                        df_disparo_result[df_disparo_result["Venda_Confirmada"] == "SIM"],
                        use_container_width=True, height=250,
                    )

            # ── Download ───────────────────────────────────────────────────────
            st.divider()
            st.download_button(
                label="📥  Baixar Excel — Resultado Completo",
                data=excel_bytes,
                file_name="proc_bulls_resultado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.caption(
                "💡 Para usar no Google Sheets: faça upload do .xlsx no Google Drive → "
                "clique com botão direito → Abrir com → Planilhas Google."
            )

        except Exception:
            progress.empty()
            st.error("Erro ao processar. Detalhes abaixo:")
            st.code(traceback.format_exc())

else:
    st.info("⬆️ Carregue as duas planilhas acima para continuar.")

# ── Rodapé ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    '<div class="footer">Desenvolvido por João &nbsp;·&nbsp; Proc-Bulls v1.0 &nbsp;·&nbsp; Aure Digital</div>',
    unsafe_allow_html=True,
)
