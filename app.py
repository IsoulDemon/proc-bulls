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

    # Trata "66999873776.0" — número do Excel lido como float string
    if re.match(r"^\d+\.0+$", raw):
        raw = str(int(float(raw)))

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
    """Últimos 8 dígitos — padroniza números com/sem 9 e com/sem DDD.
    Retorna '' se < 8 dígitos para evitar colisões com números curtos."""
    if not phone_clean:
        return ""
    digits = re.sub(r"\D", "", str(phone_clean))
    if len(digits) < 8:
        return ""  # Muito curto para ser telefone — descarta
    return digits[-8:]


# ── Detecção automática de colunas ─────────────────────────────────────────────

def _match_keywords(col_name: str, keywords: list[str]) -> bool:
    c = col_name.lower().strip()
    return any(kw in c for kw in keywords)


def detect_phone_col(df: pd.DataFrame) -> Optional[str]:
    """
    Detecta coluna de telefone por conteúdo (7+ dígitos) + bônus por nome.
    Funciona com colunas sem nome óbvio — cada planilha é um caso.
    """
    kws = [
        "telefone", "celular", "whatsapp", "wpp", "fone", "phone",
        "mobile", "cel", "numero", "número", "tel", "contato",
        "phone_number", "nr_tel", "nro", "n_tel",
    ]
    skip_kws = ["data", "date", "valor", "preco", "preço", "total",
                "cpf", "cnpj", "cep", "id", "código", "codigo"]

    def _phone_hits(col: str) -> int:
        sample = df[col].dropna().astype(str).head(30)
        if len(sample) == 0:
            return 0
        return int(sample.apply(
            lambda x: bool(re.search(r"\d{7,}", re.sub(r"[\s\-\(\)\.+]", "", x)))
        ).sum())

    candidates = []
    for col in df.columns:
        if _match_keywords(col, skip_kws):
            continue
        hits = _phone_hits(col)
        if hits < 2:
            continue
        sample_size = max(1, len(df[col].dropna().head(30)))
        content_score = int((hits / sample_size) * 60)
        name_bonus = 40 if _match_keywords(col, kws) else 0
        candidates.append((col, content_score + name_bonus))

    if not candidates:
        return None
    return max(candidates, key=lambda x: x[1])[0]


def detect_tag_col(df: pd.DataFrame) -> Optional[str]:
    kws = ["tag", "etiqueta", "label", "categoria", "tipo", "fonte", "origem"]
    for col in df.columns:
        if _match_keywords(col, kws):
            return col
    return None


def detect_value_col(df: pd.DataFrame) -> Optional[str]:
    """
    Detecta coluna de valor monetário por análise estatística — sem depender de keywords.
    Uma coluna de preço tem: separadores decimais, valores não-zero, faixa razoável, variância.
    Keywords de nome dão um bônus pequeno mas não são obrigatórias.
    """
    # Colunas que NUNCA são preço — só por nome óbvio
    hard_skip = ["cpf", "cnpj", "rg", "tel", "cel", "fone", "phone",
                 "whatsapp", "wpp", "percentual", "percent", "taxa", "mrg",
                 "margem", "desconto", "devolucao", "devolução"]

    # Keywords de nome para bônus (não obrigatório)
    bonus_kws = ["valor", "value", "total", "preco", "preço", "receita",
                 "ticket", "fatura", "price", "vl", "vlr", "bruto", "liquido",
                 "líquido", "sale", "compra", "amount"]

    def _parse_price(v):
        """Retorna float se o valor parece preço, None caso contrário."""
        s = re.sub(r"[R$\s]", "", str(v).strip())
        if not s or not re.search(r"\d", s):
            return None
        if "," in s:
            try:
                val = float(s.replace(".", "").replace(",", "."))
                return val if val < 10_000_000 else None
            except ValueError:
                return None
        # Sem decimal: só aceita se < 8 dígitos (afasta CPF/tel/CNPJ)
        pure = re.sub(r"\D", "", s)
        if 0 < len(pure) < 8:
            try:
                return float(pure)
            except ValueError:
                return None
        return None

    candidates: list = []
    for col in df.columns:
        col_lower = col.lower()
        if any(kw in col_lower for kw in hard_skip) or _is_doc_col(col):
            continue

        sample = df[col].dropna().head(50)
        if len(sample) < 3:
            continue

        parsed = [_parse_price(v) for v in sample]
        vals = [v for v in parsed if v is not None]
        if len(vals) < max(2, len(sample) * 0.35):
            continue

        n_nonzero = sum(1 for v in vals if v > 0)
        n_decimal = sum(1 for v in sample if "," in str(v))
        mean_val = sum(vals) / len(vals)

        # Coeficiente de variação: preços reais têm variância razoável
        # (não são todos iguais como "01", "01", "01")
        if mean_val > 0:
            std_val = (sum((v - mean_val) ** 2 for v in vals) / len(vals)) ** 0.5
            cv = std_val / mean_val
        else:
            cv = 0

        # Pontuação estatística — sem depender de nome
        score = 0
        score += int((n_decimal / len(sample)) * 35)          # decimal → forte sinal
        score += int((n_nonzero / max(len(vals), 1)) * 30)    # não-zero → valor real
        score += 20 if 1 <= mean_val <= 500_000 else 0        # faixa de preço razoável
        score += 10 if 0.05 <= cv <= 8.0 else 0               # variância adequada
        score -= 15 if cv < 0.01 and len(vals) > 3 else 0     # penaliza coluna uniforme

        # Bônus por nome (pequeno — não é obrigatório)
        if any(kw in col_lower for kw in bonus_kws):
            score += 15

        if score > 10:
            candidates.append((col, score))

    if not candidates:
        return None
    return max(candidates, key=lambda x: x[1])[0]


def parse_value(v) -> float:
    """Converte string de valor monetário para float."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return 0.0
    s = str(v).strip()
    s = re.sub(r"[R$\s]", "", s)
    # Formato brasileiro: 1.234,56 → 1234.56
    if re.match(r"^\d{1,3}(\.\d{3})*(,\d+)?$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


# ── Carregamento de arquivo ────────────────────────────────────────────────────

class _FileLike:
    """Wrapper de bytes para simular um arquivo uploadado (necessário para cache)."""
    def __init__(self, data: bytes, name: str):
        self._buf = io.BytesIO(data)
        self.name = name

    def __getattr__(self, attr):
        """Delega QUALQUER método não definido aqui para o BytesIO interno.
        Isso cobre tell, seekable, readable, getvalue, etc. sem precisar
        declarar cada um individualmente."""
        return getattr(self._buf, attr)

    def read(self, *a): return self._buf.read(*a)
    def seek(self, *a): return self._buf.seek(*a)
    def tell(self, *a): return self._buf.tell(*a)


@st.cache_data(show_spinner=False)
def _load_cached(file_bytes: bytes, file_name: str, sheets: tuple):
    """Versão cacheada do load_file_multisheet — evita re-parsear no mesmo arquivo."""
    return load_file_multisheet(_FileLike(file_bytes, file_name), list(sheets))


@st.cache_data(show_spinner=False)
def _get_sheets_cached(file_bytes: bytes, file_name: str) -> list:
    return get_excel_sheets(_FileLike(file_bytes, file_name))


def get_excel_sheets(uploaded) -> list:
    """Retorna lista de abas de um Excel. Lista vazia indica CSV."""
    name = uploaded.name.lower()
    if not name.endswith((".xlsx", ".xls", ".xlsm")):
        return []
    try:
        return _open_excel_file(uploaded).sheet_names
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


def _open_excel_file(uploaded):
    """Tenta abrir um ExcelFile com múltiplos engines (openpyxl, xlrd)."""
    for engine in (None, "openpyxl", "xlrd"):
        try:
            uploaded.seek(0)
            if engine:
                return pd.ExcelFile(uploaded, engine=engine)
            return pd.ExcelFile(uploaded)
        except Exception:
            continue
    raise ValueError("Não foi possível abrir o arquivo Excel. Verifique se o arquivo não está corrompido.")


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


def _score_sheet(df: pd.DataFrame) -> int:
    """
    Pontua uma aba como planilha de vendas (0-100).
    Abas primárias têm: data, valor, telefone e poucos 'Unnamed'.
    """
    score = 0
    if detect_date_col(df) is not None:
        score += 35   # data é o sinal mais forte
    if detect_value_col(df) is not None:
        score += 30
    if detect_phone_col(df) is not None:
        score += 20
    unnamed = sum(1 for c in df.columns if str(c).lower().startswith("unnamed"))
    unnamed_ratio = unnamed / max(len(df.columns), 1)
    if unnamed_ratio < 0.2:
        score += 15
    return score


def _normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Remove espaços e resolve colunas duplicadas."""
    cols = [str(c).strip() for c in df.columns]
    seen: dict = {}
    clean = []
    for c in cols:
        if c in seen:
            seen[c] += 1
            clean.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            clean.append(c)
    df = df.copy()
    df.columns = clean
    return df


def load_file_multisheet(
    uploaded, selected_sheets: list
) -> tuple[Optional[pd.DataFrame], dict]:
    """
    Carrega uma ou mais abas inteligentemente.
    Quando há múltiplas abas, pontua cada uma como planilha de vendas.
    Abas primárias (score ≥ 50) vêm primeiro no DataFrame combinado — isso
    garante que o lookup de telefone prefira linhas com data e valor.
    CSV: selected_sheets ignorado.
    """
    name = uploaded.name.lower()
    sheet_info: dict = {}
    scored_dfs: list = []   # (score, sheet_name, df, hrow)

    try:
        if name.endswith(".csv"):
            df, hrow = _load_csv_smart(uploaded)
            if len(df) > 0:
                sheet_info["CSV"] = hrow
                scored_dfs.append((100, "CSV", df, hrow))
        elif name.endswith((".xlsx", ".xls", ".xlsm")):
            xls = _open_excel_file(uploaded)
            for sheet in selected_sheets:
                df, hrow = _load_single_sheet(xls, sheet)
                if len(df) > 0:
                    score = _score_sheet(df)
                    sheet_info[sheet] = hrow
                    scored_dfs.append((score, sheet, df, hrow))
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {e}")
        return None, {}

    if not scored_dfs:
        return None, {}

    # Ordena: abas primárias (maior score) primeiro
    # Isso faz com que o lookup de telefone prefira linhas com data+valor
    scored_dfs.sort(key=lambda x: x[0], reverse=True)

    dfs = []
    primary_sheets = []
    secondary_sheets = []
    for score, sheet_name, df, hrow in scored_dfs:
        if len(scored_dfs) > 1:
            df = df.copy()
            df.insert(0, "_Planilha", sheet_name)
            df.insert(1, "_Score_Aba", score)
        if score >= 50:
            primary_sheets.append(sheet_name)
        else:
            secondary_sheets.append(sheet_name)
        dfs.append(df)

    combined = _normalize_cols(pd.concat(dfs, ignore_index=True))

    # Aviso sobre abas secundárias (para uso na UI)
    if secondary_sheets:
        combined.attrs["secondary_sheets"] = secondary_sheets
        combined.attrs["primary_sheets"] = primary_sheets

    return combined, sheet_info


# ── Lógica principal do PROCV ──────────────────────────────────────────────────

def _is_phone_col(col_data: pd.Series) -> bool:
    """Retorna True se a coluna tem pelo menos 2 valores com 8+ dígitos após limpeza."""
    if len(col_data) == 0:
        return False
    phone_like = int(col_data.apply(lambda v: len(clean_phone(str(v))) >= 8).sum())
    return phone_like >= 2


_DOC_COLS = {"cpf", "cnpj", "rg", "documento", "doc", "inscricao", "inscrição",
             "cadastro_pessoa", "cpf_cnpj", "identificacao", "identificação"}

def _is_doc_col(col_name: str) -> bool:
    """Retorna True se o nome da coluna sugere CPF/CNPJ/RG — evita falsos matches."""
    c = col_name.lower()
    return any(kw in c for kw in _DOC_COLS)


def _build_extended_lookup(df: pd.DataFrame, ds_with_meta: pd.DataFrame) -> dict:
    """
    Constrói {right8_key: row_dict} varrendo TODAS as colunas de telefone de df.
    Pula colunas de documentos. Linhas de abas primárias (_Score_Aba >= 50)
    têm prioridade — como o df já está ordenado por score (maior primeiro),
    o primeiro match por chave sempre será da aba melhor.
    """
    lookup: dict = {}
    for col in df.columns:
        if _is_doc_col(col) or col.startswith("_"):
            continue
        col_data = df[col].dropna()
        if not _is_phone_col(col_data):
            continue
        for idx in col_data.index:
            cleaned = clean_phone(str(df.at[idx, col]))
            if len(cleaned) >= 8:
                key = right8(cleaned)
                if key and key not in lookup:
                    lookup[key] = ds_with_meta.loc[idx].to_dict()
    return lookup


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
    Varre todas as colunas de telefone para maximizar matches.
    """
    ds = df_sales.copy()
    pos = ds.columns.get_loc(sales_phone_col) + 1
    ds.insert(pos, "Tel_Limpo_Vendas", ds[sales_phone_col].apply(clean_phone))
    ds.insert(pos + 1, "Tel_8dig_Vendas", ds["Tel_Limpo_Vendas"].apply(right8))

    dk = df_kommo.copy()
    pos_k = dk.columns.get_loc(kommo_phone_col) + 1
    dk.insert(pos_k, "Tel_Limpo_Kommo", dk[kommo_phone_col].apply(clean_phone))
    dk.insert(pos_k + 1, "Tel_8dig_Kommo", dk["Tel_Limpo_Kommo"].apply(right8))

    extended_lookup = _build_extended_lookup(df_sales, ds)
    lookup_keys = set(extended_lookup.keys())

    # Pré-computa right8 de cada coluna do Kommo (vetorizado) e monta sets por linha
    kommo_row_keys: List[set] = [set() for _ in range(len(dk))]
    for col in df_kommo.columns:
        r8_col = dk[col].fillna("").astype(str).apply(
            lambda v: right8(clean_phone(v)) if v else ""
        )
        for i, k in enumerate(r8_col):
            if k:
                kommo_row_keys[i].add(k)

    result_rows = []
    for i, (_, kr) in enumerate(dk.iterrows()):
        k8_primary = kr.get("Tel_8dig_Kommo", "")
        tag_raw = "" if pd.isna(kr.get(kommo_tag_col, np.nan)) else str(kr[kommo_tag_col])
        is_traffic = traffic_keyword.lower() in tag_raw.lower()

        sales_match = extended_lookup.get(k8_primary)
        if not sales_match:
            matched_key = next(iter(kommo_row_keys[i] & lookup_keys), None)
            if matched_key:
                sales_match = extended_lookup.get(matched_key)

        row_out = {
            "Tag_Kommo": tag_raw,
            "Telefone_Kommo": kr.get(kommo_phone_col, ""),
            "Tel_8dig": k8_primary,
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
    # Bug 2: deduplica por comprador único — mesmo telefone em vários leads Kommo
    # não deve contar como múltiplas conversões
    if len(df_trafego) > 0:
        df_trafego = df_trafego.drop_duplicates(subset=["Tel_8dig"]).reset_index(drop=True)

    return ds, dk, df_trafego, df_full


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


def _is_real_datetime(v) -> bool:
    """Retorna True apenas para datetime real — exclui pd.NaT que herda de datetime."""
    if v is None or v is pd.NaT:
        return False
    try:
        v.strftime("%Y")
        return True
    except Exception:
        return False


def parse_date(value) -> Optional[datetime]:
    if value is None or value is pd.NaT:
        return None
    if isinstance(value, datetime) and _is_real_datetime(value):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime() if _is_real_datetime(value) else None

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

    # Formato M/D/YY ou M/D/YYYY (americano, comum em exports de sistemas BR)
    # Ex: "4/30/26 18:05" → April 30, 2026
    m_us = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})(?:\s+\d{1,2}:\d{2})?$", s)
    if m_us:
        mo, dy, yr = int(m_us.group(1)), int(m_us.group(2)), int(m_us.group(3))
        if yr < 100:
            yr += 2000
        # Se dia > 12, certamente é MM/DD; se mês > 12, é DD/MM
        try:
            if mo <= 12 and dy <= 31:
                return datetime(yr, mo, dy)
        except ValueError:
            pass

    # Formatos padrão via pandas — tenta dayfirst=False primeiro para M/D/YY
    for dayfirst in (False, True):
        try:
            dt = pd.to_datetime(s, dayfirst=dayfirst)
            if _is_real_datetime(dt):
                return dt.to_pydatetime()
        except Exception:
            continue

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

    # Extrai data embutida em texto livre (ex: "Disparo dia das Mães 22/04" ou "23/04/2026 17:42")
    dm = re.search(r"\b(\d{1,2})[/\-\.](\d{1,2})(?:[/\-\.](\d{2,4}))?\b", s)
    if dm:
        try:
            day, month = int(dm.group(1)), int(dm.group(2))
            if 1 <= day <= 31 and 1 <= month <= 12:
                if dm.group(3):
                    year = int(dm.group(3))
                    if year < 100:
                        year += 2000
                else:
                    # Sem ano: infere pelo mês atual
                    # Se o mês extraído > mês atual, é do ano anterior
                    now = datetime.now()
                    year = now.year if month <= now.month else now.year - 1
                return datetime(year, month, day)
        except (ValueError, TypeError):
            pass

    return None


def detect_date_col(df: pd.DataFrame) -> Optional[str]:
    """
    Detecta coluna de data por conteúdo primeiro, nome como bônus.
    Funciona com colunas chamadas "Data", "Dt", "Date", "Período", etc.
    Cada planilha é um caso — não assume padronização.
    """
    kws = [
        "data", "date", "dt", "dia", "quando", "periodo", "período",
        "competencia", "competência", "venda_em", "criado", "created",
        "hora", "timestamp", "tempo", "time", "mes", "mês", "ano",
        "year", "month", "compra", "pedido", "emissao", "emissão",
        "lancamento", "lançamento", "movimento", "referencia", "referência",
    ]
    # Colunas que provavelmente NÃO são datas (evita falsos positivos)
    skip_kws = ["telefone", "celular", "phone", "tel", "whatsapp", "wpp",
                "valor", "preco", "preço", "total", "cpf", "cnpj", "cep"]

    candidates: list = []
    for col in df.columns:
        if _match_keywords(col, skip_kws):
            continue
        sample = df[col].dropna().head(25)
        if len(sample) == 0:
            continue
        hits = sum(1 for v in sample if parse_date(v) is not None)
        if hits < 2:
            continue
        content_score = int((hits / len(sample)) * 60)  # 0-60
        name_bonus = 40 if _match_keywords(col, kws) else 0
        candidates.append((col, content_score + name_bonus))

    if not candidates:
        return None
    # Retorna a coluna com maior pontuação combinada (conteúdo + nome)
    return max(candidates, key=lambda x: x[1])[0]


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

    ds = df_sales.copy()
    ds["_tel8"] = ds[sales_phone_col].apply(
        lambda v: right8(clean_phone(str(v))) if pd.notna(v) else ""
    )
    if sales_date_col:
        ds["_dt_venda"] = ds[sales_date_col].apply(parse_date)

    # Lookup estendido: todas as colunas de telefone das vendas
    # added_sale_idx: evita Bug 4 — mesma linha sob chaves diferentes (2 colunas de tel)
    added_sale_keys: set = set()  # (key, idx) para evitar duplicar por key
    sales_lookup: dict = {}
    for col in df_sales.columns:
        if _is_doc_col(col):
            continue
        col_data = df_sales[col].dropna()
        if not _is_phone_col(col_data):
            continue
        for idx in col_data.index:
            cleaned = clean_phone(str(df_sales.at[idx, col]))
            if len(cleaned) >= 8:
                key = right8(cleaned)
                if key and (key, idx) not in added_sale_keys:
                    sales_lookup.setdefault(key, []).append(ds.loc[idx].to_dict())
                    added_sale_keys.add((key, idx))

    result_rows = []
    for _, kr in df_disp_leads.iterrows():
        tel_raw = kr.get(kommo_phone_col, "")
        tel8_primary = right8(clean_phone(str(tel_raw))) if pd.notna(tel_raw) else ""

        # Coleta todas as chaves right8 deste lead (todas as colunas)
        all_tel8: list = []
        if tel8_primary:
            all_tel8.append(tel8_primary)
        for col in df_kommo.columns:
            v = kr.get(col, "")
            if pd.isna(v):
                continue
            cleaned = clean_phone(str(v))
            if len(cleaned) >= 8:
                k8 = right8(cleaned)
                if k8 and k8 not in all_tel8:
                    all_tel8.append(k8)

        if not all_tel8:
            continue

        # Tenta obter a data do disparo de múltiplas fontes (inteligência)
        disp_date = parse_date(kr.get(kommo_date_col)) if kommo_date_col else None
        if not _is_real_datetime(disp_date):
            # Fallback: extrai data embutida no texto da tag ("Disparo dia das Mães 22/04")
            disp_date = parse_date(str(kr.get(kommo_tag_col, "")))

        # Dois níveis de match:
        # 1. confirmed: telefone bate + data da venda DENTRO da janela (0..30 dias após disparo)
        # 2. phone_only: telefone bate + data da venda não parseável (inclui para revisão manual)
        # Vendas com data ANTES do disparo são explicitamente rejeitadas.
        confirmed_sale = None
        phone_only_sale = None

        for tel8 in all_tel8:
            if confirmed_sale:
                break
            for sale in sales_lookup.get(tel8, []):
                if _is_real_datetime(disp_date) and sales_date_col:
                    sale_dt = sale.get("_dt_venda")
                    if _is_real_datetime(sale_dt):
                        delta = (sale_dt - disp_date).days
                        if 0 <= delta <= window_days:
                            confirmed_sale = sale
                            break
                        # delta < 0 = venda antes do disparo → rejeita explicitamente
                    elif phone_only_sale is None:
                        # data não parseável → guarda como fallback
                        phone_only_sale = sale
                else:
                    # sem data de disparo ou sem coluna de data → match por telefone
                    confirmed_sale = sale
                    break

        matched_sale = confirmed_sale if confirmed_sale else phone_only_sale

        tag_raw = str(kr.get(kommo_tag_col, ""))
        row_out: dict = {
            "Tag_Kommo": tag_raw,
            "Telefone_Disparo": tel_raw,
            "Tel_8dig": tel8_primary,
        }

        if kommo_date_col:
            row_out["Data_Disparo"] = kr.get(kommo_date_col, "")

        if matched_sale:
            row_out["Venda_Confirmada"] = "SIM"
            if sales_date_col:
                sale_dt = matched_sale.get("_dt_venda")
                if _is_real_datetime(sale_dt):
                    row_out["Data_Venda"] = sale_dt.strftime("%d/%m/%Y")
                    if _is_real_datetime(disp_date):
                        row_out["Dias_Após_Disparo"] = (sale_dt - disp_date).days
                else:
                    row_out["Data_Venda"] = str(matched_sale.get(sales_date_col, ""))
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

    if not result_rows:
        return pd.DataFrame()
    df_res = pd.DataFrame(result_rows)
    # Deduplica por Tel_8dig: mesmo comprador em múltiplos leads Kommo não conta N vezes.
    # Prioriza Venda_Confirmada=SIM sobre NÃO antes de desduplicar.
    if "Tel_8dig" in df_res.columns:
        df_res = (df_res
                  .sort_values("Venda_Confirmada", ascending=True)  # NÃO < SIM
                  .drop_duplicates(subset=["Tel_8dig"])
                  .reset_index(drop=True))
    return df_res


# ── Análise de Duplicatas ──────────────────────────────────────────────────────

def _is_id_col(col_data: pd.Series, col_name: str) -> bool:
    """Detecta colunas de ID sequencial que não devem entrar no fingerprint."""
    id_kws = ["id", "codigo", "código", "protocolo", "numero", "número",
              "registro", "pedido", "ordem", "seq", "nro", "nº", "num",
              "index", "índice", "indice", "linha", "row"]
    if any(kw in col_name.lower() for kw in id_kws):
        return True
    non_null = col_data.dropna()
    if len(non_null) < 4:
        return False
    if non_null.nunique() == len(non_null):
        is_int = non_null.apply(
            lambda v: str(v).strip().replace(".0", "").lstrip("0").isdigit()
        ).sum()
        if int(is_int) >= 0.8 * len(non_null):
            return True
    return False


def analyze_duplicates(
    df: pd.DataFrame,
    phone_col: str,
    date_col: Optional[str] = None,
) -> pd.DataFrame:
    """
    Classifica registros com telefone repetido.
    Adiciona coluna 'Situacao_Venda':
      Única               → só uma venda para este número
      Multi-compra        → mesmo número, datas diferentes (compras reais)
      Multi-compra (mesmo dia) → mesmo número, mesma data, conteúdo diferente
      Multi-compra (sem data)  → mesmo número, sem data, conteúdo diferente
      DUPLICATA           → mesmo número + mesma data + conteúdo idêntico
      DUPLICATA (sem data)→ mesmo número, sem data, conteúdo idêntico
    """
    result = df.copy()
    phones = df[phone_col].apply(
        lambda v: right8(clean_phone(str(v))) if pd.notna(v) else ""
    )

    situacao = pd.Series(["Única"] * len(df), index=df.index, dtype=str)

    counts = phones.value_counts()
    multi_phones = counts[counts > 1].index

    id_cols = {col for col in df.columns if _is_id_col(df[col], col)}
    skip_set = {phone_col} | id_cols
    if date_col:
        skip_set.add(date_col)
    fp_cols = [c for c in df.columns if c not in skip_set]

    def _fingerprint(row) -> str:
        parts = []
        for c in fp_cols:
            v = row.get(c, "")
            if pd.isna(v):
                parts.append("")
                continue
            s = str(v).strip().lower()
            try:
                s = str(round(float(s.replace(",", ".").replace(" ", "")), 2))
            except ValueError:
                pass
            parts.append(s)
        return "|".join(parts)

    for phone in multi_phones:
        mask = phones == phone
        group = df[mask]

        if date_col and date_col in df.columns:
            date_keys = group[date_col].apply(
                lambda v: (parse_date(v).strftime("%Y-%m-%d") if parse_date(v) else "")
            )
            unique_dates = set(d for d in date_keys if d)
        else:
            date_keys = pd.Series([""] * len(group), index=group.index)
            unique_dates = set()

        if len(unique_dates) > 1:
            situacao[mask] = "Multi-compra"
        elif len(unique_dates) == 1:
            fps = group.apply(_fingerprint, axis=1)
            if fps.nunique() == 1:
                situacao[mask] = "DUPLICATA"
            else:
                situacao[mask] = "Multi-compra (mesmo dia)"
        else:
            fps = group.apply(_fingerprint, axis=1)
            if fps.nunique() == 1:
                situacao[mask] = "DUPLICATA (sem data)"
            else:
                situacao[mask] = "Multi-compra (sem data)"

    result.insert(0, "Situacao_Venda", situacao)
    return result


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
            elif col_name == "Situacao_Venda":
                if str(value).startswith("DUPLICATA"):
                    cell.fill = PatternFill("solid", fgColor="C0392B")
                    cell.font = Font(color="FFFFFF", bold=True, size=10)
                elif "mesmo dia" in str(value) or "sem data" in str(value):
                    cell.fill = PatternFill("solid", fgColor="E67E22")
                    cell.font = Font(color="FFFFFF", bold=True, size=10)
                elif str(value).startswith("Multi"):
                    cell.fill = PatternFill("solid", fgColor="2471A3")
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


def _rename_result_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Renomeia colunas técnicas para linguagem legível no Excel."""
    return df.rename(columns={
        "Venda_Confirmada": "Comprou?",
        "É_Tráfego": "É Lead de Tráfego?",
        "Tag_Kommo": "Tag no Kommo",
        "Telefone_Kommo": "Telefone (Kommo)",
        "Telefone_Disparo": "Telefone (Disparo)",
        "Tel_8dig": "Tel. (8 dígitos)",
        "Tel_Limpo_Vendas": "Tel. Limpo (Vendas)",
        "Tel_8dig_Vendas": "Tel. 8 dígitos (Vendas)",
        "Tel_Limpo_Kommo": "Tel. Limpo (Kommo)",
        "Tel_8dig_Kommo": "Tel. 8 dígitos (Kommo)",
        "Data_Disparo": "Data do Disparo",
        "Data_Venda": "Data da Venda",
        "Dias_Após_Disparo": "Dias após o Disparo",
        "Situacao_Venda": "Situação do Registro",
        "Origem": "Origem da Venda",
    })


def _write_guide_sheet(ws, summary_data: dict):
    """Escreve a aba 'Como Ler' com explicações claras para o usuário final."""
    PURPLE = "7B2FBE"
    DARK   = "1C1C1E"
    WHITE  = "FFFFFF"
    GRAY   = "F5F5F7"
    GREEN  = "27AE60"
    ORANGE = "E67E22"
    RED    = "C0392B"
    BLUE   = "2471A3"

    def _title(row, text, hex_bg=PURPLE):
        ws.merge_cells(f"A{row}:D{row}")
        c = ws.cell(row, 1, text)
        c.font = Font(color=WHITE, bold=True, size=13)
        c.fill = PatternFill("solid", fgColor=hex_bg)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 28

    def _row(row, label, desc, label_bold=False):
        ws.merge_cells(f"B{row}:D{row}")
        a = ws.cell(row, 1, label)
        a.font = Font(bold=label_bold, size=10)
        a.fill = PatternFill("solid", fgColor=GRAY)
        a.alignment = Alignment(vertical="center", indent=1)
        b = ws.cell(row, 2, desc)
        b.font = Font(size=10)
        b.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[row].height = 20

    def _color_legend(row, color_hex, label, meaning):
        ws.cell(row, 1, "  ██").font = Font(color=color_hex, size=14, bold=True)
        ws.cell(row, 1).fill = PatternFill("solid", fgColor=GRAY)
        ws.merge_cells(f"B{row}:C{row}")
        ws.cell(row, 2, label).font = Font(bold=True, size=10)
        ws.cell(row, 4, meaning).font = Font(size=10)
        ws.row_dimensions[row].height = 18

    r = 1
    # ── Cabeçalho ─────────────────────────────────────────────────────
    ws.merge_cells(f"A{r}:D{r}")
    hdr = ws.cell(r, 1, "📋  PROC-BULLS — GUIA DE LEITURA")
    hdr.font = Font(color=WHITE, bold=True, size=16)
    hdr.fill = PatternFill("solid", fgColor=PURPLE)
    hdr.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 40
    r += 1

    ws.merge_cells(f"A{r}:D{r}")
    sub = ws.cell(r, 1, "Desenvolvido por João  ·  Aure Digital  ·  Proc-Bulls")
    sub.font = Font(italic=True, size=10, color="888888")
    sub.alignment = Alignment(horizontal="center")
    ws.row_dimensions[r].height = 18
    r += 2

    # ── O que é este relatório ────────────────────────────────────────
    _title(r, "O QUE É ESTE RELATÓRIO?", DARK); r += 1
    ws.merge_cells(f"A{r}:D{r}")
    ws.cell(r, 1, (
        "Este relatório cruza sua planilha de vendas com o Kommo CRM para identificar "
        "quantas vendas vieram do tráfego pago e/ou de campanhas de disparo (WhatsApp)."
    )).font = Font(size=10)
    ws.cell(r, 1).alignment = Alignment(wrap_text=True, indent=1)
    ws.row_dimensions[r].height = 30
    r += 2

    # ── Abas do relatório ─────────────────────────────────────────────
    _title(r, "ABAS DESTE ARQUIVO — O QUE CADA UMA CONTÉM", DARK); r += 1
    sheets_info = [
        ("📋 Como Ler",              "Esta aba. Leia primeiro."),
        ("📊 Resumo",                "Números principais: conversões, receita, taxas. Comece aqui."),
        ("✅ Vendas — Tráfego",      "Leads que vieram do tráfego pago E fizeram uma compra."),
        ("📣 Vendas — Disparo",      "Leads que receberam disparo E fizeram uma compra dentro de 30 dias."),
        ("🔍 Duplicatas",            "Registros duplicados na planilha de vendas (se houver)."),
        ("📋 Todos os Leads",        "Todos os leads do Kommo, com coluna 'Comprou?' para filtrar."),
        ("📦 Vendas (dados brutos)", "Sua planilha de vendas original, processada."),
        ("🗂️ Kommo (dados brutos)",  "Sua planilha do Kommo, processada."),
    ]
    for name, desc in sheets_info:
        _row(r, name, desc, label_bold=True); r += 1
    r += 1

    # ── Colunas importantes ───────────────────────────────────────────
    _title(r, "COLUNAS IMPORTANTES — O QUE CADA UMA SIGNIFICA", DARK); r += 1
    cols_info = [
        ("Comprou?",             "SIM = lead fez uma compra | NÃO = não encontramos venda para este lead"),
        ("É Lead de Tráfego?",   "SIM = lead tem a tag de tráfego pago no Kommo"),
        ("Origem da Venda",      "Tráfego / Disparo / Tráfego + Disparo (foi impactado pelos dois)"),
        ("Tel. (8 dígitos)",     "Últimos 8 dígitos do telefone — usados para comparar as listas"),
        ("Dias após o Disparo",  "Quantos dias após o disparo a venda aconteceu (máx. 30 dias conta)"),
        ("Situação do Registro", "Única / Multi-compra / DUPLICATA — análise de repetição nas vendas"),
    ]
    for col, desc in cols_info:
        _row(r, col, desc); r += 1
    r += 1

    # ── Legenda de cores ──────────────────────────────────────────────
    _title(r, "LEGENDA DE CORES NAS TABELAS", DARK); r += 1
    colors = [
        (GREEN,  "Verde",   "Venda confirmada (Comprou? = SIM)"),
        (RED,    "Vermelho","Não comprou (Comprou? = NÃO) | Duplicata"),
        (BLUE,   "Azul",    "Lead de tráfego pago | Multi-compra"),
        (ORANGE, "Laranja", "Atenção: mesmo dia / sem data confirmada"),
    ]
    for hex_c, name, meaning in colors:
        _color_legend(r, hex_c, name, meaning); r += 1
    r += 1

    # ── Como interpretar ──────────────────────────────────────────────
    _title(r, "COMO INTERPRETAR OS RESULTADOS", DARK); r += 1
    tips = [
        ("Tráfego puro",      "Lead veio de anúncio e comprou — conversão de tráfego."),
        ("Disparo puro",      "Lead recebeu WhatsApp e comprou em até 30 dias — conversão de disparo."),
        ("Tráfego + Disparo", "Lead veio de anúncio MAS só comprou depois do disparo. A equipe decide a quem atribuir."),
        ("Duplicata",         "Mesmo cliente, mesma data, mesmo valor — provável lançamento duplo no sistema. Verificar."),
        ("Multi-compra",      "Mesmo número, datas diferentes — cliente recorrente. Contar normalmente."),
    ]
    for label, desc in tips:
        _row(r, label, desc, label_bold=True); r += 1
    r += 1

    # ── Resumo de números ─────────────────────────────────────────────
    if summary_data:
        _title(r, "RESUMO RÁPIDO DOS NÚMEROS", "FF6B35"); r += 1
        for label, val in summary_data.items():
            _row(r, label, str(val)); r += 1

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 40


def _sum_result_vcol(df_res: Optional[pd.DataFrame], value_col: Optional[str]) -> float:
    """Soma [Venda] {value_col} no DataFrame de resultado. Retorna 0.0 se não disponível."""
    if not value_col or df_res is None or len(df_res) == 0:
        return 0.0
    vcol = f"[Venda] {value_col}"
    if vcol not in df_res.columns:
        return 0.0
    return float(df_res[vcol].apply(parse_value).sum())


def _brl(value) -> str:
    """Formata float como R$ brasileiro."""
    try:
        return f"R$ {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "—"


def _add_total_row(ws, df: pd.DataFrame, value_col_name: str = ""):
    """Adiciona linha de TOTAL somando APENAS a coluna de valor identificada."""
    if len(df) == 0 or not value_col_name or value_col_name not in df.columns:
        return
    last_data_row = len(df) + 2
    total_row = last_data_row + 1
    BOLD_FILL = PatternFill("solid", fgColor="1C1C1E")

    for c, col in enumerate(df.columns, 1):
        cell = ws.cell(total_row, c)
        cell.fill = BOLD_FILL
        if c == 1:
            cell.value = "TOTAL"
            cell.font = Font(color="FFFFFF", bold=True, size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col == value_col_name:
            try:
                total = df[col].apply(parse_value).sum()
                cell.value = _brl(total)
                cell.font = Font(bold=True, size=12, color="FF6B35")
                cell.alignment = Alignment(horizontal="center")
            except Exception:
                pass
    ws.row_dimensions[total_row].height = 24


def build_excel(
    ds: pd.DataFrame,
    dk: pd.DataFrame,
    df_result: pd.DataFrame,
    df_full: pd.DataFrame,
    df_disparo_result: Optional[pd.DataFrame] = None,
    df_dup_analysis: Optional[pd.DataFrame] = None,
    sales_value_col: Optional[str] = None,
) -> bytes:
    wb = Workbook()

    total_traffic = int((df_full["É_Tráfego"] == "SIM").sum()) if "É_Tráfego" in df_full.columns else 0
    total_sales   = len(ds)
    confirmed     = len(df_result)
    conv_rate     = f"{confirmed/total_traffic*100:.1f}%" if total_traffic > 0 else "—"

    # Prepara resumo para a aba guia
    guide_summary = {
        "Vendas carregadas": total_sales,
        "Leads no Kommo": len(dk),
        "Leads de tráfego": total_traffic,
        "Conversões de tráfego": confirmed,
        "Taxa de conversão (tráfego)": conv_rate,
    }
    if df_disparo_result is not None and len(df_disparo_result) > 0:
        d_conv = int((df_disparo_result["Venda_Confirmada"] == "SIM").sum())
        d_tot  = len(df_disparo_result)
        guide_summary["Conversões de disparo"] = d_conv
        guide_summary["Taxa de conversão (disparo)"] = f"{d_conv/d_tot*100:.1f}%" if d_tot > 0 else "—"

    # ── Aba 1: Como Ler (primeira aba — lida antes de tudo) ───────────
    ws_guide = wb.active
    ws_guide.title = "📋 Como Ler"
    _write_guide_sheet(ws_guide, guide_summary)

    # ── Aba 2: Resumo ─────────────────────────────────────────────────
    ws_res = wb.create_sheet("📊 Resumo")
    ws_res.merge_cells("A1:C1")
    t = ws_res.cell(1, 1, "RESUMO DO PROC-BULLS")
    t.font = Font(color="FFFFFF", bold=True, size=14)
    t.fill = PatternFill("solid", fgColor="7B2FBE")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws_res.row_dimensions[1].height = 36

    summary_rows = [
        ("Vendas carregadas", total_sales),
        ("Leads no Kommo", len(dk)),
        ("— TRÁFEGO —", ""),
        ("Leads com tag de tráfego", total_traffic),
        ("Conversões confirmadas", confirmed),
        ("Taxa de conversão", conv_rate),
        ("→ Veja os valores na aba ✅ Vendas — Tráfego", ""),
    ]
    if df_disparo_result is not None and len(df_disparo_result) > 0:
        d_tot  = len(df_disparo_result)
        d_conv = int((df_disparo_result["Venda_Confirmada"] == "SIM").sum())
        d_rate = f"{d_conv/d_tot*100:.1f}%" if d_tot > 0 else "—"
        summary_rows += [
            ("— DISPARO —", ""),
            ("Leads de disparo analisados", d_tot),
            ("Conversões confirmadas", d_conv),
            ("Taxa de conversão", d_rate),
            ("→ Veja os valores na aba 📣 Vendas — Disparo", ""),
        ]
    if df_dup_analysis is not None and "Situacao_Venda" in df_dup_analysis.columns:
        n_dup   = int((df_dup_analysis["Situacao_Venda"].str.startswith("DUPLICATA")).sum())
        n_multi = int((df_dup_analysis["Situacao_Venda"].str.startswith("Multi")).sum())
        summary_rows += [
            ("— QUALIDADE DOS DADOS —", ""),
            ("Duplicatas suspeitas", n_dup),
            ("Multi-compras (cliente recorrente)", n_multi),
        ]
    for i, (label, val) in enumerate(summary_rows, 2):
        lc = ws_res.cell(i, 1, label)
        lc.font = Font(bold=True, size=11)
        if str(label).startswith("—"):
            lc.fill = PatternFill("solid", fgColor="7B2FBE")
            lc.font = Font(bold=True, size=11, color="FFFFFF")
        vc = ws_res.cell(i, 2, val)
        vc.font = Font(bold=True, size=13, color="FF6B35")
        vc.alignment = Alignment(horizontal="center")
        ws_res.row_dimensions[i].height = 26
    ws_res.column_dimensions["A"].width = 46
    ws_res.column_dimensions["B"].width = 22

    # Coluna de valor renomeada (para _add_total_row)
    _vcol = f"[Venda] {sales_value_col}" if sales_value_col else ""

    # ── Aba 3: Vendas — Tráfego ───────────────────────────────────────
    ws3 = wb.create_sheet("✅ Vendas — Tráfego")
    if len(df_result) == 0:
        ws3.cell(1, 1, "Nenhuma venda de tráfego confirmada.").font = Font(italic=True, color="888888")
    else:
        df_result_renamed = _rename_result_cols(df_result)
        _write_sheet(ws3, df_result_renamed,
                     f"VENDAS DO TRÁFEGO PAGO — {confirmed} CONVERSÕES CONFIRMADAS", "27AE60")
        _add_total_row(ws3, df_result_renamed, _vcol)

    # ── Aba 4: Vendas — Disparo ───────────────────────────────────────
    all_sheets = [ws_guide, ws_res, ws3]
    if df_disparo_result is not None and len(df_disparo_result) > 0:
        ws_disp = wb.create_sheet("📣 Vendas — Disparo")
        d_conv  = int((df_disparo_result["Venda_Confirmada"] == "SIM").sum())
        df_disp_renamed = _rename_result_cols(df_disparo_result)
        _write_sheet(
            ws_disp, df_disp_renamed,
            f"VENDAS DO DISPARO — {d_conv} CONVERSÕES CONFIRMADAS",
            "8E44AD",
            ["Data do Disparo", "Data da Venda", "Dias após o Disparo"],
        )
        _add_total_row(ws_disp, df_disp_renamed, _vcol)
        all_sheets.append(ws_disp)

    # ── Aba 5: Duplicatas ─────────────────────────────────────────────
    if df_dup_analysis is not None and "Situacao_Venda" in df_dup_analysis.columns:
        df_non_unique = df_dup_analysis[df_dup_analysis["Situacao_Venda"] != "Única"]
        if len(df_non_unique) > 0:
            n_dup   = int((df_dup_analysis["Situacao_Venda"].str.startswith("DUPLICATA")).sum())
            n_multi = int((df_dup_analysis["Situacao_Venda"].str.startswith("Multi")).sum())
            ws_dup  = wb.create_sheet("🔍 Duplicatas")
            _write_sheet(
                ws_dup, _rename_result_cols(df_non_unique),
                f"DUPLICATAS ({n_dup}) E MULTI-COMPRAS ({n_multi}) — VERIFIQUE",
                "922B21", ["Situação do Registro"],
            )
            all_sheets.append(ws_dup)

    # ── Aba 6: Todos os Leads ─────────────────────────────────────────
    ws4 = wb.create_sheet("📋 Todos os Leads")
    _write_sheet(ws4, _rename_result_cols(df_full),
                 "TODOS OS LEADS DO KOMMO — FILTRE PELA COLUNA 'COMPROU?'", "6C3483")
    all_sheets.append(ws4)

    # ── Aba 7: Dados Brutos — Vendas ──────────────────────────────────
    ws1 = wb.create_sheet("📦 Vendas (bruto)")
    _write_sheet(ws1, _rename_result_cols(ds),
                 "PLANILHA DE VENDAS PROCESSADA", "FF6B35",
                 ["Tel. Limpo (Vendas)", "Tel. 8 dígitos (Vendas)"])
    all_sheets.append(ws1)

    # ── Aba 8: Dados Brutos — Kommo ───────────────────────────────────
    ws2 = wb.create_sheet("🗂️ Kommo (bruto)")
    _write_sheet(ws2, _rename_result_cols(dk),
                 "PLANILHA KOMMO PROCESSADA", "2980B9",
                 ["Tel. Limpo (Kommo)", "Tel. 8 dígitos (Kommo)"])
    all_sheets.append(ws2)

    # ── Rodapé em todas as abas ───────────────────────────────────────
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
        sales_bytes = sales_file.read()  # lê UMA vez aqui
        sales_sheets = _get_sheets_cached(sales_bytes, sales_file.name)
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
            df_sales_raw, sales_info = _load_cached(
                sales_bytes, sales_file.name, tuple(selected_sales_sheets or ["__csv__"])
            )
            if df_sales_raw is not None:
                for sheet, hrow in sales_info.items():
                    if hrow > 0:
                        st.info(f"📋 '{sheet}': cabeçalho detectado na linha {hrow + 1} — título(s) anteriores ignorados.")
                # Aviso sobre abas primárias vs secundárias
                sec = df_sales_raw.attrs.get("secondary_sheets", [])
                pri = df_sales_raw.attrs.get("primary_sheets", [])
                if sec:
                    st.warning(
                        f"🧠 **Análise inteligente de abas:** "
                        f"A aba **'{', '.join(sec)}'** não tem estrutura completa de vendas "
                        f"(sem coluna de data ou muitos campos sem nome). "
                        f"Ela será usada apenas para cruzamento de telefones. "
                        f"A análise principal (datas, valores, duplicatas) usa: **'{', '.join(pri)}'**."
                    )
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
        kommo_bytes = kommo_file.read()  # lê UMA vez aqui
        kommo_sheets = _get_sheets_cached(kommo_bytes, kommo_file.name)
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
            df_kommo_raw, kommo_info = _load_cached(
                kommo_bytes, kommo_file.name, tuple(selected_kommo_sheets or ["__csv__"])
            )
            if df_kommo_raw is not None:
                for sheet, hrow in kommo_info.items():
                    if hrow > 0:
                        st.info(f"📋 '{sheet}': cabeçalho detectado na linha {hrow + 1} — título(s) anteriores ignorados.")
                st.success(f"✅ {len(df_kommo_raw):,} linhas · {len(df_kommo_raw.columns)} colunas")
                with st.expander("Prévia"):
                    st.dataframe(df_kommo_raw.head(6), use_container_width=True)

st.markdown("#### 📣 Planilha do Kommo — Disparo *(opcional, se arquivo separado do tráfego)*")
kommo_disparo_file = st.file_uploader(
    "Se o Kommo de disparo é um arquivo diferente do de tráfego, faça upload aqui",
    type=["xlsx", "xls", "csv", "xlsm"],
    key="kommo_disparo_upload",
    help="Opcional. Se não carregar, o arquivo Kommo principal será usado para ambos tráfego e disparo.",
)
df_kommo_disparo_raw: Optional[pd.DataFrame] = None
if kommo_disparo_file:
    kd_bytes = kommo_disparo_file.read()  # lê UMA vez aqui
    kd_sheets = _get_sheets_cached(kd_bytes, kommo_disparo_file.name)
    selected_kd_sheets = kd_sheets
    if len(kd_sheets) > 1:
        selected_kd_sheets = st.multiselect(
            "Planilhas do Kommo Disparo:",
            options=kd_sheets, default=kd_sheets, key="kommo_disp_sheets",
        )
    if selected_kd_sheets or not kd_sheets:
        df_kommo_disparo_raw, kd_info = _load_cached(
            kd_bytes, kommo_disparo_file.name, tuple(selected_kd_sheets or ["__csv__"])
        )
        if df_kommo_disparo_raw is not None:
            st.success(f"✅ Kommo Disparo: {len(df_kommo_disparo_raw):,} linhas · {len(df_kommo_disparo_raw.columns)} colunas")
            with st.expander("Prévia"):
                st.dataframe(df_kommo_disparo_raw.head(6), use_container_width=True)

st.divider()

# ── Passo 2: Configuração ──────────────────────────────────────────────────────
if df_sales_raw is not None and df_kommo_raw is not None:
    st.markdown('<div class="step-wrap"><div class="step-num">2</div><div class="step-text">Configure as colunas</div></div>', unsafe_allow_html=True)

    auto_sp = detect_phone_col(df_sales_raw)
    auto_kp = detect_phone_col(df_kommo_raw)
    auto_kt = detect_tag_col(df_kommo_raw)
    auto_vl = detect_value_col(df_sales_raw)

    def _idx(df, col):
        cols = list(df.columns)
        return cols.index(col) if col in cols else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        sales_phone_col = st.selectbox(
            "Coluna de telefone — Vendas",
            list(df_sales_raw.columns),
            index=_idx(df_sales_raw, auto_sp),
            help="Coluna da planilha de vendas com o telefone/WhatsApp do cliente. "
                 "A ferramenta usa os últimos 8 dígitos para comparar — funciona mesmo com DDD, sem DDD, com ou sem o 9.",
        )
        if auto_sp == sales_phone_col:
            st.caption("✨ Auto-detectado")
    with c2:
        kommo_phone_col = st.selectbox(
            "Coluna de telefone — Kommo",
            list(df_kommo_raw.columns),
            index=_idx(df_kommo_raw, auto_kp),
            help="Coluna do Kommo com o telefone do lead (geralmente 'Celular'). "
                 "Mesmo que o número esteja em outra coluna, a ferramenta testa todas automaticamente.",
        )
        if auto_kp == kommo_phone_col:
            st.caption("✨ Auto-detectado")
    with c3:
        kommo_tag_col = st.selectbox(
            "Coluna de tags — Kommo",
            list(df_kommo_raw.columns),
            index=_idx(df_kommo_raw, auto_kt) if auto_kt else 0,
            help="Coluna do Kommo onde estão as etiquetas dos leads (geralmente 'Tags'). "
                 "É aqui que a ferramenta procura as palavras-chave de tráfego e de disparo.",
        )
        if auto_kt == kommo_tag_col:
            st.caption("✨ Auto-detectado")
    with c4:
        traffic_keyword = st.text_input(
            "Tag de tráfego — palavra-chave",
            value="trafego",
            placeholder="Cole aqui como está no Kommo...",
            help="Cole aqui um trecho da tag de tráfego exatamente como aparece no Kommo. "
                 "Não precisa ser a tag completa — um pedaço já resolve. "
                 "A busca ignora maiúsculas/minúsculas e acentos. "
                 "Exemplos: 'trafego', 'pago', 'ads', 'lead trafego'.",
        )

    # Coluna de valor monetário
    none_val = "(não usar)"
    val_opts = [none_val] + list(df_sales_raw.columns)
    val_default = val_opts.index(auto_vl) if auto_vl and auto_vl in val_opts else 0
    sales_value_col_sel = st.selectbox(
        "Coluna de valor da venda *(opcional)*",
        val_opts, index=val_default, key="sales_value_col",
        help="Se a planilha de vendas tiver uma coluna com o valor em R$ de cada venda, selecione aqui. "
             "A ferramenta vai calcular e exibir a receita total de tráfego e de disparo.",
    )
    sales_value_col = None if sales_value_col_sel == none_val else sales_value_col_sel
    if auto_vl == sales_value_col:
        st.caption("✨ Auto-detectado")

    # ── Configuração de Disparo ────────────────────────────────────────────────
    st.markdown("##### 📣 Disparo")
    none_opt = "(não usar)"

    # Decide qual Kommo usar para disparo
    df_kommo_disp = df_kommo_disparo_raw if df_kommo_disparo_raw is not None else df_kommo_raw
    if df_kommo_disparo_raw is not None:
        st.caption("📣 Usando o arquivo Kommo Disparo separado para análise de disparo.")

    auto_kd = detect_date_col(df_kommo_disp)
    auto_sd = detect_date_col(df_sales_raw)

    # Se há segundo Kommo, deixa o usuário escolher as colunas dele
    if df_kommo_disparo_raw is not None:
        auto_kdp = detect_phone_col(df_kommo_disparo_raw)
        auto_kdt = detect_tag_col(df_kommo_disparo_raw)
        dd1, dd2 = st.columns(2)
        with dd1:
            kommo_disp_phone_col = st.selectbox(
                "Telefone — Kommo Disparo",
                list(df_kommo_disparo_raw.columns),
                index=_idx(df_kommo_disparo_raw, auto_kdp),
                key="kommo_disp_phone",
            )
            if auto_kdp == kommo_disp_phone_col:
                st.caption("✨ Auto-detectado")
        with dd2:
            kommo_disp_tag_col = st.selectbox(
                "Coluna de tags — Kommo Disparo",
                list(df_kommo_disparo_raw.columns),
                index=_idx(df_kommo_disparo_raw, auto_kdt) if auto_kdt else 0,
                key="kommo_disp_tag",
            )
            if auto_kdt == kommo_disp_tag_col:
                st.caption("✨ Auto-detectado")
    else:
        kommo_disp_phone_col = kommo_phone_col
        kommo_disp_tag_col = kommo_tag_col

    e1, e2, e3 = st.columns(3)
    with e1:
        disparo_keyword = st.text_input(
            "Tag de disparo — palavra-chave",
            value="disparo",
            placeholder="Cole aqui como está a tag no Kommo...",
            help="Cole aqui um trecho da tag de disparo exatamente como aparece no Kommo. "
                 "Não precisa ser a tag completa — um pedaço já basta. "
                 "Exemplos: se no Kommo a tag for 'recebeu disparo dia das mães', "
                 "pode digitar apenas 'disparo' ou 'recebeu disparo' ou 'dia das mães'. "
                 "A busca é parcial e ignora maiúsculas/minúsculas.",
            key="disparo_kw",
        )
    with e2:
        kommo_date_opts = [none_opt] + list(df_kommo_disp.columns)
        kommo_date_default = (kommo_date_opts.index(auto_kd)
                              if auto_kd and auto_kd in kommo_date_opts else 0)
        kommo_date_sel = st.selectbox(
            "Data do disparo no Kommo *(opcional)*",
            kommo_date_opts,
            index=kommo_date_default,
            key="kommo_date",
            help="Selecione a coluna do Kommo que contém a data em que o disparo foi feito. "
                 "Se a data estiver embutida no texto da tag (ex: 'Disparo 22/04'), "
                 "a ferramenta extrai automaticamente — deixe em '(não usar)' nesse caso. "
                 "A data do disparo é usada para garantir que só vendas APÓS o disparo sejam contadas.",
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
            help="Coluna da planilha de vendas com a data em que a compra foi realizada. "
                 "Usada para validar que a venda aconteceu DEPOIS do disparo (janela de 30 dias). "
                 "Se não tiver essa coluna, a ferramenta ainda encontra as vendas — "
                 "mas não consegue garantir a ordem cronológica.",
        )
        sales_date_col = None if sales_date_sel == none_opt else sales_date_sel
        if auto_sd == sales_date_col:
            st.caption("✨ Auto-detectado")

    if not kommo_date_col:
        st.caption("ℹ️ Sem coluna de data do disparo: a ferramenta tentará extrair a data do texto da tag automaticamente (ex: 'Disparo 22/04').")

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
            progress.progress(10, text="Cruzando vendas com Kommo...")
            ds_t, dk_t, df_result, df_full = run_procv(
                df_sales_raw, sales_phone_col,
                df_kommo_raw, kommo_phone_col, kommo_tag_col,
                traffic_keyword,
            )

            # Compradores únicos (deduplicado por Tel_8dig no run_procv)
            confirmed = len(df_result)

            progress.progress(60, text="Analisando disparo...")
            df_disparo_result = None
            if considerar_disparo and disparo_keyword.strip():
                df_disparo_result = run_disparo(
                    df_sales_raw, sales_phone_col, sales_date_col,
                    df_kommo_disp, kommo_disp_phone_col, kommo_disp_tag_col,
                    disparo_keyword, kommo_date_col,
                )

            # ── Atribuição: sobreposição tráfego × disparo ────────────────────
            trafego_phones: set = set()
            if len(df_result) > 0 and "Tel_8dig" in df_result.columns:
                trafego_phones = set(df_result["Tel_8dig"].dropna().values)

            disparo_phones: set = set()
            if df_disparo_result is not None and len(df_disparo_result) > 0:
                disp_sim = df_disparo_result[df_disparo_result["Venda_Confirmada"] == "SIM"]
                if "Tel_8dig" in disp_sim.columns:
                    disparo_phones = set(disp_sim["Tel_8dig"].dropna().values)

            overlap_phones = trafego_phones & disparo_phones

            # Adiciona coluna "Origem" em cada resultado
            if len(df_result) > 0:
                df_result["Origem"] = df_result["Tel_8dig"].apply(
                    lambda t: "Tráfego + Disparo" if t in overlap_phones else "Tráfego"
                )
            if df_disparo_result is not None and len(df_disparo_result) > 0:
                df_disparo_result["Origem"] = df_disparo_result["Tel_8dig"].apply(
                    lambda t: "Tráfego + Disparo" if t in overlap_phones else "Disparo"
                )

            disp_sim_df = None
            if df_disparo_result is not None and len(df_disparo_result) > 0:
                disp_sim_df = df_disparo_result[df_disparo_result["Venda_Confirmada"] == "SIM"]

            progress.progress(80, text="Analisando duplicatas e multi-compras...")
            # Duplicatas: só em abas primárias (score ≥ 50) para evitar falsos positivos
            # de telefones que aparecem em abas de cadastro/auxiliares
            df_for_dup = df_sales_raw
            if "_Score_Aba" in df_sales_raw.columns:
                df_primary = df_sales_raw[
                    pd.to_numeric(df_sales_raw["_Score_Aba"], errors="coerce").fillna(0) >= 50
                ]
                if len(df_primary) >= 2:
                    df_for_dup = df_primary
            df_dup_analysis = analyze_duplicates(df_for_dup, sales_phone_col, sales_date_col)

            progress.progress(90, text="Gerando relatório Excel...")
            excel_bytes = build_excel(ds_t, dk_t, df_result, df_full, df_disparo_result, df_dup_analysis,
                                       sales_value_col=sales_value_col)
            st.session_state["excel_bytes"] = excel_bytes

            progress.progress(100, text="Concluído!")
            progress.empty()

            total_traffic = int((df_full["É_Tráfego"] == "SIM").sum())
            disp_conv  = 0
            disp_total = 0
            if df_disparo_result is not None and len(df_disparo_result) > 0:
                disp_conv  = int((df_disparo_result["Venda_Confirmada"] == "SIM").sum())
                disp_total = len(df_disparo_result)

            n_overlap    = len(overlap_phones)
            n_total_uniq = len(trafego_phones | disparo_phones)

            # Qualidade dos dados de vendas
            n_sales_total  = len(df_sales_raw)
            phones_series  = df_sales_raw[sales_phone_col].apply(
                lambda v: right8(clean_phone(str(v))) if pd.notna(v) else "")
            n_sem_tel = int((phones_series == "").sum())
            n_com_tel = n_sales_total - n_sem_tel

            # ── Resumo ─────────────────────────────────────────────────────────
            st.markdown('<div class="step-wrap"><div class="step-num">✓</div><div class="step-text">Resultados</div></div>', unsafe_allow_html=True)

            r1, r2, r3, r4 = st.columns(4)
            r1.metric("Vendas carregadas", f"{n_sales_total:,}")
            r2.metric("Com telefone válido", f"{n_com_tel:,}",
                      delta=f"{n_sem_tel:,} sem tel." if n_sem_tel > 0 else None,
                      delta_color="off")
            r3.metric("Leads no Kommo", f"{len(dk_t):,}")
            r4.metric("Leads de tráfego no Kommo", f"{total_traffic:,}")

            if n_sem_tel > 0:
                pct = n_sem_tel / n_sales_total * 100
                st.warning(
                    f"⚠️ **{n_sem_tel} de {n_sales_total} vendas ({pct:.0f}%) não têm telefone preenchido** "
                    f"e não podem ser cruzadas com o Kommo. Os resultados abaixo refletem apenas as {n_com_tel} vendas com telefone."
                )

            st.divider()

            # ── Tráfego ────────────────────────────────────────────────────────
            st.markdown('<div class="step-wrap"><div class="step-num">📊</div><div class="step-text">Tráfego Pago</div></div>', unsafe_allow_html=True)
            if confirmed > 0:
                taxa_t = f"{confirmed/total_traffic*100:.1f}%" if total_traffic > 0 else "—"
                st.success(f"**{confirmed} vendas** encontradas de tráfego pago — {taxa_t} de conversão sobre {total_traffic:,} leads.")
                st.caption("Veja a lista completa na aba **✅ Vendas — Tráfego** do Excel.")
                with st.expander(f"Prévia — {confirmed} vendas de tráfego"):
                    st.dataframe(df_result, use_container_width=True, height=260)
            else:
                # ── Quando 0 tráfego: ajuda o usuário a entender o que existe ──
                all_phone_matches = df_full[df_full["Venda_Confirmada"] == "SIM"]
                n_any_match = len(all_phone_matches)

                if n_any_match > 0:
                    st.warning(
                        f"Nenhuma venda encontrada com a tag **\"{traffic_keyword}\"**. "
                        f"Porém, **{n_any_match} leads do Kommo têm telefone que bate com suas vendas** — "
                        f"pode ser que a tag de tráfego use um nome diferente."
                    )
                    # Top tags no Kommo para o usuário identificar a tag certa
                    top_tags = (
                        df_kommo_raw[kommo_tag_col]
                        .fillna("").astype(str)
                        .str.split(",")
                        .explode()
                        .str.strip()
                        .loc[lambda s: s != ""]
                        .value_counts()
                        .head(12)
                    )
                    if len(top_tags) > 0:
                        with st.expander("💡 Tags mais usadas no Kommo — qual identifica o tráfego pago?"):
                            st.caption("Cole uma dessas tags no campo 'Tag de tráfego' acima e rode novamente:")
                            for tag, cnt in top_tags.items():
                                st.markdown(f"• **{tag}** — {cnt:,} leads")

                    with st.expander(f"📋 {n_any_match} leads com telefone encontrado nas vendas (sem filtro de tag)"):
                        st.caption(
                            "Esses leads do Kommo têm telefone que bate com uma venda. "
                            "Veja a coluna 'Tag_Kommo' para identificar quais são do tráfego pago."
                        )
                        st.dataframe(all_phone_matches, use_container_width=True, height=300)
                else:
                    st.info(
                        f"Nenhuma venda encontrada com a tag **\"{traffic_keyword}\"** e "
                        f"nenhum telefone do Kommo bateu com as vendas. "
                        f"Verifique se a coluna de telefone está correta nos dois arquivos."
                    )

            # ── Disparo ────────────────────────────────────────────────────────
            if considerar_disparo and disparo_keyword.strip():
                st.divider()
                st.markdown('<div class="step-wrap"><div class="step-num">📣</div><div class="step-text">Disparo (WhatsApp)</div></div>', unsafe_allow_html=True)

                if df_disparo_result is None or disp_total == 0:
                    # Sugere tags alternativas
                    top_tags_d = (
                        df_kommo_disp[kommo_disp_tag_col]
                        .fillna("").astype(str)
                        .str.split(",")
                        .explode()
                        .str.strip()
                        .loc[lambda s: s != ""]
                        .value_counts()
                        .head(12)
                    )
                    st.warning(f"Nenhum lead com a tag **\"{disparo_keyword}\"** encontrado no Kommo.")
                    if len(top_tags_d) > 0:
                        with st.expander("💡 Tags disponíveis no Kommo"):
                            st.caption("Cole uma dessas no campo 'Tag de disparo' acima:")
                            for tag, cnt in top_tags_d.items():
                                st.markdown(f"• **{tag}** — {cnt:,} leads")
                elif disp_conv > 0:
                    taxa_d = f"{disp_conv/disp_total*100:.1f}%"
                    st.success(f"**{disp_conv} vendas** encontradas de disparo — {taxa_d} de conversão sobre {disp_total:,} leads disparados.")
                    if not (kommo_date_col or sales_date_col):
                        st.caption("Sem datas configuradas — todos os matches de telefone foram incluídos. Filtre por Data_Venda no Excel.")
                    st.caption("Veja a lista completa na aba **📣 Vendas — Disparo** do Excel.")
                    with st.expander(f"Prévia — {disp_conv} vendas de disparo"):
                        st.dataframe(
                            df_disparo_result[df_disparo_result["Venda_Confirmada"] == "SIM"],
                            use_container_width=True, height=260,
                        )
                else:
                    st.warning(
                        f"**{disp_total:,} leads de disparo** encontrados, mas nenhuma venda confirmada. "
                        f"Possíveis causas: as vendas ocorreram antes do disparo, fora da janela de 30 dias, "
                        f"ou os telefones não bateram."
                    )

            # ── Sobreposição ────────────────────────────────────────────────────
            if n_overlap > 0:
                st.divider()
                st.info(
                    f"🔀 **{n_overlap} comprador(es)** aparecem em tráfego **e** disparo. "
                    f"Total único de compradores: **{n_total_uniq}**. "
                    f"Detalhes na coluna **'Origem da Venda'** do Excel."
                )

            # ── Qualidade dos dados de vendas ───────────────────────────────────
            st.divider()
            st.markdown('<div class="step-wrap"><div class="step-num">🔍</div><div class="step-text">Qualidade dos Dados</div></div>', unsafe_allow_html=True)

            sit      = df_dup_analysis["Situacao_Venda"]
            n_dup    = int(sit.str.startswith("DUPLICATA").sum())
            n_multi  = int(sit.str.startswith("Multi").sum())

            qd1, qd2, qd3, qd4 = st.columns(4)
            qd1.metric("Total de vendas", f"{n_sales_total:,}")
            qd2.metric("Com telefone", f"{n_com_tel:,}")
            qd3.metric("Sem telefone", f"{n_sem_tel:,}", delta_color="off")
            qd4.metric("Duplicatas suspeitas", f"{n_dup:,}")

            if n_dup > 0:
                st.error(f"⚠️ **{n_dup} registros duplicados** — mesmo número, data e conteúdo idêntico. Veja aba **🔍 Duplicatas** no Excel.")
            if n_multi > 0:
                st.caption(f"🔁 {n_multi} multi-compras identificadas (mesmo cliente, datas diferentes) — normal para clientes recorrentes. Ver aba 🔍 Duplicatas no Excel.")
            if n_dup == 0 and n_multi == 0:
                st.success("Nenhuma duplicata detectada nos dados de vendas.")

            st.divider()

        except Exception:
            progress.empty()
            st.error("Erro ao processar. Detalhes abaixo:")
            st.code(traceback.format_exc())

    # ── Download — fora do bloco de processamento para persistir entre reruns ──
    if "excel_bytes" in st.session_state:
        st.divider()
        st.download_button(
            label="📥  Baixar Excel — Resultado Completo",
            data=st.session_state["excel_bytes"],
            file_name="proc_bulls_resultado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption(
            "💡 Para usar no Google Sheets: faça upload do .xlsx no Google Drive → "
            "clique com botão direito → Abrir com → Planilhas Google."
        )

else:
    st.info("⬆️ Carregue as duas planilhas acima para continuar.")

# ── Rodapé ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    '<div class="footer">Desenvolvido por João &nbsp;·&nbsp; Proc-Bulls v1.0 &nbsp;·&nbsp; Aure Digital</div>',
    unsafe_allow_html=True,
)
