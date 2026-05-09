"""
Proc-Bulls — Ferramenta de PROCV Inteligente
Desenvolvido por João · Aure Digital
"""

import io
import re
import traceback
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
    """
<style>
  .main-title {
    font-size: 3.2rem; font-weight: 900; text-align: center;
    background: linear-gradient(135deg, #A855F7, #7C3AED);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    margin-bottom: 0;
  }
  .subtitle {
    text-align: center; color: #888; font-size: 1rem;
    margin-top: 2px; margin-bottom: 2rem;
  }
  .footer {
    text-align: center; color: #666; font-size: 0.8rem;
    margin-top: 3rem; padding-top: 1rem; border-top: 1px solid #333;
  }
  div[data-testid="stDownloadButton"] button {
    background: linear-gradient(135deg, #27AE60, #1E8449) !important;
    color: white !important; border: none !important;
    border-radius: 8px !important; font-weight: 700 !important;
    font-size: 1rem !important; width: 100% !important;
  }
</style>
""",
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

def load_file(uploaded) -> Optional[pd.DataFrame]:
    name = uploaded.name.lower()
    try:
        if name.endswith(".csv"):
            for enc in ("utf-8", "latin-1", "cp1252", "iso-8859-1"):
                try:
                    uploaded.seek(0)
                    return pd.read_csv(uploaded, encoding=enc, dtype=str)
                except (UnicodeDecodeError, Exception):
                    continue
        elif name.endswith((".xlsx", ".xls", ".xlsm")):
            uploaded.seek(0)
            # Tenta ler todas as abas e pega a primeira com dados
            xls = pd.ExcelFile(uploaded)
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
                if len(df) > 0:
                    return df
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {e}")
    return None


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

    # Sheet 5 — Resumo
    ws5 = wb.create_sheet("Resumo")
    total_traffic = int((df_full["É_Tráfego"] == "SIM").sum())
    total_sales = len(ds)
    confirmed = len(df_result)
    conv_rate = f"{confirmed/total_traffic*100:.1f}%" if total_traffic > 0 else "—"

    ws5.merge_cells("A1:C1")
    t = ws5.cell(1, 1, "RESUMO DO PROC-BULLS")
    t.font = Font(color="FFFFFF", bold=True, size=14)
    t.fill = PatternFill("solid", fgColor="FF6B35")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 36

    rows = [
        ("Total de vendas carregadas", total_sales),
        ("Total de leads no Kommo", len(dk)),
        ("Leads com tag de tráfego", total_traffic),
        ("Conversões confirmadas (tráfego → venda)", confirmed),
        ("Taxa de conversão do tráfego", conv_rate),
    ]
    for i, (label, val) in enumerate(rows, 2):
        ws5.cell(i, 1, label).font = Font(bold=True, size=11)
        c = ws5.cell(i, 2, val)
        c.font = Font(bold=True, size=13, color="FF6B35")
        c.alignment = Alignment(horizontal="center")
        ws5.row_dimensions[i].height = 26

    ws5.column_dimensions["A"].width = 46
    ws5.column_dimensions["B"].width = 22

    # Rodapé em todas as abas
    for ws in [ws1, ws2, ws3, ws4, ws5]:
        r = ws.max_row + 2
        ws.cell(r, 1, "Desenvolvido por João  ·  Proc-Bulls  ·  Aure Digital").font = Font(
            italic=True, color="AAAAAA", size=9
        )

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Interface ──────────────────────────────────────────────────────────────────

st.markdown('<h1 class="main-title" translate="no">🎯 PROC-BULLS</h1>', unsafe_allow_html=True)
st.markdown(
    '<p class="subtitle">Ferramenta de PROCV Inteligente · Análise de Conversão de Tráfego · Aure Digital</p>',
    unsafe_allow_html=True,
)
st.divider()

# ── Passo 1: Upload ────────────────────────────────────────────────────────────
st.markdown("### 1 · Carregue as planilhas")

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
        df_sales_raw = load_file(sales_file)
        if df_sales_raw is not None:
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
        df_kommo_raw = load_file(kommo_file)
        if df_kommo_raw is not None:
            st.success(f"✅ {len(df_kommo_raw):,} linhas · {len(df_kommo_raw.columns)} colunas")
            with st.expander("Prévia"):
                st.dataframe(df_kommo_raw.head(6), use_container_width=True)

st.divider()

# ── Passo 2: Configuração ──────────────────────────────────────────────────────
if df_sales_raw is not None and df_kommo_raw is not None:
    st.markdown("### 2 · Configure as colunas")

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

    st.divider()

    # ── Passo 3: Processar ─────────────────────────────────────────────────────
    st.markdown("### 3 · Processar")

    if st.button("🎯  RODAR PROC-BULLS", use_container_width=True, type="primary"):
        with st.spinner("Tratando dados com maldade... 🧠"):
            try:
                ds_t, dk_t, df_result, df_full = run_procv(
                    df_sales_raw, sales_phone_col,
                    df_kommo_raw, kommo_phone_col, kommo_tag_col,
                    traffic_keyword,
                )

                # ── Métricas ───────────────────────────────────────────────────
                st.markdown("### Resultados")
                total_traffic = int((df_full["É_Tráfego"] == "SIM").sum())
                confirmed = len(df_result)

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

                if confirmed > 0:
                    st.success(
                        f"🎉 {confirmed} leads de tráfego com venda confirmada!"
                    )
                    st.dataframe(df_result, use_container_width=True, height=280)
                else:
                    st.warning(
                        "Nenhuma conversão encontrada. "
                        "Verifique a palavra-chave da tag e as colunas selecionadas."
                    )

                # ── Download ───────────────────────────────────────────────────
                excel_bytes = build_excel(ds_t, dk_t, df_result, df_full)
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
