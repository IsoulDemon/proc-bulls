"""
Bateria PROFUNDA de cenários — Proc Aure.
Exercita ingestão, telefone, datas, valores, PROCV (tráfego), disparo,
duplicatas, end-to-end e performance, em dezenas de cenários variados.

Rodar:  python3 tests/test_cenarios.py
"""
import io
import os
import sys
import time
import warnings
from datetime import datetime

warnings.filterwarnings("ignore")
os.environ["STREAMLIT_GLOBAL_DISABLE_WATCHDOG_WARNING"] = "true"
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd  # noqa: E402
import app  # noqa: E402

_results = []   # (categoria, nome, ok, detalhe)
_cur_cat = "?"


def cat(c):
    global _cur_cat
    _cur_cat = c
    print(f"\n── {c} " + "─" * (60 - len(c)))


def ck(nome, cond, detalhe=""):
    ok = bool(cond)
    _results.append((_cur_cat, nome, ok, detalhe))
    print(f"  [{'PASS' if ok else 'FALHOU'}] {nome}" + (f"  → {detalhe}" if not ok else ""))


# ── Helpers de arquivo em memória ────────────────────────────────────────────
def make_xlsx(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, rows in sheets.items():
            pd.DataFrame(rows).to_excel(w, sheet_name=name, header=False, index=False)
    return buf.getvalue()


def load_xlsx(sheets: dict):
    b = make_xlsx(sheets)
    fl = app._FileLike(b, "t.xlsx")
    names = app.get_excel_sheets(fl)
    fl2 = app._FileLike(b, "t.xlsx")
    df, info = app.load_file_multisheet(fl2, names)
    return df


def load_csv(text: str):
    fl = app._FileLike(text.encode("utf-8"), "t.csv")
    df, info = app.load_file_multisheet(fl, [])
    return df


# ════════════════════════════════════════════════════════════════════════════
# A. Limpeza e chave de telefone
# ════════════════════════════════════════════════════════════════════════════
def test_telefone():
    cat("A. Telefone — limpeza e chave")
    casos = {
        "+55 (66) 99987-3776": ("66", "99873776"),
        "5511988887777":       ("11", "88887777"),
        "551133334444":        ("11", "33334444"),
        "11 9 9999-8888":      ("11", "99998888"),
        "(21)97777-6666":      ("21", "77776666"),
        "66999873776.0":       ("66", "99873776"),
        "6.6e+10":             (None, ""),   # placeholder 66000000000
        "tel: 66 99987-3776":  ("66", "99873776"),
    }
    for raw, exp in casos.items():
        got = app.phone_key(raw)
        ck(f"phone_key({raw!r})=={exp}", got == exp, str(got))

    # múltiplos números numa célula
    ks = app._phone_keys_in_cell("66999873776 / 11988887777")
    subs = sorted(k[1] for k in ks)
    ck("2 números numa célula extraídos", subs == ["88887777", "99873776"], str(subs))

    # casamento DDD-aware
    ck("com/sem 9 casa", app.phones_match(app.phone_key("66999873776"), app.phone_key("669987 3776")))
    ck("com/sem DDD casa", app.phones_match(app.phone_key("66999873776"), app.phone_key("99873776")))
    ck("DDDs diferentes não casam",
       not app.phones_match(app.phone_key("11999998888"), app.phone_key("21999998888")))


# ════════════════════════════════════════════════════════════════════════════
# B. Discriminação telefone × CPF/CNPJ/data/ID
# ════════════════════════════════════════════════════════════════════════════
def test_discriminacao():
    cat("B. Discriminação telefone × não-telefone")
    ck("data ISO não é telefone", not app._looks_like_phone("2026-03-15"))
    ck("CPF (3º díg != 9) não é telefone", not app._looks_like_phone("123.456.789-09"))
    ck("CNPJ não é telefone", not app._looks_like_phone("12.345.678/0001-99"))
    ck("DDD inválido (00) não é telefone", not app._looks_like_phone("0099998888"))
    ck("celular real é telefone", app._looks_like_phone("66999873776"))
    ck("fixo com DDD é telefone", app._looks_like_phone("1133334444"))

    # detecção de coluna: CPF + telefone reais → escolhe telefone
    df = pd.DataFrame({
        "CPF": ["123.456.789-09", "987.654.321-00", "111.222.333-44"],
        "Celular": ["66999873776", "11988887777", "21999990000"],
    })
    ck("detect_phone_col escolhe Celular (não CPF)", app.detect_phone_col(df) == "Celular",
       str(app.detect_phone_col(df)))

    # coluna de ID numérica de 8 dígitos não vira telefone (skip por nome)
    df2 = pd.DataFrame({"id_pedido": ["10000001", "10000002", "10000003"],
                        "Telefone": ["66999873776", "11988887777", "21999990000"]})
    ck("detect_phone_col ignora id_pedido", app.detect_phone_col(df2) == "Telefone",
       str(app.detect_phone_col(df2)))


# ════════════════════════════════════════════════════════════════════════════
# C. Datas
# ════════════════════════════════════════════════════════════════════════════
def test_datas():
    cat("C. parse_date — formatos variados")
    def y_m_d(s, y, m, d):
        dt = app.parse_date(s)
        ck(f"{s!r} -> {y}-{m:02d}-{d:02d}",
           dt is not None and (dt.year, dt.month, dt.day) == (y, m, d), str(dt))
    y_m_d("15/03/2026", 2026, 3, 15)
    y_m_d("2026-03-15", 2026, 3, 15)
    y_m_d("03/15/2026", 2026, 3, 15)           # US: dia 15 inválido como mês → inverte
    y_m_d("abril/26", 2026, 4, 1)
    y_m_d("março 2024", 2024, 3, 1)
    y_m_d("22/04/2026 17:42", 2026, 4, 22)
    dt = app.parse_date("45292")               # serial Excel
    ck("serial 45292 -> 2024-01", dt is not None and (dt.year, dt.month) == (2024, 1), str(dt))
    for vazio in ("", "nan", "-", "n/a", None):
        ck(f"{vazio!r} -> None", app.parse_date(vazio) is None)
    # data embutida em texto livre
    dt = app.parse_date("Disparo dia das Mães 22/04")
    ck("texto livre '...22/04' extrai dia 22/04", dt is not None and (dt.month, dt.day) == (4, 22), str(dt))


# ════════════════════════════════════════════════════════════════════════════
# D. Valores monetários
# ════════════════════════════════════════════════════════════════════════════
def test_valores():
    cat("D. Valor — parse e detecção")
    ck("'1.234,56' -> 1234.56", abs(app.parse_value("1.234,56") - 1234.56) < 0.001, str(app.parse_value("1.234,56")))
    ck("'R$ 99,90' -> 99.90", abs(app.parse_value("R$ 99,90") - 99.90) < 0.001, str(app.parse_value("R$ 99,90")))
    ck("'1234.56' -> 1234.56", abs(app.parse_value("1234.56") - 1234.56) < 0.001, str(app.parse_value("1234.56")))
    ck("vazio -> 0", app.parse_value("") == 0.0)
    df = pd.DataFrame({
        "Telefone": ["66999873776", "11988887777", "21999990000", "31988887777"],
        "Valor": ["199,90", "1.250,00", "89,90", "450,00"],
    })
    ck("detect_value_col escolhe Valor (não Telefone)", app.detect_value_col(df) == "Valor",
       str(app.detect_value_col(df)))


# ════════════════════════════════════════════════════════════════════════════
# E. Ingestão de arquivos (load_file_multisheet)
# ════════════════════════════════════════════════════════════════════════════
def test_ingestao():
    cat("E. Ingestão — header, sem header, multi-aba, blocos, CSV")

    # E1: título antes do cabeçalho
    df = load_xlsx({"Plan1": [
        ["Relatório de Vendas - Abril", None, None],
        ["Nome", "Telefone", "Valor"],
        ["Joao", "11988887777", "100"],
        ["Maria", "21999990000", "200"],
    ]})
    ck("E1 título ignorado, header certo",
       df is not None and list(df.columns)[:3] == ["Nome", "Telefone", "Valor"] and len(df) == 2,
       None if df is None else f"cols={list(df.columns)} n={len(df)}")

    # E2: CSV sem cabeçalho
    df = load_csv("Joao,11988887777,100\nMaria,21999990000,200\nJose,31988887777,150")
    ck("E2 CSV sem cabeçalho -> Coluna N",
       df is not None and len(df) == 3 and any(str(c).startswith("Coluna") for c in df.columns),
       None if df is None else f"cols={list(df.columns)} n={len(df)}")

    # E3: CSV com ponto-e-vírgula
    df = load_csv("Nome;Telefone;Valor\nAna;11988887777;100\nBia;21999990000;200")
    ck("E3 CSV ; detecta separador",
       df is not None and "Telefone" in df.columns and len(df) == 2,
       None if df is None else f"cols={list(df.columns)} n={len(df)}")

    # E4: duas abas com cabeçalho divergente (alinhamento)
    df = load_xlsx({
        "Parte1": [["Telefone", "Valor"], ["11988887777", "100"], ["31988887777", "300"]],
        "Parte2": [["Telefone ", "Valor"], ["21999990000", "200"]],   # espaço no fim
    })
    tel_cols = [c for c in df.columns if app._canon_colname(c) == "telefone"] if df is not None else []
    ck("E4 colunas Telefone unificadas (sem fragmentar)",
       df is not None and len(tel_cols) == 1 and df[tel_cols[0]].notna().sum() == 3,
       None if df is None else f"tel_cols={tel_cols} cols={list(df.columns)}")

    # E5: blocos lado a lado
    df = load_xlsx({"Plan1": [
        ["Nome", "Tel", None, "Nome", "Tel"],
        ["A", "11988887777", None, "B", "21999990000"],
        ["C", "31988887777", None, "D", "41977776666"],
    ]})
    ck("E5 blocos lado a lado reempilhados",
       df is not None and len(df) == 4 and len([c for c in df.columns if app._canon_colname(c) == "tel"]) == 1,
       None if df is None else f"cols={list(df.columns)} n={len(df)}")

    # E6: planilha normal multi-aba (vendas + cadastro auxiliar)
    df = load_xlsx({
        "Vendas": [["Nome", "Telefone", "Data", "Valor"],
                   ["Joao", "11988887777", "10/04/2026", "100"],
                   ["Maria", "21999990000", "11/04/2026", "200"]],
        "Cadastro": [["Cliente", "Fone"],
                     ["Pedro", "31988887777"]],
    })
    ck("E6 multi-aba carrega tudo", df is not None and len(df) == 3, None if df is None else f"n={len(df)}")


# ════════════════════════════════════════════════════════════════════════════
# F. PROCV — tráfego
# ════════════════════════════════════════════════════════════════════════════
def _procv(sales, sphone, kommo, kphone, ktag, kw, **kw2):
    return app.run_procv(sales, sphone, kommo, kphone, ktag, kw, **kw2)


def test_procv():
    cat("F. PROCV — tráfego")

    # F1: tag COM acento, keyword sem acento
    sales = pd.DataFrame({"Nome": ["Joao"], "Telefone": ["11988887777"], "Valor": ["100"]})
    kommo = pd.DataFrame({"Celular": ["11988887777"], "Tags": ["Tráfego Pago"]})
    _, _, traf, _ = _procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("F1 tag com acento casa keyword sem acento", len(traf) == 1, f"traf={len(traf)}")

    # F2: precisão por DDD
    sales = pd.DataFrame({"Nome": ["SP", "RJ"], "Telefone": ["11999998888", "21999998888"], "Valor": ["1", "2"]})
    kommo = pd.DataFrame({"Celular": ["11999998888", "31999998888"], "Tags": ["trafego", "trafego"]})
    _, _, _, full = _procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("F2 DDD: lead SP casa SP", full.iloc[0]["[Venda] Nome"] == "SP", str(full.iloc[0]["[Venda] Nome"]))
    ck("F2 DDD: lead 31 não casa", full.iloc[1]["Venda_Confirmada"] == "NÃO")

    # F3: telefone em coluna alternativa do Kommo (Celular vazio, número em Telefone2)
    sales = pd.DataFrame({"Nome": ["Joao", "Ana"], "Telefone": ["11988887777", "21999990000"], "Valor": ["100", "200"]})
    kommo = pd.DataFrame({"Celular": ["", ""], "Telefone2": ["11988887777", "21999990000"],
                          "Tags": ["trafego", "trafego pago"]})
    _, _, traf, _ = _procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("F3 match por coluna alternativa do Kommo", len(traf) == 2, f"traf={len(traf)}")

    # F4: fallback por nome (telefone não bate)
    sales = pd.DataFrame({"Nome": ["Joao Carlos Silva"], "Telefone": ["11988887777"], "Valor": ["100"]})
    kommo = pd.DataFrame({"Celular": ["99999999999"], "Tags": ["trafego"], "Nome": ["Joao Carlos Silva"]})
    _, _, traf, _ = _procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego",
                           sales_name_col="Nome", kommo_name_col="Nome")
    ck("F4 fallback por nome completo", len(traf) == 1, f"traf={len(traf)}")

    # F5: dedup — mesmo comprador em 2 leads conta 1
    sales = pd.DataFrame({"Nome": ["Joao"], "Telefone": ["11988887777"], "Valor": ["100"]})
    kommo = pd.DataFrame({"Celular": ["11988887777", "11988887777"], "Tags": ["trafego", "trafego pago"]})
    _, _, traf, _ = _procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("F5 dedup por comprador único", len(traf) == 1, f"traf={len(traf)}")

    # F6: telefone solto fora da coluna (vendas)
    sales = pd.DataFrame({"Nome": ["A", "B"], "Telefone": ["11988887777", ""],
                          "Obs": ["ok", "zap 21999990000"], "Valor": ["1", "2"]})
    kommo = pd.DataFrame({"Celular": ["21999990000"], "Tags": ["trafego"]})
    _, _, traf, full = _procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("F6 telefone solto é cruzado", len(traf) == 1 and full.attrs.get("n_stray_sales", 0) >= 1,
       f"traf={len(traf)} stray={full.attrs.get('n_stray_sales')}")


# ════════════════════════════════════════════════════════════════════════════
# G. Disparo
# ════════════════════════════════════════════════════════════════════════════
def _disp(sales, sphone, sdate, kommo, kphone, ktag, kw, kdate=None, **kw2):
    return app.run_disparo(sales, sphone, sdate, kommo, kphone, ktag, kw, kdate, **kw2)


def test_disparo():
    cat("G. Disparo — janela, datas, acento, multi-compra")

    base_sales = lambda dt: pd.DataFrame({
        "Nome": ["Cliente"], "Telefone": ["66999873776"], "Data": [dt], "Valor": ["100"]})

    # G1: venda DEPOIS do disparo, dentro da janela (data na tag)
    k = pd.DataFrame({"Celular": ["66999873776"], "Tags": ["recebeu disparo 22/04"]})
    r = _disp(base_sales("25/04/2026"), "Telefone", "Data", k, "Celular", "Tags", "disparo")
    ck("G1 venda 3 dias após disparo = confirma", int((r["Venda_Confirmada"] == "SIM").sum()) == 1,
       f"conf={int((r['Venda_Confirmada']=='SIM').sum())}")

    # G2: venda ANTES do disparo = não conta
    r = _disp(base_sales("10/04/2026"), "Telefone", "Data", k, "Celular", "Tags", "disparo")
    ck("G2 venda antes do disparo = NÃO conta", int((r["Venda_Confirmada"] == "SIM").sum()) == 0,
       f"conf={int((r['Venda_Confirmada']=='SIM').sum())}")

    # G3: venda além da janela (>30 dias) = não conta
    r = _disp(base_sales("30/05/2026"), "Telefone", "Data", k, "Celular", "Tags", "disparo")
    ck("G3 venda 38 dias após = NÃO conta", int((r["Venda_Confirmada"] == "SIM").sum()) == 0,
       f"conf={int((r['Venda_Confirmada']=='SIM').sum())}")

    # G4: sem datas configuradas = casamento por telefone (lenient)
    k2 = pd.DataFrame({"Celular": ["66999873776"], "Tags": ["disparo"]})
    r = _disp(base_sales("25/04/2026"), "Telefone", None, k2, "Celular", "Tags", "disparo")
    ck("G4 sem datas = casa por telefone", int((r["Venda_Confirmada"] == "SIM").sum()) == 1)

    # G5: tag de disparo COM acento (keyword 'maes' deve achar 'Mães')
    k3 = pd.DataFrame({"Celular": ["66999873776"], "Tags": ["Disparo Dia das Mães 22/04"]})
    r = _disp(base_sales("25/04/2026"), "Telefone", "Data", k3, "Celular", "Tags", "maes")
    ck("G5 keyword sem acento acha tag com acento", len(r) >= 1 and int((r["Venda_Confirmada"] == "SIM").sum()) == 1,
       f"linhas={len(r)} conf={int((r['Venda_Confirmada']=='SIM').sum()) if len(r) else 0}")

    # G6: multi-compra — uma antes, uma depois → conta a de depois
    sales = pd.DataFrame({
        "Nome": ["Cliente", "Cliente"], "Telefone": ["66999873776", "66999873776"],
        "Data": ["10/04/2026", "25/04/2026"], "Valor": ["50", "100"]})
    r = _disp(sales, "Telefone", "Data", k, "Celular", "Tags", "disparo")
    ck("G6 multi-compra: pega a venda pós-disparo", int((r["Venda_Confirmada"] == "SIM").sum()) == 1,
       f"conf={int((r['Venda_Confirmada']=='SIM').sum())}")

    # G7: precisão por DDD no disparo
    sales = pd.DataFrame({"Nome": ["SP", "RJ"], "Telefone": ["11999998888", "21999998888"],
                          "Data": ["25/04/2026", "25/04/2026"], "Valor": ["1", "2"]})
    k4 = pd.DataFrame({"Celular": ["11999998888"], "Tags": ["disparo 22/04"]})
    r = _disp(sales, "Telefone", "Data", k4, "Celular", "Tags", "disparo")
    conf = r[r["Venda_Confirmada"] == "SIM"]
    ck("G7 disparo casa só a venda do DDD certo",
       len(conf) == 1 and conf.iloc[0]["[Venda] Nome"] == "SP",
       f"conf={len(conf)}")

    # G8: data do disparo na COLUNA do kommo (não na tag)
    k5 = pd.DataFrame({"Celular": ["66999873776"], "Tags": ["disparo"], "DataDisparo": ["22/04/2026"]})
    r = _disp(base_sales("25/04/2026"), "Telefone", "Data", k5, "Celular", "Tags", "disparo", kdate="DataDisparo")
    ck("G8 data do disparo pela coluna do kommo", int((r["Venda_Confirmada"] == "SIM").sum()) == 1)


# ════════════════════════════════════════════════════════════════════════════
# H. Duplicatas
# ════════════════════════════════════════════════════════════════════════════
def test_duplicatas():
    cat("H. Duplicatas — classificação")

    # H1: duplicata exata
    df = pd.DataFrame({"Telefone": ["11988887777", "11988887777"],
                       "Data": ["10/04/2026", "10/04/2026"], "Produto": ["X", "X"], "Valor": ["100", "100"]})
    res = app.analyze_duplicates(df, "Telefone", "Data")
    ck("H1 duplicata exata -> DUPLICATA", (res["Situacao_Venda"] == "DUPLICATA").all(),
       str(list(res["Situacao_Venda"])))

    # H2: multi-compra (datas diferentes)
    df = pd.DataFrame({"Telefone": ["11988887777", "11988887777"],
                       "Data": ["10/04/2026", "20/04/2026"], "Produto": ["X", "Y"], "Valor": ["100", "200"]})
    res = app.analyze_duplicates(df, "Telefone", "Data")
    ck("H2 datas diferentes -> Multi-compra", (res["Situacao_Venda"] == "Multi-compra").all(),
       str(list(res["Situacao_Venda"])))

    # H3: DDDs diferentes não são duplicata
    df = pd.DataFrame({"Telefone": ["11999998888", "21999998888"],
                       "Data": ["10/04/2026", "10/04/2026"], "Produto": ["X", "X"], "Valor": ["100", "100"]})
    res = app.analyze_duplicates(df, "Telefone", "Data")
    ck("H3 DDDs diferentes -> Única", (res["Situacao_Venda"] == "Única").all(),
       str(list(res["Situacao_Venda"])))

    # H4: telefones vazios não viram grupo de duplicata
    df = pd.DataFrame({"Telefone": ["", "", ""], "Data": ["10/04/2026", "11/04/2026", "12/04/2026"],
                       "Produto": ["X", "Y", "Z"], "Valor": ["1", "2", "3"]})
    res = app.analyze_duplicates(df, "Telefone", "Data")
    ck("H4 telefones vazios -> todos Única", (res["Situacao_Venda"] == "Única").all(),
       str(list(res["Situacao_Venda"])))


# ════════════════════════════════════════════════════════════════════════════
# I. End-to-end + build_excel
# ════════════════════════════════════════════════════════════════════════════
def test_end_to_end():
    cat("I. End-to-end + geração de Excel")
    sales = pd.DataFrame({
        "Nome": ["Joao SP", "Maria RJ", "Pedro MT", "Ana"],
        "Telefone": ["11999998888", "21999998888", "66999873776", ""],
        "Tel2": ["", "", "", "31988887777"],
        "Data": ["25/04/2026", "10/04/2026", "25/04/2026", "26/04/2026"],
        "Valor": ["199,90", "299,00", "150,00", "99,90"],
    })
    kommo = pd.DataFrame({
        "Celular": ["11999998888", "31999998888", "66999873776", "31988887777"],
        "Tags": ["Tráfego Pago", "trafego", "recebeu disparo 22/04", "trafego"],
        "Nome": ["Joao SP", "X", "Pedro MT", "Ana"],
    })
    ds, dk, traf, full = app.run_procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego",
                                       sales_name_col="Nome", kommo_name_col="Nome")
    disp = app.run_disparo(sales, "Telefone", "Data", kommo, "Celular", "Tags", "disparo", None,
                           sales_name_col="Nome", kommo_name_col="Nome")
    dup = app.analyze_duplicates(sales, "Telefone", "Data")
    xls = app.build_excel(ds, dk, traf, full, disp, dup, sales_value_col="Valor")
    ck("I1 pipeline completo gera Excel", isinstance(xls, (bytes, bytearray)) and len(xls) > 2000, f"{len(xls)}B")
    ck("I1 tráfego com acento contado", len(traf) >= 1, f"traf={len(traf)}")

    # I2: build_excel com resultados vazios (nenhuma conversão) não quebra
    empty_kommo = pd.DataFrame({"Celular": ["99999999999"], "Tags": ["outro"]})
    ds2, dk2, traf2, full2 = app.run_procv(sales, "Telefone", empty_kommo, "Celular", "Tags", "trafego")
    dup2 = app.analyze_duplicates(sales, "Telefone", "Data")
    xls2 = app.build_excel(ds2, dk2, traf2, full2, None, dup2, sales_value_col=None)
    ck("I2 Excel sem conversões não quebra", isinstance(xls2, (bytes, bytearray)) and len(xls2) > 1000, f"{len(xls2)}B")


# ════════════════════════════════════════════════════════════════════════════
# J. Performance
# ════════════════════════════════════════════════════════════════════════════
def test_performance():
    cat("J. Performance — dataset grande")
    n = 5000
    ddds = ["11", "21", "31", "41", "51"]
    sales = pd.DataFrame({
        "Nome": [f"Cliente {i}" for i in range(n)],
        "Telefone": [f"{ddds[i % 5]}9{(80000000 + i):08d}" for i in range(n)],
        "Data": ["20/04/2026"] * n,
        "Valor": ["100,00"] * n,
    })
    kommo = pd.DataFrame({
        "Celular": [f"{ddds[i % 5]}9{(80000000 + i):08d}" for i in range(0, n, 2)],
        "Tags": ["trafego pago"] * (n // 2),
    })
    t0 = time.time()
    _, _, traf, _ = app.run_procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    dt = time.time() - t0
    ck(f"J1 run_procv {n}x{n//2} < 25s", dt < 25, f"{dt:.2f}s")
    ck("J1 metade dos leads convertida", abs(len(traf) - n // 2) <= 2, f"traf={len(traf)}")


# ════════════════════════════════════════════════════════════════════════════
# K. Cenários adversariais
# ════════════════════════════════════════════════════════════════════════════
def test_adversarial():
    cat("K. Adversariais — combinações difíceis")

    # K1: TRÊS blocos lado a lado
    df = load_xlsx({"Plan1": [
        ["Nome", "Tel", None, "Nome", "Tel", None, "Nome", "Tel"],
        ["A", "11988887777", None, "B", "21999990000", None, "C", "31988887777"],
        ["D", "41977776666", None, "E", "51966665555", None, "F", "61955554444"],
    ]})
    ck("K1 três blocos reempilhados", df is not None and len(df) == 6,
       None if df is None else f"n={len(df)}")

    # K2: nome ambíguo (2 pessoas, mesmo nome) NÃO gera falso match
    sales = pd.DataFrame({"Nome": ["Joao Silva", "Joao Silva"],
                          "Telefone": ["11988887777", "11977776666"], "Valor": ["1", "2"]})
    kommo = pd.DataFrame({"Celular": ["99999999999"], "Tags": ["trafego"], "Nome": ["Joao Silva"]})
    _, _, traf, _ = app.run_procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego",
                                  sales_name_col="Nome", kommo_name_col="Nome")
    ck("K2 nome ambíguo não vira falso match", len(traf) == 0, f"traf={len(traf)}")

    # K3: aba secundária (cadastro) ainda cruza telefone
    combined = load_xlsx({
        "Vendas": [["Nome", "Telefone", "Data", "Valor"], ["Joao", "11988887777", "10/04/2026", "100"]],
        "Cadastro": [["Cliente", "Fone"], ["Pedro", "66999873776"], ["Lucas", "21999990000"]],
    })
    kommo = pd.DataFrame({"Celular": ["66999873776"], "Tags": ["trafego"]})
    _, _, traf, _ = app.run_procv(combined, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("K3 telefone de aba secundária é cruzado", len(traf) == 1, f"traf={len(traf)}")

    # K4: formatos MISTOS de telefone na mesma coluna (load + procv)
    sales_b = make_xlsx({"V": [["Nome", "Telefone"],
                               ["A", "+55 (11) 98888-7777"], ["B", "21 99999-0000"], ["C", "5531988887777"]]})
    sales = app.load_file_multisheet(app._FileLike(sales_b, "v.xlsx"), app.get_excel_sheets(app._FileLike(sales_b, "v.xlsx")))[0]
    kommo = pd.DataFrame({"Celular": ["11988887777", "21999990000", "31988887777"],
                          "Tags": ["trafego", "trafego", "trafego"]})
    _, _, traf, _ = app.run_procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("K4 formatos mistos casam todos", len(traf) == 3, f"traf={len(traf)}")

    # K5: arquivo de disparo SEPARADO, keyword diferente
    sales = pd.DataFrame({"Nome": ["Cliente"], "Telefone": ["66999873776"], "Data": ["25/04/2026"], "Valor": ["100"]})
    kommo_disp = pd.DataFrame({"Celular": ["66999873776"], "Tags": ["Black Friday Disparo"]})
    r = app.run_disparo(sales, "Telefone", None, kommo_disp, "Celular", "Tags", "black friday", None)
    ck("K5 disparo de arquivo/keyword separados", int((r["Venda_Confirmada"] == "SIM").sum()) == 1,
       f"conf={int((r['Venda_Confirmada']=='SIM').sum()) if len(r) else 0}")

    # K6: keyword da tag com espaços nas pontas
    kommo = pd.DataFrame({"Celular": ["66999873776"], "Tags": ["Tráfego Pago"]})
    _, _, traf, _ = app.run_procv(sales, "Telefone", kommo, "Celular", "Tags", "  trafego  ")
    ck("K6 keyword com espaços nas pontas casa", len(traf) == 1, f"traf={len(traf)}")

    # K7: data BR ambígua "05/04/2026" = 5 de abril
    dt = app.parse_date("05/04/2026")
    ck("K7 '05/04/2026' = 5 abril (BR)", dt is not None and (dt.month, dt.day) == (4, 5), str(dt))


# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    for fn in (test_telefone, test_discriminacao, test_datas, test_valores,
               test_ingestao, test_procv, test_disparo, test_duplicatas,
               test_end_to_end, test_adversarial, test_performance):
        try:
            fn()
        except Exception as e:
            import traceback
            _results.append((_cur_cat, fn.__name__ + " (EXCEÇÃO)", False, repr(e)))
            print(f"  [EXCEÇÃO] {fn.__name__}: {e}")
            traceback.print_exc()

    print("\n" + "=" * 64)
    cats = {}
    for c, _, ok, _ in _results:
        p, t = cats.get(c, (0, 0))
        cats[c] = (p + (1 if ok else 0), t + 1)
    for c in sorted(cats):
        p, t = cats[c]
        mark = "✅" if p == t else "❌"
        print(f"  {mark} {c}: {p}/{t}")
    fails = [(c, n, d) for c, n, ok, d in _results if not ok]
    total = len(_results)
    print(f"\n  TOTAL: {total - len(fails)}/{total} passaram")
    if fails:
        print("\n  FALHAS:")
        for c, n, d in fails:
            print(f"   • [{c}] {n}" + (f"  → {d}" if d else ""))
        sys.exit(1)
    print("\n✅ Bateria completa passou.")
