"""
Bateria MULTI-MÊS / MULTI-ABA — Proc Aure.
Cenário real: cliente manda 1 planilha de vendas com várias abas (1 por mês) e
1 export do Kommo. A ferramenta cruza o Kommo contra CADA mês → vendas de
tráfego (e disparo) mês a mês, sem deduplicar o comprador recorrente entre meses.

Rodar:  python3 tests/test_multimes.py
"""
import io
import os
import sys
import time
import warnings

warnings.filterwarnings("ignore")
os.environ["STREAMLIT_GLOBAL_DISABLE_WATCHDOG_WARNING"] = "true"
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd  # noqa: E402
import app  # noqa: E402

_results = []
_cur = "?"


def cat(c):
    global _cur
    _cur = c
    print(f"\n── {c} " + "─" * max(2, 58 - len(c)))


def ck(nome, cond, detalhe=""):
    ok = bool(cond)
    _results.append((_cur, nome, ok, detalhe))
    print(f"  [{'PASS' if ok else 'FALHOU'}] {nome}" + (f"  → {detalhe}" if not ok else ""))


def make_wb(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, rows in sheets.items():
            pd.DataFrame(rows).to_excel(w, sheet_name=name, header=False, index=False)
    return buf.getvalue()


def load_wb(sheets: dict, selected=None):
    b = make_wb(sheets)
    names = app.get_excel_sheets(app._FileLike(b, "v.xlsx"))
    return app.load_file_multisheet(app._FileLike(b, "v.xlsx"), selected or names)[0]


MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto"]
RECORRENTE = "11988887777"  # compra todo mês


def workbook_8_meses():
    """8 abas. Recorrente compra em todos os meses; cada mês tem 1 cliente único."""
    sheets = {}
    for i, m in enumerate(MESES, 1):
        sheets[m] = [
            ["Nome", "Telefone", "Data", "Valor"],
            [f"Único {m}", f"119{i:08d}", f"15/{i:02d}/2026", "100,00"],
            ["Recorrente", RECORRENTE, f"20/{i:02d}/2026", "200,00"],
        ]
    return sheets


# ════════════════════════════════════════════════════════════════════════════
# M1. Carga do workbook de 8 meses
# ════════════════════════════════════════════════════════════════════════════
def test_carga():
    cat("M1. Carga — workbook de 8 meses")
    sales = load_wb(workbook_8_meses())
    ck("8 abas → coluna _Planilha presente", sales is not None and "_Planilha" in sales.columns)
    ck("16 vendas no total", sales is not None and len(sales) == 16, None if sales is None else f"n={len(sales)}")
    ck("8 meses distintos preservados",
       sales is not None and sales["_Planilha"].nunique() == 8,
       None if sales is None else str(sales["_Planilha"].unique()))


# ════════════════════════════════════════════════════════════════════════════
# M2. Recorrente: combinado deduplica (1), por-mês conta cada mês (8)
# ════════════════════════════════════════════════════════════════════════════
def test_recorrente():
    cat("M2. Comprador recorrente — combinado × mês a mês")
    sales = load_wb(workbook_8_meses())
    kommo = pd.DataFrame({"Celular": [RECORRENTE], "Tags": ["trafego pago"]})

    # combinado: comprador único → 1 conversão
    _, _, traf, _ = app.run_procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("combinado deduplica recorrente → 1", len(traf) == 1, f"traf={len(traf)}")

    # mês a mês: conta o recorrente em CADA mês → 8
    bd = app.run_breakdown_by_sheet(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("breakdown retorna 8 linhas (meses)", bd is not None and len(bd) == 8, None if bd is None else f"n={len(bd)}")
    if bd is not None:
        ck("soma das vendas de tráfego mês a mês = 8", int(bd["Vendas de Tráfego"].sum()) == 8,
           str(bd["Vendas de Tráfego"].tolist()))
        ck("cada mês tem exatamente 1 conversão de tráfego",
           (bd["Vendas de Tráfego"] == 1).all(), str(bd["Vendas de Tráfego"].tolist()))


# ════════════════════════════════════════════════════════════════════════════
# M3. Atribuição correta por mês (cliente que só comprou num mês)
# ════════════════════════════════════════════════════════════════════════════
def test_atribuicao():
    cat("M3. Atribuição por mês")
    sheets = workbook_8_meses()
    # cliente que só comprou em Março, lead de tráfego
    so_marco = "21999990000"
    sheets["Março"].append(["Cliente Março", so_marco, "10/03/2026", "300,00"])
    sales = load_wb(sheets)
    kommo = pd.DataFrame({"Celular": [RECORRENTE, so_marco], "Tags": ["trafego", "trafego"]})
    bd = app.run_breakdown_by_sheet(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    by = {r["Mês / Aba"]: r["Vendas de Tráfego"] for _, r in bd.iterrows()}
    ck("Março tem 2 conversões (recorrente + cliente de março)", by.get("Março") == 2, str(by))
    ck("Janeiro tem 1 conversão", by.get("Janeiro") == 1, str(by))
    ck("total = 9 (8 recorrente + 1 março)", int(bd["Vendas de Tráfego"].sum()) == 9, str(by))


# ════════════════════════════════════════════════════════════════════════════
# M4. Layouts DIFERENTES por mês (nome da coluna de telefone varia)
# ════════════════════════════════════════════════════════════════════════════
def test_layouts_diferentes():
    cat("M4. Colunas de telefone diferentes por mês")
    sheets = {
        "Jan": [["Nome", "Telefone", "Valor"], ["A", "11988887777", "100"], ["A2", "11955554444", "100"]],
        "Fev": [["Nome", "Celular", "Valor"], ["B", "21999990000", "200"], ["B2", "21955554444", "200"]],
        "Mar": [["Nome", "Fone", "Valor"], ["C", "31988887777", "300"], ["C2", "31955554444", "300"]],
    }
    sales = load_wb(sheets)
    sp = app.detect_phone_col(sales)
    kommo = pd.DataFrame({"Celular": ["11988887777", "21999990000", "31988887777"],
                          "Tags": ["trafego", "trafego", "trafego"]})
    bd = app.run_breakdown_by_sheet(sales, sp, kommo, "Celular", "Tags", "trafego")
    ck("cada mês cruza apesar da coluna de telefone ter nome diferente (1/mês)",
       bd is not None and int(bd["Vendas de Tráfego"].sum()) == 3,
       None if bd is None else str(bd[["Mês / Aba", "Vendas de Tráfego"]].values.tolist()))


# ════════════════════════════════════════════════════════════════════════════
# M5. Recorrente com FORMATO de telefone diferente a cada mês
# ════════════════════════════════════════════════════════════════════════════
def test_formato_variando():
    cat("M5. Recorrente com formato variando entre meses")
    fmts = ["11988887777", "(11) 98888-7777", "+55 11 98888-7777", "11 9 8888-7777",
            "011 98888-7777", "11.98888.7777", "55 11 98888 7777", "11988887777.0"]
    sheets = {}
    for i, m in enumerate(MESES):
        sheets[m] = [["Nome", "Telefone", "Data", "Valor"],
                     ["Recorrente", fmts[i], f"15/{i+1:02d}/2026", "100,00"]]
    sales = load_wb(sheets)
    kommo = pd.DataFrame({"Celular": ["11988887777"], "Tags": ["trafego"]})
    bd = app.run_breakdown_by_sheet(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("recorrente reconhecido em todos os 8 meses apesar dos formatos",
       bd is not None and int(bd["Vendas de Tráfego"].sum()) == 8,
       None if bd is None else str(bd["Vendas de Tráfego"].tolist()))


# ════════════════════════════════════════════════════════════════════════════
# M6. Abas-lixo: mês vazio + aba de resumo no meio
# ════════════════════════════════════════════════════════════════════════════
def test_abas_lixo():
    cat("M6. Mês vazio + aba de Resumo")
    sheets = {
        "Resumo": [["Relatório consolidado"], ["Total de vendas:"], ["1500"]],
        "Jan": [["Nome", "Telefone", "Valor"], ["A", "11988887777", "100"], ["B", "21999990000", "200"]],
        "Fev (sem vendas)": [["Nome", "Telefone", "Valor"]],   # só cabeçalho, sem dados
        "Mar": [["Nome", "Telefone", "Valor"], ["C", "31988887777", "300"]],
    }
    sales = load_wb(sheets)
    ck("carregou sem quebrar com Resumo/mês vazio", sales is not None, None)
    kommo = pd.DataFrame({"Celular": ["11988887777", "21999990000", "31988887777"],
                          "Tags": ["trafego", "trafego", "trafego"]})
    bd = app.run_breakdown_by_sheet(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    # Jan(2) + Mar(1) = 3 conversões; Resumo/Fev não atrapalham
    ck("3 conversões reais (Resumo/mês-vazio não inflam)",
       bd is not None and int(bd["Vendas de Tráfego"].sum()) == 3,
       None if bd is None else str(bd[["Mês / Aba", "Vendas de Tráfego"]].values.tolist()))


# ════════════════════════════════════════════════════════════════════════════
# M7. Seleção de UM mês só (multiselect) → sem _Planilha → breakdown None
# ════════════════════════════════════════════════════════════════════════════
def test_um_mes():
    cat("M7. Selecionar um único mês")
    sales = load_wb(workbook_8_meses(), selected=["Abril"])
    ck("uma aba só → sem coluna _Planilha", sales is not None and "_Planilha" not in sales.columns)
    kommo = pd.DataFrame({"Celular": [RECORRENTE], "Tags": ["trafego"]})
    _, _, traf, _ = app.run_procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("cruza só o mês selecionado (1 conversão)", len(traf) == 1, f"traf={len(traf)}")
    bd = app.run_breakdown_by_sheet(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("breakdown é None para aba única", bd is None)


# ════════════════════════════════════════════════════════════════════════════
# M8. Disparo mês a mês + janela respeitada por mês
# ════════════════════════════════════════════════════════════════════════════
def test_disparo_mes():
    cat("M8. Disparo mês a mês")
    sheets = {}
    for i, m in enumerate(["Abril", "Maio", "Junho"], 4):
        # venda no dia 25; disparo dia 22 do mesmo mês (dentro da janela)
        sheets[m] = [["Nome", "Telefone", "Data", "Valor"],
                     ["Cliente", "66999873776", f"25/{i:02d}/2026", "100,00"]]
    sales = load_wb(sheets)
    # 1 Kommo de disparo, tags com a data de cada mês embutida não dá — usa data fixa de abril:
    kommo = pd.DataFrame({"Celular": ["66999873776"], "Tags": ["recebeu disparo"], "DataDisp": ["22/04/2026"]})
    bd = app.run_breakdown_by_sheet(sales, "Telefone", kommo, "Celular", "Tags", "trafego",
                                    sales_date_col="Data", disparo_keyword="disparo", kommo_date_col="DataDisp")
    ck("breakdown tem coluna de disparo", bd is not None and "Vendas de Disparo" in bd.columns)
    # só Abril está dentro da janela de 30 dias após 22/04; Maio/Junho fora
    by = {r["Mês / Aba"]: r["Vendas de Disparo"] for _, r in bd.iterrows()} if bd is not None else {}
    ck("Abril dentro da janela conta disparo", by.get("Abril") == 1, str(by))
    ck("Maio/Junho fora da janela não contam", by.get("Maio") == 0 and by.get("Junho") == 0, str(by))


# ════════════════════════════════════════════════════════════════════════════
# M9. Mês com título + linha TOTAL (confusão) ainda conta certo
# ════════════════════════════════════════════════════════════════════════════
def test_mes_baguncado():
    cat("M9. Meses bagunçados (título + TOTAL)")
    sheets = {
        "ABRIL 2026": [
            ["Vendas de Abril", None, None],
            [None, None, None],
            ["Cliente", "Telefone", "Valor"],
            ["A", "11988887777", "R$ 100,00"],
            ["B", "21999990000", "R$ 200,00"],
            ["TOTAL", "", "R$ 300,00"],
        ],
        "MAIO 2026": [
            ["Vendas de Maio", None, None],
            ["Cliente", "Telefone", "Valor"],
            ["C", "31988887777", "R$ 150,00"],
            ["SOMA", "", "R$ 150,00"],
        ],
    }
    sales = load_wb(sheets)
    kommo = pd.DataFrame({"Celular": ["11988887777", "21999990000", "31988887777"],
                          "Tags": ["Tráfego Pago", "trafego", "Tráfego Ads"]})
    bd = app.run_breakdown_by_sheet(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    by = {r["Mês / Aba"]: (r["Vendas no mês"], r["Vendas de Tráfego"]) for _, r in bd.iterrows()} if bd is not None else {}
    ck("Abril: 2 vendas (TOTAL fora), 2 tráfego", by.get("ABRIL 2026") == (2, 2), str(by))
    ck("Maio: 1 venda (SOMA fora), 1 tráfego", by.get("MAIO 2026") == (1, 1), str(by))


# ════════════════════════════════════════════════════════════════════════════
# M10. Performance — 8 meses grandes × Kommo grande
# ════════════════════════════════════════════════════════════════════════════
def test_performance():
    cat("M10. Performance — 8 meses grandes")
    per_month = 1000
    ddds = ["11", "21", "31", "41", "51", "61", "71", "81"]
    sheets = {}
    for mi, m in enumerate(MESES):
        rows = [["Nome", "Telefone", "Data", "Valor"]]
        for j in range(per_month):
            rows.append([f"C{mi}_{j}", f"{ddds[mi]}9{(70000000 + j):08d}", f"15/{mi+1:02d}/2026", "100,00"])
        sheets[m] = rows
    sales = load_wb(sheets)
    # Kommo cobre metade de cada mês
    cel = []
    for mi in range(8):
        for j in range(0, per_month, 2):
            cel.append(f"{ddds[mi]}9{(70000000 + j):08d}")
    kommo = pd.DataFrame({"Celular": cel, "Tags": ["trafego"] * len(cel)})
    t0 = time.time()
    bd = app.run_breakdown_by_sheet(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    dt = time.time() - t0
    ck(f"breakdown 8 meses × {per_month} (Kommo {len(cel)}) < 40s", dt < 40, f"{dt:.1f}s")
    ck("cada mês ~500 conversões",
       bd is not None and all(abs(v - per_month // 2) <= 2 for v in bd["Vendas de Tráfego"]),
       None if bd is None else str(bd["Vendas de Tráfego"].tolist()))


# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    for fn in (test_carga, test_recorrente, test_atribuicao, test_layouts_diferentes,
               test_formato_variando, test_abas_lixo, test_um_mes, test_disparo_mes,
               test_mes_baguncado, test_performance):
        try:
            fn()
        except Exception as e:
            import traceback
            _results.append((_cur, fn.__name__ + " (EXCEÇÃO)", False, repr(e)))
            print(f"  [EXCEÇÃO] {fn.__name__}: {e}")
            traceback.print_exc()

    print("\n" + "=" * 64)
    cats = {}
    for c, _, ok, _ in _results:
        p, t = cats.get(c, (0, 0))
        cats[c] = (p + (1 if ok else 0), t + 1)
    for c in sorted(cats):
        p, t = cats[c]
        print(f"  {'✅' if p == t else '❌'} {c}: {p}/{t}")
    fails = [(c, n, d) for c, n, ok, d in _results if not ok]
    total = len(_results)
    print(f"\n  TOTAL: {total - len(fails)}/{total} passaram")
    if fails:
        print("\n  FALHAS:")
        for c, n, d in fails:
            print(f"   • [{c}] {n}" + (f"  → {d}" if d else ""))
        sys.exit(1)
    print("\n✅ Bateria multi-mês passou.")
