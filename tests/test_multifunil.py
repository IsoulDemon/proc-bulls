"""
Bateria MULTI-FUNIL — Proc Aure.
Cenário real do usuário: 1 planilha de vendas com várias páginas (mês a mês) E
VÁRIOS arquivos do Kommo (um por funil de venda). A ferramenta combina os funis
num só Kommo e cruza com cada mês.

Rodar:  python3 tests/test_multifunil.py
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
    print(f"\n── {c} " + "─" * max(2, 56 - len(c)))


def ck(nome, cond, detalhe=""):
    ok = bool(cond)
    _results.append((_cur, nome, ok, detalhe))
    print(f"  [{'PASS' if ok else 'FALHOU'}] {nome}" + (f"  → {detalhe}" if not ok else ""))


def load_wb(sheets: dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, rows in sheets.items():
            pd.DataFrame(rows).to_excel(w, sheet_name=name, header=False, index=False)
    b = buf.getvalue()
    return app.load_file_multisheet(app._FileLike(b, "v.xlsx"), app.get_excel_sheets(app._FileLike(b, "v.xlsx")))[0]


# ════════════════════════════════════════════════════════════════════════════
# N1. Combinar funis
# ════════════════════════════════════════════════════════════════════════════
def test_combinar():
    cat("N1. Combinar vários funis num só Kommo")
    f1 = pd.DataFrame({"Nome": ["A", "B"], "Celular": ["11988887777", "21999990000"], "Tags": ["trafego", "x"]})
    f2 = pd.DataFrame({"Nome": ["C", "D"], "Celular": ["31988887777", "41977776666"], "Tags": ["trafego", "y"]})
    f3 = pd.DataFrame({"Nome": ["E"], "Celular": ["51966665555"], "Tags": ["trafego"]})
    c = app.combine_kommo_sources([f1, f2, f3], ["Funil Vendas", "Funil SDR", "Funil Recompra"])
    ck("combinou 5 leads", c is not None and len(c) == 5, None if c is None else f"n={len(c)}")
    ck("marca a origem em _Funil", c is not None and "_Funil" in c.columns and c["_Funil"].nunique() == 3,
       None if c is None else str(c["_Funil"].unique()))


# ════════════════════════════════════════════════════════════════════════════
# N2. Lead repetido entre funis não dobra a conversão
# ════════════════════════════════════════════════════════════════════════════
def test_lead_repetido():
    cat("N2. Lead em 2 funis = 1 conversão")
    f1 = pd.DataFrame({"Nome": ["Joao"], "Celular": ["11988887777"], "Tags": ["trafego"]})
    f2 = pd.DataFrame({"Nome": ["Joao"], "Celular": ["11988887777"], "Tags": ["trafego pago"]})  # mesmo lead
    kommo = app.combine_kommo_sources([f1, f2], ["A", "B"])
    sales = pd.DataFrame({"Nome": ["Joao"], "Telefone": ["11988887777"], "Valor": ["100"]})
    _, _, traf, _ = app.run_procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("comprador em 2 funis conta 1 vez", len(traf) == 1, f"traf={len(traf)}")


# ════════════════════════════════════════════════════════════════════════════
# N3. Funis com NOMES DE COLUNA diferentes (unificação)
# ════════════════════════════════════════════════════════════════════════════
def test_colunas_diferentes():
    cat("N3. Funis com nomes de coluna diferentes")
    fA = pd.DataFrame({"Nome": ["A", "A2"], "Celular": ["11988887777", "11955554444"], "Tags": ["trafego", "trafego"]})
    fB = pd.DataFrame({"Nome": ["B", "B2"], "Telefone": ["21999990000", "21955554444"], "Etiquetas": ["trafego pago", "trafego"]})
    kommo = app.combine_kommo_sources([fA, fB], ["A", "B"])
    ck("coluna de telefone unificada (sem 'Telefone' órfã)",
       kommo is not None and "Telefone" not in kommo.columns, None if kommo is None else str(list(kommo.columns)))
    sphone = app.detect_phone_col(kommo)
    stag = app.detect_tag_col(kommo)
    sales = pd.DataFrame({"Nome": ["a", "a2", "b", "b2"],
                          "Telefone": ["11988887777", "11955554444", "21999990000", "21955554444"],
                          "Valor": ["1", "1", "1", "1"]})
    _, _, traf, _ = app.run_procv(sales, "Telefone", kommo, sphone, stag, "trafego")
    ck("cruza leads dos 2 funis apesar de nomes diferentes", len(traf) == 4, f"traf={len(traf)} phone={sphone} tag={stag}")


# ════════════════════════════════════════════════════════════════════════════
# N4. Tag de tráfego em apenas um dos funis
# ════════════════════════════════════════════════════════════════════════════
def test_tag_em_um_funil():
    cat("N4. Tag de tráfego só num funil")
    # mesmo lead: no funil A está como 'lead frio', no funil B como 'trafego'
    fA = pd.DataFrame({"Nome": ["Joao", "Ana"], "Celular": ["11988887777", "21999990000"], "Tags": ["lead frio", "lead frio"]})
    fB = pd.DataFrame({"Nome": ["Joao"], "Celular": ["11988887777"], "Tags": ["trafego pago"]})
    kommo = app.combine_kommo_sources([fA, fB], ["A", "B"])
    sales = pd.DataFrame({"Nome": ["Joao"], "Telefone": ["11988887777"], "Valor": ["1"]})
    _, _, traf, _ = app.run_procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("conta porque um funil tem a tag de tráfego", len(traf) == 1, f"traf={len(traf)}")


# ════════════════════════════════════════════════════════════════════════════
# N5. CENÁRIO COMPLETO: vendas multi-mês × Kommo de vários funis
# ════════════════════════════════════════════════════════════════════════════
def test_cenario_completo():
    cat("N5. Completo — vendas 6 meses × 3 funis")
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho"]
    REC = "11988887777"  # recorrente (compra todo mês)
    sheets = {}
    for i, m in enumerate(meses, 1):
        sheets[m] = [
            ["Cliente", "WhatsApp", "Data", "Valor"],
            [f"Único {m}", f"219{i:08d}", f"15/{i:02d}/2026", "100,00"],
            ["Recorrente", REC, f"20/{i:02d}/2026", "200,00"],
        ]
    sales = load_wb(sheets)

    # 3 funis: recorrente no funil 1; os 'únicos' espalhados nos funis 2 e 3
    f1 = pd.DataFrame({"Nome": ["Recorrente"], "Celular": [REC], "Tags": ["trafego pago"]})
    unicos = [f"219{i:08d}" for i in range(1, 7)]
    f2 = pd.DataFrame({"Nome": [f"U{i}" for i in (1, 2, 3)], "Celular": unicos[:3], "Tags": ["trafego"] * 3})
    f3 = pd.DataFrame({"Nome": [f"U{i}" for i in (4, 5, 6)], "Celular": unicos[3:], "Tags": ["trafego"] * 3})
    kommo = app.combine_kommo_sources([f1, f2, f3], ["Vendas", "SDR", "Recompra"])
    ck("Kommo combinado tem 7 leads", kommo is not None and len(kommo) == 7, None if kommo is None else f"n={len(kommo)}")

    sp = app.detect_phone_col(sales)
    bd = app.run_breakdown_by_sheet(sales, sp, kommo, "Celular", "Tags", "trafego")
    by = {r["Mês / Aba"]: r["Vendas de Tráfego"] for _, r in bd.iterrows()} if bd is not None else {}
    # cada mês: recorrente (1) + o único daquele mês (1) = 2
    ck("breakdown: 6 meses", bd is not None and len(bd) == 6, str(by))
    ck("cada mês tem 2 conversões (recorrente + único do mês)",
       bool(by) and all(v == 2 for v in by.values()), str(by))
    ck("total mês a mês = 12 (6 recorrente + 6 únicos)",
       bd is not None and int(bd["Vendas de Tráfego"].sum()) == 12, str(by))


# ════════════════════════════════════════════════════════════════════════════
# N6/N7. Casos de borda
# ════════════════════════════════════════════════════════════════════════════
def test_bordas():
    cat("N6/N7. Um funil só, vazios, None")
    f1 = pd.DataFrame({"Nome": ["A"], "Celular": ["11988887777"], "Tags": ["trafego"]})
    c = app.combine_kommo_sources([f1], ["Único"])
    ck("um funil só → sem coluna _Funil", c is not None and "_Funil" not in c.columns)
    c2 = app.combine_kommo_sources([None, pd.DataFrame(), f1, pd.DataFrame()], ["x", "y", "bom", "z"])
    ck("ignora None/vazios e usa o válido", c2 is not None and len(c2) == 1)
    ck("nenhum válido → None", app.combine_kommo_sources([None, pd.DataFrame()]) is None)


# ════════════════════════════════════════════════════════════════════════════
# N8. Disparo + exclusão de tags sobre o Kommo combinado
# ════════════════════════════════════════════════════════════════════════════
def test_disparo_e_exclusao():
    cat("N8. Disparo e exclusão sobre funis combinados")
    f1 = pd.DataFrame({"Celular": ["66999873776"], "Tags": ["recebeu disparo 22/04"]})
    f2 = pd.DataFrame({"Celular": ["11988887777"], "Tags": ["trafego"]})
    kommo = app.combine_kommo_sources([f1, f2], ["Disparo", "Trafego"])
    sales = pd.DataFrame({"Nome": ["C"], "Telefone": ["66999873776"], "Data": ["25/04/2026"], "Valor": ["1"]})
    r = app.run_disparo(sales, "Telefone", "Data", kommo, "Celular", "Tags", "disparo", None)
    ck("disparo encontra o lead no funil certo", int((r["Venda_Confirmada"] == "SIM").sum()) == 1,
       f"conf={int((r['Venda_Confirmada']=='SIM').sum()) if len(r) else 0}")

    # exclusão: orgânico não conta
    fa = pd.DataFrame({"Celular": ["11988887777"], "Tags": ["Tráfego Pago"]})
    fb = pd.DataFrame({"Celular": ["21999990000"], "Tags": ["Tráfego Orgânico"]})
    kommo2 = app.combine_kommo_sources([fa, fb], ["A", "B"])
    sales2 = pd.DataFrame({"Nome": ["a", "b"], "Telefone": ["11988887777", "21999990000"], "Valor": ["1", "1"]})
    _, _, traf, _ = app.run_procv(sales2, "Telefone", kommo2, "Celular", "Tags", "trafego", traffic_exclude="organico")
    ck("exclusão de 'orgânico' funciona no combinado", len(traf) == 1, f"traf={len(traf)}")


# ════════════════════════════════════════════════════════════════════════════
# N9. Performance
# ════════════════════════════════════════════════════════════════════════════
def test_performance():
    cat("N9. Performance — 5 funis grandes")
    ddds = ["11", "21", "31", "41", "51"]
    funis = []
    for fi in range(5):
        cel = [f"{ddds[fi]}9{(70000000 + j):08d}" for j in range(1000)]
        funis.append(pd.DataFrame({"Celular": cel, "Tags": ["trafego"] * 1000}))
    t0 = time.time()
    kommo = app.combine_kommo_sources(funis, [f"F{i}" for i in range(5)])
    dt = time.time() - t0
    ck(f"combinar 5×1000 leads < 5s", dt < 5, f"{dt:.2f}s")
    ck("combinado tem 5000 leads", kommo is not None and len(kommo) == 5000, None if kommo is None else f"n={len(kommo)}")


# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    for fn in (test_combinar, test_lead_repetido, test_colunas_diferentes, test_tag_em_um_funil,
               test_cenario_completo, test_bordas, test_disparo_e_exclusao, test_performance):
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
    print("\n✅ Bateria multi-funil passou.")
