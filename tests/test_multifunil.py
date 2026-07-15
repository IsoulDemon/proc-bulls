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


def test_listas_de_vendas():
    cat("N10. Várias LISTAS de vendas (uma por vendedora)")
    # 3 vendedoras; a da Bia usa 'Celular' em vez de 'Telefone' (unificação)
    ana = pd.DataFrame({"Cliente": ["A1", "A2"],
                        "Telefone": ["11988880001", "11988880002"],
                        "Valor": ["100,00", "200,00"]})
    bia = pd.DataFrame({"Cliente": ["B1", "B2"],
                        "Celular": ["21988880003", "21988880005"],
                        "Valor": ["300,00", "150,00"]})
    ca = pd.DataFrame({"Cliente": ["C1", "A1 de novo"],
                       "Telefone": ["31988880004", "11988880001"],  # A1 repetida!
                       "Valor": ["50,00", "80,00"]})
    comb = app.combine_sales_sources([ana, bia, ca], ["Ana.xlsx", "Bia.csv", "Carla.xlsx"])
    ck("combina 3 listas (6 vendas)", comb is not None and len(comb) == 6,
       None if comb is None else f"n={len(comb)}")
    ck("_Planilha marca a lista de origem",
       sorted(comb["_Planilha"].unique()) == ["Ana", "Bia", "Carla"],
       str(sorted(comb["_Planilha"].unique())))
    tel_col = app.detect_phone_col(comb)
    ck("coluna de telefone unificada entre as listas",
       tel_col is not None and int(comb[tel_col].notna().sum()) == 6,
       f"col={tel_col} preenchidos={int(comb[tel_col].notna().sum()) if tel_col else 0}")

    kommo = pd.DataFrame({
        "Celular": ["11988880001", "21988880003", "31988880004"],
        "Tags": ["TRAFEGO", "TRAFEGO", "INSTAGRAM"],
    })
    # Total combinado: A1 (repetida em 2 listas) conta UMA vez
    _, _, traf, _ = app.run_procv(comb, tel_col, kommo, "Celular", "Tags", "trafego")
    ck("total combinado dedup entre listas (A1 conta 1×)", len(traf) == 2, f"traf={len(traf)}")

    # Breakdown por lista: Ana=1 (A1), Bia=1, Carla=1 (A1 conta na lista dela também)
    bd = app.run_breakdown_by_sheet(comb, tel_col, kommo, "Celular", "Tags", "trafego")
    by = {r["Mês / Aba"]: r["Vendas de Tráfego"] for _, r in bd.iterrows()} if bd is not None else {}
    ck("breakdown por vendedora (Ana=1, Bia=1, Carla=1)",
       by == {"Ana": 1, "Bia": 1, "Carla": 1}, str(by))


def test_grupo_vip():
    cat("N11. Grupo VIP — venda com telefone na lista, sem janela de data")
    # A engine é a MESMA do tráfego, só sem tag e sem data.
    vendas = pd.DataFrame({
        "Cliente": ["Bolsa De Couro", "Fulano", "Ciclana"],
        "Telefone": ["+5593991848744", "(11) 3333-2222", "(21) 97777-6666"],
        "Data": ["05/01/2026", "10/06/2026", "20/06/2026"],  # datas espalhadas
        "Valor": ["3048,00", "100,00", "80,00"],
    })
    # membro VIP com o MESMO número em formato diferente (sem DDI, com traço)
    vip = pd.DataFrame({"Nome": ["Membro A"], "Telefone": ["93 99184-8744"]})
    _, _, conv, full = app.run_vip(vendas, "Telefone", vip, "Telefone", sales_name_col="Cliente")
    ck("Bolsa De Couro casa por telefone (sem depender de data)",
       len(conv) == 1 and conv.iloc[0]["Criterio_Match"] == "Telefone", f"conv={len(conv)}")
    ck("valor da venda VIP preservado (3048)",
       len(conv) == 1 and "3048" in str(conv.iloc[0]["[Venda] Valor"]),
       str(conv.iloc[0]["[Venda] Valor"]) if len(conv) else "—")
    ck("colunas do VIP renomeadas (Telefone_VIP/Nome_VIP, sem Tag_Kommo)",
       "Telefone_VIP" in conv.columns and "Tag_Kommo" not in conv.columns,
       str([c for c in conv.columns if not c.startswith("[Venda]")]))

    # Dedup: mesma pessoa com 2 números no VIP e nas vendas → 1 conversão
    vendas2 = pd.DataFrame({"Cliente": ["Ana"], "Fixo": ["(31) 3333-4444"], "Cel": ["(31) 98888-7777"]})
    vip2 = pd.DataFrame({"Tel": ["3133334444", "31988887777"]})
    _, _, conv2, _ = app.run_vip(vendas2, "Fixo", vip2, "Tel")
    ck("dedup por comprador/venda no VIP (2 números → 1)", len(conv2) == 1, f"conv={len(conv2)}")

    # DDD divergente NÃO casa (mesma trava anti-falso-positivo)
    vendas3 = pd.DataFrame({"Cliente": ["X"], "Tel": ["(11) 98888-7777"]})
    vip3 = pd.DataFrame({"Tel": ["(21) 98888-7777"]})
    _, _, conv3, _ = app.run_vip(vendas3, "Tel", vip3, "Tel")
    ck("DDD divergente não casa no VIP", len(conv3) == 0, f"conv={len(conv3)}")

    # Resgate de telefone escondido na lista VIP (fora da coluna de telefone)
    vendas4 = pd.DataFrame({"Cliente": ["Ana", "Bia"], "Tel": ["(11) 98888-7777", "(21) 97777-6666"]})
    vip4 = pd.DataFrame({"Celular": ["", ""], "Obs": ["p:+5511988887777", "sem numero"]})
    _, _, conv4, _ = app.run_vip(vendas4, "Tel", vip4, "Celular")
    ck("telefone escondido na lista VIP é resgatado", len(conv4) == 1, f"conv={len(conv4)}")

    # Overlap 3 canais: comprador em Tráfego E VIP
    vendas5 = pd.DataFrame({"Cliente": ["Ana", "Bia"], "Tel": ["11988887777", "21999990000"]})
    kommo5 = pd.DataFrame({"Celular": ["11988887777"], "Tags": ["TRAFEGO"]})  # Ana = tráfego
    vip5 = pd.DataFrame({"Tel": ["11988887777", "21999990000"]})              # Ana E Bia = VIP
    _, _, traf5, _ = app.run_procv(vendas5, "Tel", kommo5, "Celular", "Tags", "trafego")
    _, _, vipc5, _ = app.run_vip(vendas5, "Tel", vip5, "Tel")
    tset = {v for v in traf5["Tel_8dig"] if v}
    vset = {v for v in vipc5["Tel_8dig"] if v}
    ck("overlap T∩V = 1 e únicos = 2 (não conta 2×)",
       len(tset & vset) == 1 and len(tset | vset) == 2,
       f"T∩V={len(tset & vset)} uniq={len(tset | vset)}")

    # Breakdown por lista com coluna VIP
    ana = pd.DataFrame({"Cliente": ["A1", "A2"], "Telefone": ["11988880001", "11988880002"]})
    bia = pd.DataFrame({"Cliente": ["B1"], "Telefone": ["21988880003"]})
    comb = app.combine_sales_sources([ana, bia], ["Ana.xlsx", "Bia.csv"])
    tel = app.detect_phone_col(comb)
    vipb = pd.DataFrame({"Tel": ["21988880003", "11988880002"]})  # B1 e A2 são VIP
    bd = app.run_breakdown_by_sheet(comb, tel, pd.DataFrame({"Celular": ["11988880001"], "Tags": ["TRAFEGO"]}),
                                    "Celular", "Tags", "trafego", df_vip=vipb, vip_phone_col="Tel")
    by_vip = {r["Mês / Aba"]: r["Vendas Grupo VIP"] for _, r in bd.iterrows()} if bd is not None else {}
    ck("breakdown mostra VIP por vendedora (Ana=1, Bia=1)",
       by_vip == {"Ana": 1, "Bia": 1}, str(by_vip))


# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    for fn in (test_combinar, test_lead_repetido, test_colunas_diferentes, test_tag_em_um_funil,
               test_cenario_completo, test_bordas, test_disparo_e_exclusao, test_performance,
               test_listas_de_vendas, test_grupo_vip):
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
