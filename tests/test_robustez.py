"""
Bateria de ROBUSTEZ — Proc Aure.
Tortura a percepção e a tratativa de dados: telefones sujos (pontos, vírgulas,
traços, ruído, tipos nativos do Excel), valores BR/US, datas variadas, confusões
de coluna, encodings, planilhas bagunçadas e a "planilha do inferno".

Rodar:  python3 tests/test_robustez.py
"""
import io
import os
import sys
import warnings
from datetime import datetime

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


def make_xlsx(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, rows in sheets.items():
            pd.DataFrame(rows).to_excel(w, sheet_name=name, header=False, index=False)
    return buf.getvalue()


def load_xlsx(sheets):
    b = make_xlsx(sheets)
    return app.load_file_multisheet(app._FileLike(b, "t.xlsx"), app.get_excel_sheets(app._FileLike(b, "t.xlsx")))[0]


def load_bytes_csv(data: bytes):
    return app.load_file_multisheet(app._FileLike(data, "t.csv"), [])[0]


# ════════════════════════════════════════════════════════════════════════════
# R1. Telefone — separadores, ruído, formatos
# ════════════════════════════════════════════════════════════════════════════
def test_telefone_sujo():
    cat("R1. Telefone sujo — separadores e ruído")
    # todos devem normalizar para o celular 11 98888-7777 → (11, 88887777)
    iguais = [
        "11988887777", "11.98888.7777", "11-98888-7777", "11/98888/7777",
        "11,98888,7777", "(11) 98888-7777", "(11)98888-7777", "11 98888 7777",
        "11 9 8888-7777", "  11988887777  ", "Tel: 11988887777", "cel 11 98888-7777",
        "whatsapp: (11)98888-7777", "📱 11 98888-7777", "'11988887777",
        "+55 11 98888-7777", "55 11 98888-7777", "011 98888-7777",
        "11988887777.0", " 11988887777\t", "11 98888–7777",
    ]
    for raw in iguais:
        ck(f"{raw!r} → (11,88887777)", app.phone_key(raw) == ("11", "88887777"), str(app.phone_key(raw)))

    # fixo com separadores
    for raw in ["1133334444", "11 3333-4444", "(11) 3333-4444", "11.3333.4444"]:
        ck(f"fixo {raw!r} → (11,33334444)", app.phone_key(raw) == ("11", "33334444"), str(app.phone_key(raw)))

    # múltiplos números na mesma célula
    ks = app._phone_keys_in_cell("casa 11988887777, trabalho 11 3333-4444")
    subs = sorted(k[1] for k in ks)
    ck("2 números (cel+fixo) extraídos da célula", subs == ["33334444", "88887777"], str(subs))


# ════════════════════════════════════════════════════════════════════════════
# R2. Telefone — o que NÃO é telefone
# ════════════════════════════════════════════════════════════════════════════
def test_nao_telefone():
    cat("R2. Não-telefone — CPF/CNPJ/CEP/data/ID")
    nao = {
        "CPF 3º!=9": "123.456.789-00",
        "CNPJ": "12.345.678/0001-99",
        "data ISO": "2026-03-15",
        "data BR": "15/03/2026",
        "DDD inválido 00": "0099998888",
        "DDD inválido 10": "1099998888",
        "sequência repetida": "11111111111",
        "muito curto": "1234",
        "muito longo": "1198888777766554433",
    }
    for nome, v in nao.items():
        ck(f"{nome} NÃO é telefone", not app._looks_like_phone(v), f"{v} -> {app.phone_key(v)}")
    # estes SÃO
    for v in ["66999873776", "1133334444", "988887777"]:
        ck(f"{v} é telefone", app._looks_like_phone(v))


# ════════════════════════════════════════════════════════════════════════════
# R3. parse_value — moedas e formatos
# ════════════════════════════════════════════════════════════════════════════
def test_valores():
    cat("R3. parse_value — BR, US, R$, negativos, milhar")
    casos = {
        "1.234,56": 1234.56, "R$ 1.234,56": 1234.56, "R$1.234,56": 1234.56,
        "1,234.56": 1234.56, "89,90": 89.90, "1234": 1234.0, "2.500": 2500.0,
        "R$ 2.500": 2500.0, "2500.00": 2500.0, "199.90": 199.90,
        "-50,00": -50.0, "(50,00)": -50.0, "1.234.567,89": 1234567.89,
        "0": 0.0, "": 0.0, "R$ 0,00": 0.0,
    }
    for s, exp in casos.items():
        got = app.parse_value(s)
        ck(f"parse_value({s!r})={exp}", abs(got - exp) < 0.005, str(got))


# ════════════════════════════════════════════════════════════════════════════
# R4. parse_date — variados e bagunçados
# ════════════════════════════════════════════════════════════════════════════
def test_datas():
    cat("R4. parse_date — timestamps, extenso, misto")
    def ymd(s, y, m, d):
        dt = app.parse_date(s)
        ck(f"{s!r} → {y}-{m:02d}-{d:02d}", dt is not None and (dt.year, dt.month, dt.day) == (y, m, d), str(dt))
    ymd("22/04/2026 17:42:10", 2026, 4, 22)        # timestamp com segundos
    ymd("2026-04-22 17:42:10", 2026, 4, 22)
    ymd("22/04/2026", 2026, 4, 22)
    ymd("2026-04-22", 2026, 4, 22)
    ymd("22 de abril de 2026", 2026, 4, 22)
    ymd("22-abr-2026", 2026, 4, 22)
    ymd("abril/2026", 2026, 4, 1)
    ymd("04/22/2026", 2026, 4, 22)                  # US: dia 22 inválido como mês → inverte
    dt = datetime(2026, 4, 22, 9, 30)
    ck("objeto datetime nativo", app.parse_date(dt) is not None and app.parse_date(dt).day == 22)

    # coluna com formatos MISTOS ainda é detectada como data
    df = pd.DataFrame({"Quando": ["15/03/2026", "2026-03-16", "17 de março de 2026", "45292", "18/03/2026"]})
    ck("detect_date_col em coluna de formatos mistos", app.detect_date_col(df) == "Quando", str(app.detect_date_col(df)))


# ════════════════════════════════════════════════════════════════════════════
# R5. Percepção — confusão entre tipos de coluna
# ════════════════════════════════════════════════════════════════════════════
def test_percepcao():
    cat("R5. Percepção — CPF×CEP×Telefone×Data×Valor juntos")
    df = pd.DataFrame({
        "CPF": ["123.456.789-00", "987.654.321-11", "111.222.333-96", "555.666.777-22"],
        "CEP": ["01310-100", "20040-002", "30130-010", "40020-000"],
        "Celular": ["11988887777", "21999990000", "31988887777", "41977776666"],
        "Nascimento": ["10/05/1990", "22/11/1985", "03/03/2000", "15/07/1978"],
        "Valor": ["199,90", "1.250,00", "89,90", "450,00"],
    })
    ck("phone_col = Celular", app.detect_phone_col(df) == "Celular", str(app.detect_phone_col(df)))
    ck("date_col = Nascimento", app.detect_date_col(df) == "Nascimento", str(app.detect_date_col(df)))
    ck("value_col = Valor", app.detect_value_col(df) == "Valor", str(app.detect_value_col(df)))

    # quantidade (inteiros pequenos) NÃO deve ganhar de preço com decimais
    df2 = pd.DataFrame({
        "Qtd": ["1", "2", "1", "3", "2"],
        "Preco": ["199,90", "89,90", "450,00", "1.200,00", "59,90"],
    })
    ck("value_col prefere Preco a Qtd", app.detect_value_col(df2) == "Preco", str(app.detect_value_col(df2)))

    # duas colunas de telefone reais: detecta uma, mas lookup usa as duas
    df3 = pd.DataFrame({
        "Nome": ["A", "B"], "Celular": ["11988887777", "21999990000"],
        "WhatsApp": ["31988887777", "41977776666"], "Valor": ["1", "2"],
    })
    ck("phone_col detectada entre Celular/WhatsApp",
       app.detect_phone_col(df3) in ("Celular", "WhatsApp"), str(app.detect_phone_col(df3)))


# ════════════════════════════════════════════════════════════════════════════
# R6. Ingestão — tipos NATIVOS do Excel (int, datetime, float)
# ════════════════════════════════════════════════════════════════════════════
def test_tipos_nativos():
    cat("R6. Ingestão — tipos nativos do Excel")
    df = load_xlsx({"V": [
        ["Nome", "Telefone", "Data", "Valor"],
        ["Joao", 11988887777, datetime(2026, 4, 15), 199.9],
        ["Maria", 21999990000, datetime(2026, 4, 16), 299.0],
    ]})
    ck("carregou 2 linhas", df is not None and len(df) == 2, None if df is None else f"n={len(df)}")
    if df is not None:
        ck("telefone nativo (int) vira chave certa", app.phone_key(df.iloc[0]["Telefone"]) == ("11", "88887777"),
           str(app.phone_key(df.iloc[0]["Telefone"])))
        ck("data nativa parseada", app.parse_date(df.iloc[0]["Data"]) is not None)
        # end-to-end com tipos nativos
        kommo = pd.DataFrame({"Celular": ["11988887777", "21999990000"], "Tags": ["trafego", "trafego"]})
        _, _, traf, _ = app.run_procv(df, "Telefone", kommo, "Celular", "Tags", "trafego")
        ck("PROCV casa telefones nativos do Excel", len(traf) == 2, f"traf={len(traf)}")


# ════════════════════════════════════════════════════════════════════════════
# R7. Ingestão — encodings, BOM, separadores, aspas
# ════════════════════════════════════════════════════════════════════════════
def test_encodings_csv():
    cat("R7. CSV — encoding, BOM, separador, aspas")
    # latin-1 com acentos
    df = load_bytes_csv("Nome;Telefone\nJoão Conceição;11988887777\nÂngela;21999990000".encode("latin-1"))
    ck("latin-1 acentos preservados",
       df is not None and "João Conceição" in df.iloc[:, 0].astype(str).tolist(),
       None if df is None else str(df.iloc[:, 0].tolist()))

    # BOM utf-8
    df = load_bytes_csv("﻿Nome,Telefone\nAna,11988887777".encode("utf-8"))
    ck("BOM utf-8 removido do header",
       df is not None and "Nome" in [str(c) for c in df.columns],
       None if df is None else str(list(df.columns)))

    # tab
    df = load_bytes_csv("Nome\tTelefone\nAna\t11988887777\nBia\t21999990000".encode("utf-8"))
    ck("CSV separado por TAB", df is not None and "Telefone" in df.columns and len(df) == 2,
       None if df is None else f"cols={list(df.columns)} n={len(df)}")

    # vírgula dentro de aspas
    df = load_bytes_csv('Nome,Telefone\n"Silva, Joao",11988887777\n"Souza, Ana",21999990000'.encode("utf-8"))
    ck("vírgula dentro de aspas não quebra colunas",
       df is not None and len(df) == 2 and "Silva, Joao" in df.iloc[:, 0].astype(str).tolist(),
       None if df is None else f"n={len(df)} c0={df.iloc[:,0].tolist() if df is not None else None}")


# ════════════════════════════════════════════════════════════════════════════
# R8. Ingestão — linhas-lixo (TOTAL, branco no meio)
# ════════════════════════════════════════════════════════════════════════════
def test_linhas_lixo():
    cat("R8. Linhas-lixo — TOTAL / branco no meio")
    df = load_xlsx({"V": [
        ["Nome", "Telefone", "Valor"],
        ["Joao", "11988887777", "100,00"],
        [None, None, None],                       # linha em branco no meio
        ["Maria", "21999990000", "200,00"],
        ["TOTAL", "", "300,00"],                   # linha de soma
    ]})
    nomes = [] if df is None else df["Nome"].astype(str).tolist()
    ck("linha em branco e TOTAL removidas (só 2 vendas)",
       df is not None and "TOTAL" not in nomes and len(df) == 2,
       None if df is None else f"nomes={nomes} n={len(df)}")


# ════════════════════════════════════════════════════════════════════════════
# R9. Header — título multi-linha
# ════════════════════════════════════════════════════════════════════════════
def test_header_dificil():
    cat("R9. Header — título e ruído antes do cabeçalho")
    df = load_xlsx({"V": [
        ["RELATÓRIO DE VENDAS", None, None],
        ["Período: Abril/2026", None, None],
        [None, None, None],
        ["Nome", "Telefone", "Valor"],
        ["Joao", "11988887777", "100"],
        ["Maria", "21999990000", "200"],
    ]})
    ck("título de 2 linhas ignorado, header certo",
       df is not None and list(df.columns)[:3] == ["Nome", "Telefone", "Valor"] and len(df) == 2,
       None if df is None else f"cols={list(df.columns)} n={len(df)}")


# ════════════════════════════════════════════════════════════════════════════
# R10. PROCV — confusões reais
# ════════════════════════════════════════════════════════════════════════════
def test_procv_confuso():
    cat("R10. PROCV — formatos divergentes e nomes")
    # vendas com telefone formatado, kommo cru — devem casar
    sales = pd.DataFrame({"Nome": ["Joao", "Maria"],
                          "Telefone": ["(11) 98888-7777", "+55 21 99999-0000"], "Valor": ["1", "2"]})
    kommo = pd.DataFrame({"Celular": ["11988887777", "21999990000"], "Tags": ["trafego", "trafego"]})
    _, _, traf, _ = app.run_procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego")
    ck("formatos divergentes (vendas formatado × kommo cru) casam", len(traf) == 2, f"traf={len(traf)}")

    # nome com título "Dr." deve casar com nome sem título
    sales = pd.DataFrame({"Nome": ["Dr. João Carlos Silva"], "Telefone": ["11988887777"], "Valor": ["1"]})
    kommo = pd.DataFrame({"Celular": ["99999999999"], "Tags": ["trafego"], "Nome": ["João Carlos Silva"]})
    _, _, traf, _ = app.run_procv(sales, "Telefone", kommo, "Celular", "Tags", "trafego",
                                  sales_name_col="Nome", kommo_name_col="Nome")
    ck("nome com título 'Dr.' casa com nome sem título", len(traf) == 1, f"traf={len(traf)}")


# ════════════════════════════════════════════════════════════════════════════
# R11. Disparo — confusões de data/tag
# ════════════════════════════════════════════════════════════════════════════
def test_disparo_confuso():
    cat("R11. Disparo — datas/tags bagunçadas")
    sales = pd.DataFrame({"Nome": ["C"], "Telefone": ["66999873776"],
                          "Data": ["25/04/2026 14:30:00"], "Valor": ["100"]})
    # data do disparo embutida em texto bem bagunçado
    k = pd.DataFrame({"Celular": ["66999873776"],
                      "Tags": ["Cliente VIP, recebeu disparo Dia das Mães em 22/04, promo"]})
    r = app.run_disparo(sales, "Telefone", "Data", k, "Celular", "Tags", "disparo", None)
    ck("data do disparo extraída de texto bagunçado", int((r["Venda_Confirmada"] == "SIM").sum()) == 1,
       f"conf={int((r['Venda_Confirmada']=='SIM').sum()) if len(r) else 0}")

    # janela EXATA: venda no limite (30 dias) conta; 31 não
    k2 = pd.DataFrame({"Celular": ["66999873776"], "Tags": ["disparo 01/04/2026"]})
    r30 = app.run_disparo(pd.DataFrame({"Nome": ["C"], "Telefone": ["66999873776"], "Data": ["01/05/2026"], "Valor": ["1"]}),
                          "Telefone", "Data", k2, "Celular", "Tags", "disparo", None)
    ck("venda exatamente 30 dias após = conta", int((r30["Venda_Confirmada"] == "SIM").sum()) == 1)
    r31 = app.run_disparo(pd.DataFrame({"Nome": ["C"], "Telefone": ["66999873776"], "Data": ["02/05/2026"], "Valor": ["1"]}),
                          "Telefone", "Data", k2, "Celular", "Tags", "disparo", None)
    ck("venda 31 dias após = NÃO conta", int((r31["Venda_Confirmada"] == "SIM").sum()) == 0)


# ════════════════════════════════════════════════════════════════════════════
# R12. Planilha do inferno — tudo junto
# ════════════════════════════════════════════════════════════════════════════
def test_planilha_do_inferno():
    cat("R12. Planilha do inferno — tudo junto")
    sales = load_xlsx({"Vendas 2026": [
        ["LOJA AURE — VENDAS", None, None, None, None],
        ["Atualizado em 31/05/2026", None, None, None, None],
        [None, None, None, None, None],
        ["Cliente", "WhatsApp", "CPF", "Data da Compra", "Valor (R$)"],
        ["João da Silva", "(11) 98888-7777", "123.456.789-00", "15/04/2026", "R$ 1.299,90"],
        ["Maria Souza", "55 21 99999-0000", "987.654.321-11", "10/04/2026", "299,00"],
        ["Pedro", 66999873776, "111.222.333-96", datetime(2026, 4, 20), 150.0],   # nativos
        ["Ana", "", "444.555.666-77", "26/04/2026", "99,90"],                      # sem tel
        ["TOTAL", "", "", "", "R$ 1.848,80"],                                      # soma
    ], "Cadastro Antigo": [
        ["Nome", "Fone"],
        ["Ana Lima", "31 98888-7777"],   # telefone da Ana num cadastro separado
    ]})
    # Ana sem telefone nas vendas, mas tem na aba Cadastro (31 98888-7777)
    kommo = pd.DataFrame({
        "Celular": ["11988887777", "21999990000", "66999873776", "31988887777"],
        "Tags": ["Tráfego Pago", "tráfego", "recebeu disparo 18/04", "Tráfego Ads"],
        "Nome": ["João da Silva", "Maria Souza", "Pedro", "Ana Lima"],
    })
    ck("inferno: carregou vendas (sem TOTAL)", sales is not None, None)
    if sales is None:
        return
    nomes = sales[[c for c in sales.columns if app._canon_colname(c) == "cliente"][0]].astype(str).tolist() \
        if any(app._canon_colname(c) == "cliente" for c in sales.columns) else []
    ck("inferno: linha TOTAL fora", "TOTAL" not in nomes, str(nomes))

    sphone = [c for c in sales.columns if app._canon_colname(c) == "whatsapp"]
    sphone = sphone[0] if sphone else "WhatsApp"
    _, _, traf, full = app.run_procv(sales, sphone, kommo, "Celular", "Tags", "trafego")
    # João, Maria, Ana(via cadastro), e Pedro NÃO é tráfego (é disparo)
    ck("inferno: 3 conversões de tráfego (inclui Ana via cadastro)", len(traf) == 3,
       f"traf={len(traf)} criterios={list(full[full['Venda_Confirmada']=='SIM']['Criterio_Match'])}")

    disp = app.run_disparo(sales, sphone, "Data da Compra" if "Data da Compra" in sales.columns
                           else [c for c in sales.columns if "data" in app._canon_colname(c)][0],
                           kommo, "Celular", "Tags", "disparo", None)
    ck("inferno: 1 conversão de disparo (Pedro)", int((disp["Venda_Confirmada"] == "SIM").sum()) == 1,
       f"conf={int((disp['Venda_Confirmada']=='SIM').sum()) if len(disp) else 0}")


# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    for fn in (test_telefone_sujo, test_nao_telefone, test_valores, test_datas,
               test_percepcao, test_tipos_nativos, test_encodings_csv, test_linhas_lixo,
               test_header_dificil, test_procv_confuso, test_disparo_confuso,
               test_planilha_do_inferno):
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
    print("\n✅ Bateria de robustez passou.")
