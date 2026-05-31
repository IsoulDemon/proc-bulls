"""
Testes de regressão do tratamento de planilhas e do cruzamento por telefone.
Cobre os 7 cenários do scan que originaram os erros da ferramenta.

Rodar:  python3 tests/test_tratamento.py
"""
import os
import sys
import warnings

warnings.filterwarnings("ignore")
os.environ["STREAMLIT_GLOBAL_DISABLE_WATCHDOG_WARNING"] = "true"
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd  # noqa: E402

import app  # noqa: E402

_failures = []


def check(name, cond, detail=""):
    status = "PASS" if cond else "FALHOU"
    print(f"  [{status}] {name}" + (f" — {detail}" if detail and not cond else ""))
    if not cond:
        _failures.append(name)


# ── 1. Telefone: sem colisão entre DDDs, mas casa com/sem 9 e com/sem DDD ────
def test_phone_keys():
    print("\n1. Chave de telefone (DDD + sub8)")
    sp = app.phone_key("(11) 3456-7890")
    rj = app.phone_key("(21) 3456-7890")
    check("DDDs diferentes NÃO casam (11 x 21)", not app.phones_match(sp, rj),
          f"{sp} vs {rj}")

    com9 = app.phone_key("66 99987-3776")
    sem9 = app.phone_key("66 9987-3776")
    check("mesmo número com/sem o 9 casa", app.phones_match(com9, sem9),
          f"{com9} vs {sem9}")

    com_ddd = app.phone_key("66999873776")
    sem_ddd = app.phone_key("99873776")
    check("com DDD x sem DDD casa (lado ausente)", app.phones_match(com_ddd, sem_ddd),
          f"{com_ddd} vs {sem_ddd}")


# ── 2. clean_phone: notação científica e placeholder ─────────────────────────
def test_clean_phone_formats():
    print("\n2. Limpeza de formatos estranhos")
    check("notação científica '6.6e+10'", app.clean_phone("6.6e+10") == "66000000000",
          app.clean_phone("6.6e+10"))
    check("float string '66999873776.0'", app.clean_phone("66999873776.0") == "66999873776",
          app.clean_phone("66999873776.0"))
    check("DDI 55 removido", app.clean_phone("5566999873776") == "66999873776",
          app.clean_phone("5566999873776"))
    check("placeholder 66000000000 -> chave vazia", app.phone_key("66000000000")[1] == "",
          str(app.phone_key("66000000000")))
    check("texto+número 'tel: 66 99987-3776'", app.clean_phone("tel: 66 99987-3776") == "66999873776",
          app.clean_phone("tel: 66 99987-3776"))


# ── 3. Detecção de telefone à prova de CPF / data ────────────────────────────
def test_phone_detection():
    print("\n3. Telefone x CPF x data")
    check("data ISO não é telefone", not app._looks_like_phone("2024-01-15"))
    check("CPF '12345678900' não é telefone", not app._looks_like_phone("12345678900"))
    check("celular real é telefone", app._looks_like_phone("66999873776"))
    check("fixo com DDD é telefone", app._looks_like_phone("1133334444"))
    df = pd.DataFrame({
        "data_pedido": ["2024-01-15", "2024-02-20", "2024-03-10", "2024-04-01"],
        "documento_x": ["12345678900", "98765432100", "11122233344", "55566677788"],
    })
    check("detect_phone_col não escolhe data/CPF", app.detect_phone_col(df) is None,
          str(app.detect_phone_col(df)))


# ── 4. Cabeçalho: título acima e planilha sem cabeçalho ──────────────────────
def test_header_detection():
    print("\n4. Detecção de cabeçalho")
    com_titulo = pd.DataFrame([
        ["Relatório de Vendas - Abril 2026", None, None],
        [None, None, None],
        ["Nome", "Telefone", "Valor"],
        ["Joao", "66999873776", "199,90"],
        ["Maria", "11988887777", "299,00"],
    ])
    check("título acima -> header na linha 2", app.detect_header_row(com_titulo) == 2,
          str(app.detect_header_row(com_titulo)))

    sem_header = pd.DataFrame([
        ["Joao", "66999873776", "199,90"],
        ["Maria", "11988887777", "299,00"],
        ["Jose", "21977776666", "150,00"],
    ])
    check("planilha sem cabeçalho -> -1", app.detect_header_row(sem_header) == -1,
          str(app.detect_header_row(sem_header)))


# ── 5. Número fora da coluna: telefone solto numa coluna de texto ────────────
def test_stray_phone():
    print("\n5. Número fora da coluna (telefone solto)")
    df_sales = pd.DataFrame({
        "nome": ["Cliente A", "Cliente B", "Cliente C"],
        "observacao": ["cliente novo", "contato 66999873776", "recompra"],
        "valor": ["100", "200", "300"],
    })
    df_kommo = pd.DataFrame({
        "Celular": ["66999873776", "11000000001"],
        "Tags": ["trafego pago", "outro"],
    })
    _, _, df_traf, df_full = app.run_procv(
        df_sales, "nome", df_kommo, "Celular", "Tags", "trafego")
    # o telefone solto na coluna 'observacao' deve ser encontrado
    achou = (df_full["Venda_Confirmada"] == "SIM").any()
    check("telefone solto fora da coluna é cruzado", bool(achou),
          "nenhum match encontrado")


# ── 6. Planilha dividida em blocos lado a lado ───────────────────────────────
def test_side_by_side_blocks():
    print("\n6. Blocos lado a lado na mesma aba")
    raw = pd.DataFrame([
        ["Nome", "Telefone", None, "Nome", "Telefone"],
        ["Joao", "66999873776", None, "Ana", "31988887777"],
        ["Maria", "11988887777", None, "Bia", "41977776666"],
    ])
    stacked = app._split_side_by_side_blocks(raw)
    check("blocos reempilhados (mais linhas, menos colunas)",
          stacked.shape[1] == 2 and len(stacked) >= 5,
          f"shape={stacked.shape}")


# ── 7. Abas com cabeçalho divergente não fragmentam ──────────────────────────
def test_column_alignment():
    print("\n7. Alinhamento de colunas entre abas")
    df1 = pd.DataFrame({"Telefone": ["66999873776"], "Valor": ["100"]})
    df2 = pd.DataFrame({"Telefone ": ["11988887777"], "Valor": ["200"]})
    aligned, n = app._align_columns([df1, df2])
    combined = pd.concat(aligned, ignore_index=True)
    # 'Telefone' e 'Telefone ' devem virar UMA coluna sem NaN
    tel_cols = [c for c in combined.columns if app._canon_colname(c) == "telefone"]
    check("colunas de telefone unificadas", len(tel_cols) == 1, f"colunas: {list(combined.columns)}")
    if tel_cols:
        check("nenhum telefone perdido (sem NaN)", combined[tel_cols[0]].notna().all())


# ── 8. PROCV: precisão por DDD (sem falso match entre cidades) ───────────────
def test_procv_ddd_precision():
    print("\n8. PROCV — precisão por DDD")
    df_sales = pd.DataFrame({
        "nome": ["Joao SP", "Maria RJ"],
        "telefone": ["11999998888", "21999998888"],  # mesmo final, DDDs diferentes
        "valor": ["100", "200"],
    })
    df_kommo = pd.DataFrame({
        "Celular": ["11999998888", "31999998888"],   # SP bate; 31 só tem o mesmo final
        "Tags": ["trafego pago", "trafego pago"],
    })
    _, _, _, df_full = app.run_procv(
        df_sales, "telefone", df_kommo, "Celular", "Tags", "trafego")
    linha_sp = df_full.iloc[0]
    check("lead SP casa com a venda SP (não a RJ)", linha_sp["[Venda] nome"] == "Joao SP",
          str(linha_sp["[Venda] nome"]))
    linha_31 = df_full.iloc[1]
    check("lead DDD 31 NÃO casa (mesmo final, cidade diferente)",
          linha_31["Venda_Confirmada"] == "NÃO", str(linha_31["Venda_Confirmada"]))
    check("contador de falsos matches evitados > 0",
          df_full.attrs.get("n_ddd_blocked", 0) >= 1,
          str(df_full.attrs.get("n_ddd_blocked", 0)))


# ── 9. Duplicatas: DDDs diferentes não viram duplicata ───────────────────────
def test_duplicates_ddd():
    print("\n9. Duplicatas — DDD separa pessoas diferentes")
    df = pd.DataFrame({
        "telefone": ["11999998888", "21999998888"],  # mesmo final, pessoas diferentes
        "nome": ["Joao", "Maria"],
    })
    res = app.analyze_duplicates(df, "telefone")
    check("DDDs diferentes ficam 'Única' (não duplicata)",
          (res["Situacao_Venda"] == "Única").all(),
          str(list(res["Situacao_Venda"])))


# ── 10. Disparo: janela de tempo + precisão por DDD ──────────────────────────
def test_disparo_window():
    print("\n10. Disparo — janela de tempo e DDD")
    df_sales = pd.DataFrame({
        "nome": ["Joao SP", "Maria RJ"],
        "telefone": ["11999998888", "21999998888"],
        "data": ["25/04/2026", "25/04/2026"],
        "valor": ["100", "200"],
    })
    df_kommo = pd.DataFrame({
        "Celular": ["11999998888"],
        "Tags": ["recebeu disparo dia das mães 22/04"],
    })
    res = app.run_disparo(
        df_sales, "telefone", "data", df_kommo, "Celular", "Tags", "disparo", None)
    confirmadas = res[res["Venda_Confirmada"] == "SIM"]
    check("disparo casa a venda SP dentro da janela", len(confirmadas) == 1, str(len(confirmadas)))
    if len(confirmadas):
        check("disparo não casa a venda RJ (DDD diferente)",
              confirmadas.iloc[0]["[Venda] nome"] == "Joao SP",
              str(confirmadas.iloc[0]["[Venda] nome"]))


if __name__ == "__main__":
    test_phone_keys()
    test_clean_phone_formats()
    test_phone_detection()
    test_header_detection()
    test_stray_phone()
    test_side_by_side_blocks()
    test_column_alignment()
    test_procv_ddd_precision()
    test_duplicates_ddd()
    test_disparo_window()

    print("\n" + "=" * 60)
    if _failures:
        print(f"❌ {len(_failures)} teste(s) falharam: {_failures}")
        sys.exit(1)
    print("✅ Todos os testes passaram.")
