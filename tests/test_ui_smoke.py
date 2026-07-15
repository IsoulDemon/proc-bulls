"""
Smoke da UI — executa o app INTEIRO (passos 1-2) com uploads simulados.

O boot headless sem planilhas não exercita o passo 2 (configuração):
NameError de ordem de definição só aparece com arquivos carregados
(regressão real: sales_name_col_match usado antes de sales_name_col).

Rodar:  python3 tests/test_ui_smoke.py
"""
import io
import os
import sys
import warnings

warnings.filterwarnings("ignore")
os.environ.setdefault("STREAMLIT_GLOBAL_DISABLE_WATCHDOG_WARNING", "true")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd  # noqa: E402
import streamlit as st  # noqa: E402


class FakeUpload:
    def __init__(self, data: bytes, name: str):
        self._d, self.name = data, name

    def read(self):
        return self._d


def _mk_vendas() -> FakeUpload:
    buf = io.BytesIO()
    pd.DataFrame({
        "Cliente": ["Ana Lima", "Bia Costa", "Caio Melo"],
        "Telefone": ["(31) 3333-4444", "(31) 3333-5555", ""],
        "Celular 1": ["(31) 98888-7777", "(31) 97777-6666", "(31) 96666-5555"],
        "Data da Venda": ["10/06/2026", "12/06/2026", "15/06/2026"],
        "Valor": ["150,00", "200,00", "99,90"],
    }).to_excel(buf, index=False)
    return FakeUpload(buf.getvalue(), "vendedora_ana.xlsx")


def _mk_vendas2() -> FakeUpload:
    # segunda vendedora, com o telefone noutra coluna ("Celular") — exercita
    # a unificação de colunas-chave do combine_sales_sources
    buf = io.BytesIO()
    pd.DataFrame({
        "Cliente": ["Duda Reis", "Eva Rocha"],
        "Celular": ["(31) 95555-4444", "(31) 94444-3333"],
        "Data da Venda": ["11/06/2026", "13/06/2026"],
        "Valor": ["80,00", "120,00"],
    }).to_excel(buf, index=False)
    return FakeUpload(buf.getvalue(), "vendedora_duda.xlsx")


def _mk_kommo() -> FakeUpload:
    csv = ("ID,Nome completo,Celular,Tags,Criado em\n"
           "1,Ana Lima,31988887777,TRAFEGO,01.06.2026 10:00:00\n"
           "2,Carla Dias,31966665555,\"DISPARO PROMO 10/06/26\",02.06.2026 11:00:00\n"
           "3,Duda Reis,31955554444,INSTAGRAM,03.06.2026 12:00:00\n")
    return FakeUpload(csv.encode("utf-8"), "kommo.csv")


def _mk_vip() -> FakeUpload:
    # lista do Grupo VIP: um telefone que casa com uma venda (Ana Lima)
    csv = "Nome,Telefone\nAna L,(31) 98888-7777\nZeca,(31) 90000-1111\n"
    return FakeUpload(csv.encode("utf-8"), "grupo_vip.csv")


def _fake_uploader(label, *a, **kw):
    if kw.get("key") == "sales_upload":
        return [_mk_vendas(), _mk_vendas2()]  # 2 listas → combine_sales_sources
    if kw.get("key") == "vip_upload":
        return [_mk_vip()]  # exercita o canal Grupo VIP no app inteiro
    if kw.get("accept_multiple_files"):
        return [_mk_kommo()] if kw.get("key") == "kommo_upload" else []
    return _mk_vendas()


_orig = st.file_uploader
st.file_uploader = _fake_uploader
try:
    # Executa o script completo em bare mode com as planilhas "carregadas":
    # todo o passo 2 (detecção, IA desligada, checkboxes, disparo) roda aqui.
    import app  # noqa: F401
finally:
    st.file_uploader = _orig

print("✅ Smoke da UI passou — o app executa com planilhas carregadas (passos 1-2).")
