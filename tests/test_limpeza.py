"""Testes das funções de limpeza (src/limpeza.py).

Rodar com:  .venv/bin/python -m pytest -v
"""

import math
from datetime import datetime

import pandas as pd
import pytest

from limpeza import (
    converter_data,
    converter_preco,
    limpar_dados,
    limpar_espacos,
    padronizar_maiusculas,
    preencher_categoria,
    preencher_vendedor,
    remover_duplicatas,
    remover_fora_do_mes,
    remover_linhas_incompletas,
)

DIA_15_MARCO = pd.Timestamp("2026-03-15")


# ---------------------------------------------------------------------------
# Datas
# ---------------------------------------------------------------------------

# parametrize: o mesmo teste roda uma vez para cada par (entrada, esperado).
@pytest.mark.parametrize("valor", [
    datetime(2026, 3, 15),             # célula de data do Excel
    datetime(2026, 3, 15, 14, 30),     # data com hora: a hora é descartada
    pd.Timestamp("2026-03-15"),
    46096,                             # número serial do Excel (inteiro)
    46096.0,                           # número serial do Excel (decimal)
    "15/03/2026",                      # dd/mm/aaaa
    "5/3/2026",                        # dia e mês com um dígito
    "2026-03-15",                      # aaaa-mm-dd
    "2026-03-15 00:00:00",             # aaaa-mm-dd com hora zerada
    "15 de março de 2026",             # por extenso
    "15 de marco de 2026",             # por extenso, sem acento
    "  15 DE MARÇO DE 2026 ",          # maiúsculas e espaços nas pontas
])
def test_converter_data_formatos_aceitos(valor):
    esperado = pd.Timestamp("2026-03-05") if valor == "5/3/2026" else DIA_15_MARCO
    assert converter_data(valor) == esperado


@pytest.mark.parametrize("valor", [
    "31 de fevereiro de 2026",  # dia que não existe
    "15 de marçal de 2026",     # mês desconhecido
    "03/15/2026",               # formato americano: mês 15 não existe
    "15-03-2026",               # formato não suportado
    "ontem",
    "",
    None,
    float("nan"),
])
def test_converter_data_invalida_vira_nat(valor):
    assert pd.isna(converter_data(valor))


def test_converter_data_nunca_troca_dia_e_mes():
    # 04/05 tem que ser 4 de maio (padrão brasileiro), nunca 5 de abril.
    assert converter_data("04/05/2026") == pd.Timestamp("2026-05-04")


# ---------------------------------------------------------------------------
# Preços
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("valor, esperado", [
    ("R$ 1.234,56", 1234.56),
    ("R$ 49,90", 49.90),
    ("R$ 1.234.567,89", 1234567.89),   # mais de um ponto de milhar
    ("R$\xa01.234,56", 1234.56),       # espaço não separável (comum ao copiar do Excel)
    ("1.234,56", 1234.56),             # sem o "R$"
    ("1234.56", 1234.56),              # já no formato com ponto decimal
    (1234.56, 1234.56),                # número de verdade
    (50, 50.0),
])
def test_converter_preco(valor, esperado):
    assert converter_preco(valor) == pytest.approx(esperado)


@pytest.mark.parametrize("valor", ["R$ abc", "grátis", None, float("nan")])
def test_converter_preco_invalido_vira_nan(valor):
    assert math.isnan(converter_preco(valor))


# ---------------------------------------------------------------------------
# Textos
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("valor, esperado", [
    ("  Notebook  ", "Notebook"),
    ("Mouse   sem    fio", "Mouse sem fio"),
    ("Headset", "Headset"),
    (None, None),  # valores que não são texto passam sem mudança
])
def test_limpar_espacos(valor, esperado):
    assert limpar_espacos(valor) == esperado


def test_padronizar_maiusculas_usa_grafia_mais_frequente():
    serie = pd.Series(["Notebook", "Notebook", "NOTEBOOK", "notebook"])
    assert padronizar_maiusculas(serie).tolist() == ["Notebook"] * 4


def test_padronizar_maiusculas_preserva_nomes_especiais():
    # .title() transformaria em "Pen Drive 64Gb"; a grafia mais comum é mantida.
    serie = pd.Series(["Pen drive 64GB", "PEN DRIVE 64GB", "Pen drive 64GB"])
    assert padronizar_maiusculas(serie).tolist() == ["Pen drive 64GB"] * 3


def test_padronizar_maiusculas_mantem_vazios():
    serie = pd.Series(["Rio", None, "RIO", "Rio"])
    resultado = padronizar_maiusculas(serie)
    assert resultado[0] == resultado[2] == resultado[3] == "Rio"
    assert pd.isna(resultado[1])


# ---------------------------------------------------------------------------
# Valores faltantes
# ---------------------------------------------------------------------------

def venda(**campos):
    """Cria uma venda completa (dicionário); os campos informados substituem os padrões."""
    padrao = {
        "data": DIA_15_MARCO,
        "produto": "Notebook",
        "categoria": "Informática",
        "quantidade": 1,
        "preco_unitario": 3500.0,
        "vendedor": "Ana Souza",
        "filial": "Rio",
    }
    return {**padrao, **campos}


@pytest.mark.parametrize("coluna", ["data", "produto", "quantidade", "preco_unitario"])
def test_remover_linhas_incompletas_sem_campo_obrigatorio(coluna):
    df = pd.DataFrame([venda(), venda(**{coluna: None})])
    resultado = remover_linhas_incompletas(df)
    assert len(resultado) == 1
    assert resultado.index.tolist() == [0]


def test_remover_linhas_incompletas_mantem_sem_categoria_ou_vendedor():
    df = pd.DataFrame([venda(categoria=None), venda(vendedor=None)])
    assert len(remover_linhas_incompletas(df)) == 2


def test_preencher_categoria_com_a_mais_comum_do_produto():
    df = pd.DataFrame([
        venda(produto="Headset", categoria="Periféricos"),
        venda(produto="Headset", categoria="Periféricos"),
        venda(produto="Headset", categoria="Acessórios"),
        venda(produto="Headset", categoria=None),
    ])
    assert preencher_categoria(df)["categoria"].iloc[-1] == "Periféricos"


def test_preencher_categoria_sem_referencia():
    df = pd.DataFrame([venda(produto="Produto novo", categoria=None)])
    assert preencher_categoria(df)["categoria"].iloc[0] == "Não informada"


def test_preencher_vendedor():
    df = pd.DataFrame([venda(vendedor=None), venda(vendedor="Ana Souza")])
    assert preencher_vendedor(df)["vendedor"].tolist() == ["Não informado", "Ana Souza"]


# ---------------------------------------------------------------------------
# Duplicatas
# ---------------------------------------------------------------------------

def test_remover_duplicatas_exatas():
    df = pd.DataFrame([venda(), venda(), venda(quantidade=2)])
    resultado = remover_duplicatas(df, df_original=df)
    assert len(resultado) == 2
    assert resultado["quantidade"].tolist() == [1, 2]


def test_limpar_dados_remove_quase_duplicatas():
    """Linhas que só diferem em espaços, maiúsculas e formato de data/preço viram uma só."""
    df = pd.DataFrame([
        venda(data="15/03/2026", produto="Notebook", preco_unitario="R$ 3.500,00"),
        venda(data="2026-03-15", produto="  NOTEBOOK ", preco_unitario=3500.0),
        venda(data="15 de março de 2026", produto="Notebook", preco_unitario="R$ 3.500,00"),
        venda(data="16/03/2026", produto="Notebook", preco_unitario="R$ 3.500,00"),  # outro dia: não é duplicata
    ])
    resultado = limpar_dados(df)
    assert len(resultado) == 2
    assert resultado["produto"].unique().tolist() == ["Notebook"]
    assert resultado["faturamento"].tolist() == [3500.0, 3500.0]


# ---------------------------------------------------------------------------
# Mês (modo mensal)
# ---------------------------------------------------------------------------

def test_remover_fora_do_mes():
    df = pd.DataFrame([
        venda(data=pd.Timestamp("2026-10-01")),
        venda(data=pd.Timestamp("2026-10-31")),
        venda(data=pd.Timestamp("2026-09-30")),
        venda(data=pd.Timestamp("2026-11-01")),
    ])
    resultado = remover_fora_do_mes(df, pd.Period("2026-10", freq="M"))
    assert resultado["data"].dt.day.tolist() == [1, 31]
