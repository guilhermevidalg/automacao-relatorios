"""Testes da leitura e do mapeamento de colunas (src/leitura.py).

Rodar com:  .venv/bin/python -m pytest -v
"""

from pathlib import Path

import pandas as pd
import pytest

from gerar_dados import CONFIG_FILIAIS
from leitura import (
    COLUNAS_PADRAO,
    MAPA_COLUNAS,
    encontrar_linha_cabecalho,
    nome_da_filial,
    normalizar_nome_coluna,
    padronizar_colunas,
)


@pytest.mark.parametrize("nome, esperado", [
    ("Preço Unitário", "preco unitario"),
    ("preço unit.", "preco unit"),
    ("data_venda", "data venda"),
    ("DT VENDA", "dt venda"),
    ("Qtd.", "qtd"),
    ("Produto ", "produto"),
    (None, ""),       # célula vazia
    (42, ""),         # número no lugar do nome
])
def test_normalizar_nome_coluna(nome, esperado):
    assert normalizar_nome_coluna(nome) == esperado


# Lista (filial, nome da coluna, coluna padrão esperada) de todas as planilhas
# geradas. Liga o gerador ao leitor: se alguém mudar um nome de coluna no
# gerador sem atualizar SINONIMOS_COLUNAS, este teste falha.
COLUNAS_DO_GERADOR = [
    (filial, nome, padrao)
    for filial, config in CONFIG_FILIAIS.items()
    for nome, padrao in zip(config["colunas"], COLUNAS_PADRAO)
]


@pytest.mark.parametrize("filial, nome, padrao", COLUNAS_DO_GERADOR)
def test_colunas_de_todas_as_filiais_sao_reconhecidas(filial, nome, padrao):
    assert MAPA_COLUNAS[normalizar_nome_coluna(nome)] == padrao


def test_padronizar_colunas_renomeia_e_ordena():
    df = pd.DataFrame(columns=["VENDEDOR", "Qtd.", "DT VENDA", "Produto ", "VL UNIT", "CATEGORIA"])
    assert padronizar_colunas(df, "teste.xlsx").columns.tolist() == COLUNAS_PADRAO


def test_padronizar_colunas_descarta_desconhecidas():
    df = pd.DataFrame({
        "Data": ["15/03/2026"], "Produto": ["Notebook"], "Categoria": ["Informática"],
        "Qtd": [1], "Preço": [3500], "Vendedor": ["Ana"], "Observação": ["cliente VIP"],
    })
    resultado = padronizar_colunas(df, "teste.xlsx")
    assert "Observação" not in resultado.columns
    assert resultado.columns.tolist() == COLUNAS_PADRAO


def test_padronizar_colunas_cria_faltantes_vazias():
    df = pd.DataFrame({"Data": ["15/03/2026"], "Produto": ["Notebook"], "Qtd": [1], "Preço": [3500]})
    resultado = padronizar_colunas(df, "teste.xlsx")
    assert resultado.columns.tolist() == COLUNAS_PADRAO
    assert resultado["categoria"].isna().all()
    assert resultado["vendedor"].isna().all()


def test_encontrar_cabecalho_na_primeira_linha():
    bruto = pd.DataFrame([["Data", "Produto", "Qtd", "Preço"], ["15/03/2026", "Notebook", 1, 3500]])
    assert encontrar_linha_cabecalho(bruto) == 0


def test_encontrar_cabecalho_depois_de_titulo():
    bruto = pd.DataFrame([
        ["Relatório de Vendas – Filial Curitiba", None, None, None],
        [None, None, None, None],
        ["Data", "Produto", "Qtd.", "Preço"],
        ["15/03/2026", "Notebook", 1, "R$ 3.500,00"],
    ])
    assert encontrar_linha_cabecalho(bruto) == 2


def test_encontrar_cabecalho_inexistente_gera_erro():
    bruto = pd.DataFrame([["a", "b", "c"], [1, 2, 3]])
    with pytest.raises(ValueError):
        encontrar_linha_cabecalho(bruto)


@pytest.mark.parametrize("arquivo, esperado", [
    ("vendas_sao_paulo.xlsx", "São Paulo"),
    ("vendas_bh.xlsx", "BH"),
    ("vendas_porto_alegre.xlsx", "Porto Alegre"),  # fora da lista: nome deduzido
])
def test_nome_da_filial(arquivo, esperado):
    assert nome_da_filial(Path(arquivo)) == esperado
