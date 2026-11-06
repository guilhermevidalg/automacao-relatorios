"""Testes das funções de mês (src/periodo.py) e da variação mensal (src/relatorio.py).

Rodar com:  .venv/bin/python -m pytest -v
"""

import argparse

import pandas as pd
import pytest

from periodo import ler_mes, mes_anterior, mes_padrao, nome_do_mes
from relatorio import calcular_variacao, rotulo_semana


def test_ler_mes_valido():
    assert ler_mes("2026-10") == pd.Period("2026-10", freq="M")


@pytest.mark.parametrize("texto", ["2026-13", "2026-00", "2026-1", "10/2026", "outubro", ""])
def test_ler_mes_invalido(texto):
    with pytest.raises(argparse.ArgumentTypeError):
        ler_mes(texto)


def test_mes_padrao_e_o_mes_anterior():
    assert mes_padrao("2026-11-05") == pd.Period("2026-10", freq="M")
    assert mes_padrao("2026-01-31") == pd.Period("2025-12", freq="M")


def test_mes_anterior_vira_o_ano():
    assert mes_anterior(pd.Period("2026-01", freq="M")) == pd.Period("2025-12", freq="M")


def test_nome_do_mes():
    assert nome_do_mes(pd.Period("2026-09", freq="M")) == "set/2026"


@pytest.mark.parametrize("atual, anterior, esperado", [
    (108, 100, 0.08),
    (90, 100, -0.10),
    (100, 100, 0.0),
])
def test_calcular_variacao(atual, anterior, esperado):
    assert calcular_variacao(atual, anterior) == pytest.approx(esperado)


def test_calcular_variacao_sem_base_anterior():
    assert calcular_variacao(100, 0) is None


def test_rotulo_semana():
    outubro = pd.Period("2026-10", freq="M")
    assert rotulo_semana(0, outubro) == "01–07/out"
    assert rotulo_semana(4, outubro) == "29–31/out"
