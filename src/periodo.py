"""Funções de mês e caminhos de arquivos, compartilhadas por todas as etapas.

O pipeline pode rodar em dois modos:
- anual (mes=None): o comportamento original, com o ano de 2026 inteiro;
- mensal (mes = pd.Period, ex.: 2026-10): só as vendas daquele mês.

Os caminhos de entrada e saída de cada etapa dependem do modo, por isso
ficam todos aqui, num lugar só.
"""

import argparse
import re
from pathlib import Path

import pandas as pd

PASTA_PROJETO = Path(__file__).resolve().parent.parent

# Meses em português definidos aqui, para não depender do idioma do computador.
MESES_ABREVIADOS = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez",
}
MESES_POR_EXTENSO = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}


# ---------------------------------------------------------------------------
# Meses
# ---------------------------------------------------------------------------

def ler_mes(texto):
    """Converte um texto "AAAA-MM" em um período mensal do pandas.

    Usada como `type=` do argparse: se o texto for inválido, o argparse
    mostra a mensagem de erro e encerra o programa.
    """
    if not re.fullmatch(r"\d{4}-\d{2}", texto) or not 1 <= int(texto[5:]) <= 12:
        raise argparse.ArgumentTypeError(
            f"mês inválido: {texto!r}. Use o formato AAAA-MM, ex.: 2026-10"
        )
    return pd.Period(texto, freq="M")


def mes_padrao(hoje=None):
    """Mês anterior ao atual (o último mês completo), ex.: em 05/11/2026 -> 2026-10."""
    hoje = pd.Timestamp.today() if hoje is None else pd.Timestamp(hoje)
    return hoje.to_period("M") - 1


def mes_anterior(mes):
    """Mês anterior a `mes`, ex.: 2026-01 -> 2025-12."""
    return mes - 1


def nome_do_mes(periodo):
    """Nome curto do mês, ex.: 2026-03 -> "mar/2026"."""
    return f"{MESES_ABREVIADOS[periodo.month]}/{periodo.year}"


def mes_por_extenso(periodo):
    """Nome completo do mês, ex.: 2026-10 -> "outubro de 2026"."""
    return f"{MESES_POR_EXTENSO[periodo.month]} de {periodo.year}"


def adicionar_argumento_mes(parser):
    """Acrescenta a opção --mes (opcional) a um leitor de argumentos."""
    parser.add_argument(
        "--mes",
        type=ler_mes,
        default=None,
        help="mês no formato AAAA-MM (ex.: 2026-10); sem esta opção, usa o ano de 2026 inteiro",
    )


# ---------------------------------------------------------------------------
# Caminhos (mes=None = modo anual, com os caminhos originais)
# ---------------------------------------------------------------------------

def pasta_raw(mes=None):
    """Pasta das planilhas brutas: data/raw/ ou data/raw/AAAA-MM/."""
    pasta = PASTA_PROJETO / "data" / "raw"
    return pasta if mes is None else pasta / str(mes)


def arquivo_base(mes=None):
    """Base limpa: vendas_consolidadas.csv ou vendas_AAAA-MM.csv."""
    nome = "vendas_consolidadas.csv" if mes is None else f"vendas_{mes}.csv"
    return PASTA_PROJETO / "data" / "processed" / nome


def arquivo_relatorio(mes=None):
    """Relatório em Excel: relatorio_vendas_2026.xlsx ou relatorio_vendas_AAAA-MM.xlsx."""
    sufixo = "2026" if mes is None else str(mes)
    return PASTA_PROJETO / "output" / f"relatorio_vendas_{sufixo}.xlsx"


def complemento_comando(mes=None):
    """Texto " --mes AAAA-MM" para as mensagens que sugerem um comando ("" no modo anual)."""
    return "" if mes is None else f" --mes {mes}"
