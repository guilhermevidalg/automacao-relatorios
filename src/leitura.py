"""Leitura das planilhas de vendas de todas as filiais.

Responsabilidade deste módulo: trazer os dados "para dentro" do Python.
- encontra em qual linha começa o cabeçalho de cada planilha;
- traduz os nomes de colunas variados para um padrão único;
- adiciona a coluna "filial" e junta todas as planilhas em um só DataFrame.

A correção dos valores (datas, preços, textos etc.) fica em limpeza.py.
"""

import logging
import re
import unicodedata
from pathlib import Path

import pandas as pd

from periodo import pasta_raw

# Cada módulo usa o seu próprio "logger"; a configuração (onde gravar,
# formato das mensagens) é feita uma única vez em limpeza.py.
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

COLUNAS_PADRAO = ["data", "produto", "categoria", "quantidade", "preco_unitario", "vendedor"]

# Para cada coluna padrão, os nomes que já apareceram nas planilhas.
# Os nomes aqui estão "normalizados" (ver normalizar_nome_coluna):
# minúsculas, sem acento e sem pontuação.
SINONIMOS_COLUNAS = {
    "data": ["data", "data venda", "dt venda", "dt"],
    "produto": ["produto", "item"],
    "categoria": ["categoria"],
    "quantidade": ["quantidade", "qtd", "qtde", "quant"],
    "preco_unitario": ["preco unitario", "preco unit", "preco", "vl unit", "valor unitario"],
    "vendedor": ["vendedor", "vendedora"],
}

# Dicionário invertido: sinônimo -> coluna padrão (facilita a consulta).
MAPA_COLUNAS = {
    sinonimo: padrao
    for padrao, sinonimos in SINONIMOS_COLUNAS.items()
    for sinonimo in sinonimos
}

# Nome do arquivo (sem extensão) -> nome da filial.
NOMES_FILIAIS = {
    "vendas_rio": "Rio",
    "vendas_sao_paulo": "São Paulo",
    "vendas_bh": "BH",
    "vendas_curitiba": "Curitiba",
    "vendas_salvador": "Salvador",
}

LINHAS_PROCURA_CABECALHO = 10  # quantas linhas do topo examinar
MINIMO_COLUNAS_RECONHECIDAS = 3  # para considerar que a linha é o cabeçalho


# ---------------------------------------------------------------------------
# Funções
# ---------------------------------------------------------------------------

def remover_acentos(texto):
    """Remove acentos de um texto, ex.: "Preço" -> "Preco"."""
    # NFKD separa a letra do acento ("ç" vira "c" + cedilha); depois
    # descartamos os sinais de acentuação (categoria Unicode "Mn").
    decomposto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in decomposto if unicodedata.category(c) != "Mn")


def normalizar_nome_coluna(nome):
    """Simplifica um nome de coluna para facilitar a comparação.

    Exemplo: "Preço Unit." -> "preco unit"; "data_venda" -> "data venda".
    Valores que não são texto (ex.: células vazias) viram "".
    """
    if not isinstance(nome, str):
        return ""
    nome = remover_acentos(nome).lower()
    nome = re.sub(r"[^a-z0-9]+", " ", nome)  # pontuação e "_" viram espaço
    return nome.strip()


def encontrar_linha_cabecalho(df_bruto):
    """Devolve o número da linha (a partir de 0) onde está o cabeçalho.

    Procura nas primeiras linhas a primeira que tenha pelo menos
    MINIMO_COLUNAS_RECONHECIDAS nomes de colunas conhecidos. Assim
    funciona mesmo quando a planilha tem um título antes da tabela.
    """
    for indice in range(min(LINHAS_PROCURA_CABECALHO, len(df_bruto))):
        valores = df_bruto.iloc[indice]
        reconhecidas = sum(normalizar_nome_coluna(v) in MAPA_COLUNAS for v in valores)
        if reconhecidas >= MINIMO_COLUNAS_RECONHECIDAS:
            return indice
    raise ValueError("Cabeçalho não encontrado nas primeiras linhas da planilha.")


def padronizar_colunas(df, nome_arquivo):
    """Renomeia as colunas para o padrão e descarta as desconhecidas.

    Colunas padrão que não existirem na planilha são criadas vazias,
    para que todas as filiais tenham as mesmas colunas.
    """
    novos_nomes = {}
    for coluna in df.columns:
        padrao = MAPA_COLUNAS.get(normalizar_nome_coluna(coluna))
        if padrao:
            novos_nomes[coluna] = padrao
        elif df[coluna].notna().any():
            log.warning("%s: coluna desconhecida %r ignorada", nome_arquivo, coluna)

    df = df[list(novos_nomes)].rename(columns=novos_nomes)

    for coluna in COLUNAS_PADRAO:
        if coluna not in df.columns:
            log.warning("%s: coluna %r não encontrada; criada vazia", nome_arquivo, coluna)
            df[coluna] = None
    return df[COLUNAS_PADRAO]


def nome_da_filial(caminho):
    """Descobre o nome da filial a partir do nome do arquivo."""
    nome_arquivo = caminho.stem
    if nome_arquivo in NOMES_FILIAIS:
        return NOMES_FILIAIS[nome_arquivo]
    # Arquivo novo: deduz o nome, ex.: "vendas_porto_alegre" -> "Porto Alegre".
    filial = nome_arquivo.removeprefix("vendas_").replace("_", " ").title()
    log.warning("%s: filial fora da lista; usando o nome %r", caminho.name, filial)
    return filial


def ler_planilha(caminho):
    """Lê uma planilha e devolve um DataFrame com as colunas padrão + "filial"."""
    # header=None: lemos tudo como dados, inclusive título e cabeçalho,
    # para depois descobrir onde o cabeçalho realmente está.
    # dtype=object: não deixa o pandas converter nada por conta própria.
    df_bruto = pd.read_excel(caminho, header=None, dtype=object)

    linha_cabecalho = encontrar_linha_cabecalho(df_bruto)
    df = df_bruto.iloc[linha_cabecalho + 1:].copy()
    df.columns = df_bruto.iloc[linha_cabecalho]

    df = padronizar_colunas(df, caminho.name)
    df["filial"] = nome_da_filial(caminho)

    log.info(
        "%s: %d linhas lidas (cabeçalho na linha %d do Excel)",
        caminho.name, len(df), linha_cabecalho + 1,
    )
    return df


def ler_todas_planilhas(pasta=None, comando_gerar=".venv/bin/python src/gerar_dados.py"):
    """Lê todas as planilhas .xlsx da pasta e junta em um único DataFrame.

    Args:
        pasta: pasta das planilhas; por padrão, data/raw/ (modo anual).
        comando_gerar: comando sugerido na mensagem de erro se a pasta estiver vazia.
    """
    pasta = pasta_raw() if pasta is None else Path(pasta)
    # Arquivos que começam com "~$" são temporários do Excel (planilha aberta).
    arquivos = sorted(p for p in pasta.glob("*.xlsx") if not p.name.startswith("~$"))
    if not arquivos:
        raise FileNotFoundError(
            f"Nenhuma planilha encontrada em {pasta}. Rode primeiro: {comando_gerar}"
        )

    tabelas = [ler_planilha(caminho) for caminho in arquivos]
    # ignore_index=True: cria uma numeração única para as linhas de todas as filiais.
    return pd.concat(tabelas, ignore_index=True)
