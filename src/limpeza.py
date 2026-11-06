"""Limpeza e consolidação das vendas de todas as filiais.

Etapas, nesta ordem:
1. remove linhas totalmente vazias;
2. converte datas (vários formatos, inclusive por extenso), preços
   ("R$ 1.234,56") e quantidades para os tipos corretos;
3. padroniza textos (espaços e maiúsculas/minúsculas);
4. trata valores faltantes;
5. remove duplicatas e "quase duplicatas";
6. (só no modo mensal) remove vendas com data fora do mês;
7. calcula o faturamento e salva a base limpa em CSV.

Tudo o que é removido ou corrigido é registrado em log (terminal e
logs/limpeza.log).

Uso:
    .venv/bin/python src/limpeza.py                # data/raw/ -> vendas_consolidadas.csv
    .venv/bin/python src/limpeza.py --mes 2026-10  # data/raw/2026-10/ -> vendas_2026-10.csv
"""

import argparse
import logging
import re
from datetime import datetime

import pandas as pd

from leitura import ler_todas_planilhas, remover_acentos
from periodo import (
    PASTA_PROJETO,
    adicionar_argumento_mes,
    arquivo_base,
    complemento_comando,
    nome_do_mes,
    pasta_raw,
)

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

ARQUIVO_LOG = PASTA_PROJETO / "logs" / "limpeza.log"
FORMATO_LOG = "%(asctime)s | %(levelname)-7s | %(message)s"

COLUNAS_TEXTO = ["produto", "categoria", "vendedor"]

# Sem estas informações não dá para calcular o faturamento: a linha é removida.
COLUNAS_OBRIGATORIAS = ["data", "produto", "quantidade", "preco_unitario"]

ORDEM_COLUNAS_SAIDA = [
    "data", "filial", "produto", "categoria", "vendedor",
    "quantidade", "preco_unitario", "faturamento",
]

# Meses sem acento, pois o texto da data passa por remover_acentos antes.
MESES = {
    "janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4, "maio": 5, "junho": 6,
    "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12,
}

# O Excel guarda datas como "número de dias desde 30/12/1899".
INICIO_DATAS_EXCEL = pd.Timestamp("1899-12-30")


# ---------------------------------------------------------------------------
# Log
# ---------------------------------------------------------------------------

def configurar_log():
    """Configura o log para aparecer no terminal (o arquivo é tratado em executar)."""
    logging.basicConfig(level=logging.INFO, format=FORMATO_LOG, datefmt="%H:%M:%S")


def abrir_log_arquivo():
    """Passa a gravar o log também em logs/limpeza.log (sobrescrito a cada execução).

    O "handler" (destino do log) é ligado ao logger principal e devolvido,
    para ser desligado no fim da limpeza. Assim o arquivo funciona tanto
    rodando este script sozinho quanto dentro do pipeline (main.py).
    """
    ARQUIVO_LOG.parent.mkdir(parents=True, exist_ok=True)
    arquivo = logging.FileHandler(ARQUIVO_LOG, mode="w", encoding="utf-8")
    arquivo.setFormatter(logging.Formatter(FORMATO_LOG, datefmt="%H:%M:%S"))
    logging.getLogger().addHandler(arquivo)
    return arquivo


def fechar_log_arquivo(arquivo):
    """Desliga e fecha o arquivo de log aberto por abrir_log_arquivo."""
    logging.getLogger().removeHandler(arquivo)
    arquivo.close()


def contagem_por_filial(df):
    """Monta um texto como "Rio: 3, Salvador: 2" com o número de linhas por filial."""
    contagem = df["filial"].value_counts().sort_index()
    return ", ".join(f"{filial}: {qtd}" for filial, qtd in contagem.items())


# ---------------------------------------------------------------------------
# Linhas vazias
# ---------------------------------------------------------------------------

def remover_linhas_vazias(df):
    """Remove linhas em que todas as colunas (exceto "filial") estão vazias."""
    colunas_dados = [c for c in df.columns if c != "filial"]
    vazias = df[colunas_dados].isna().all(axis=1)
    if vazias.any():
        log.info(
            "Linhas vazias: %d removidas (%s)",
            vazias.sum(), contagem_por_filial(df[vazias]),
        )
    return df[~vazias]


# ---------------------------------------------------------------------------
# Datas
# ---------------------------------------------------------------------------

def converter_data_texto(texto):
    """Converte uma data escrita como texto. Devolve NaT se não reconhecer.

    Formatos aceitos: "15/03/2026", "2026-03-15" e "15 de março de 2026".
    Usamos formatos explícitos para nunca confundir dia com mês.
    """
    texto = remover_acentos(texto).strip().lower()

    if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", texto):
        return pd.to_datetime(texto, format="%d/%m/%Y", errors="coerce")

    # Aceita também "2026-03-15 00:00:00" (data com hora zerada).
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}( 00:00:00)?", texto):
        return pd.to_datetime(texto[:10], format="%Y-%m-%d", errors="coerce")

    por_extenso = re.fullmatch(r"(\d{1,2}) de ([a-z]+) de (\d{4})", texto)
    if por_extenso:
        dia, nome_mes, ano = por_extenso.groups()
        if nome_mes in MESES:
            try:
                return pd.Timestamp(year=int(ano), month=MESES[nome_mes], day=int(dia))
            except ValueError:  # ex.: "31 de fevereiro"
                return pd.NaT

    return pd.NaT


def converter_data(valor):
    """Converte um valor de data de qualquer formato conhecido para Timestamp."""
    if pd.isna(valor):
        return pd.NaT
    if isinstance(valor, datetime):  # já é data (célula de data do Excel)
        return pd.Timestamp(valor).normalize()
    if isinstance(valor, (int, float)):  # número serial do Excel
        return INICIO_DATAS_EXCEL + pd.Timedelta(days=int(valor))
    if isinstance(valor, str):
        return converter_data_texto(valor)
    return pd.NaT


def converter_datas(df):
    """Converte a coluna "data" e registra quantas datas eram texto ou inválidas."""
    df = df.copy()
    originais = df["data"]
    df["data"] = pd.to_datetime(originais.map(converter_data))

    eram_texto = originais.map(lambda v: isinstance(v, str)).sum()
    invalidas = originais.notna() & df["data"].isna()
    log.info("Datas: %d convertidas de texto para data", eram_texto)
    if invalidas.any():
        log.warning(
            "Datas: %d em formato não reconhecido viraram vazias: %s",
            invalidas.sum(), originais[invalidas].unique()[:5].tolist(),
        )
    return df


# ---------------------------------------------------------------------------
# Preços e quantidades
# ---------------------------------------------------------------------------

def converter_preco(valor):
    """Converte um preço para número, ex.: "R$ 1.234,56" -> 1234.56."""
    if pd.isna(valor):
        return float("nan")
    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).replace("R$", "").replace("\xa0", "").replace(" ", "")
    if "," in texto:
        # Formato brasileiro: ponto separa milhar e vírgula separa centavos.
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return float("nan")


def converter_precos(df):
    """Converte a coluna "preco_unitario" e registra quantos preços eram texto."""
    df = df.copy()
    originais = df["preco_unitario"]
    df["preco_unitario"] = originais.map(converter_preco)

    eram_texto = originais.map(lambda v: isinstance(v, str))
    if eram_texto.any():
        log.info(
            "Preços: %d convertidos de texto (\"R$ 1.234,56\") para número (%s)",
            eram_texto.sum(), contagem_por_filial(df[eram_texto]),
        )
    invalidos = originais.notna() & df["preco_unitario"].isna()
    if invalidos.any():
        log.warning("Preços: %d em formato não reconhecido viraram vazios", invalidos.sum())
    return df


def converter_quantidades(df):
    """Converte "quantidade" para número inteiro (Int64 aceita valores vazios)."""
    df = df.copy()
    originais = df["quantidade"]
    numeros = pd.to_numeric(originais, errors="coerce")
    invalidas = originais.notna() & numeros.isna()
    if invalidas.any():
        log.warning("Quantidades: %d não numéricas viraram vazias", invalidas.sum())
    df["quantidade"] = numeros.round().astype("Int64")
    return df


# ---------------------------------------------------------------------------
# Textos
# ---------------------------------------------------------------------------

def limpar_espacos(valor):
    """Remove espaços nas pontas e troca espaços repetidos por um só."""
    if not isinstance(valor, str):
        return valor
    return " ".join(valor.split())


def padronizar_maiusculas(serie):
    """Unifica variações de maiúsculas/minúsculas usando a grafia mais comum.

    Ex.: "NOTEBOOK", "notebook" e "Notebook" são agrupados (mesma versão
    em minúsculas) e todos viram a grafia que mais aparece no grupo.
    Assim preservamos nomes como "Pen drive 64GB", que .title() estragaria.
    """
    chave = serie.str.casefold()  # casefold: minúsculas, inclusive para acentos
    grafia_oficial = serie.groupby(chave).agg(lambda grupo: grupo.value_counts().idxmax())
    return chave.map(grafia_oficial)


def padronizar_textos(df):
    """Padroniza espaços e maiúsculas/minúsculas das colunas de texto."""
    df = df.copy()
    for coluna in COLUNAS_TEXTO:
        originais = df[coluna]
        df[coluna] = padronizar_maiusculas(originais.map(limpar_espacos))

        alteradas = originais.notna() & (originais != df[coluna])
        if alteradas.any():
            log.info(
                "Textos: %d valores de %r corrigidos (espaços/maiúsculas) (%s)",
                alteradas.sum(), coluna, contagem_por_filial(df[alteradas]),
            )
    return df


# ---------------------------------------------------------------------------
# Valores faltantes
# ---------------------------------------------------------------------------

def remover_linhas_incompletas(df):
    """Remove linhas sem alguma informação obrigatória para o faturamento."""
    faltando = df[COLUNAS_OBRIGATORIAS].isna()
    incompletas = faltando.any(axis=1)
    if incompletas.any():
        detalhe = ", ".join(
            f"{coluna}: {qtd}" for coluna, qtd in faltando.sum().items() if qtd > 0
        )
        log.info(
            "Faltantes: %d linhas removidas por falta de dado obrigatório "
            "(vazios por coluna -> %s) (%s)",
            incompletas.sum(), detalhe, contagem_por_filial(df[incompletas]),
        )
    return df[~incompletas]


def preencher_categoria(df):
    """Preenche categorias vazias com a categoria mais comum do mesmo produto."""
    df = df.copy()
    vazias = df["categoria"].isna()
    if not vazias.any():
        return df

    categoria_do_produto = (
        df.dropna(subset=["categoria"])
        .groupby("produto")["categoria"]
        .agg(lambda grupo: grupo.value_counts().idxmax())
    )
    df.loc[vazias, "categoria"] = df.loc[vazias, "produto"].map(categoria_do_produto)
    sem_referencia = df["categoria"].isna().sum()
    df["categoria"] = df["categoria"].fillna("Não informada")

    log.info(
        "Faltantes: %d categorias preenchidas a partir do produto (%d sem referência -> \"Não informada\")",
        vazias.sum(), sem_referencia,
    )
    return df


def preencher_vendedor(df):
    """Preenche vendedores vazios com "Não informado"."""
    df = df.copy()
    vazios = df["vendedor"].isna()
    if vazios.any():
        df["vendedor"] = df["vendedor"].fillna("Não informado")
        log.info(
            "Faltantes: %d vendedores vazios preenchidos com \"Não informado\" (%s)",
            vazios.sum(), contagem_por_filial(df[vazios]),
        )
    return df


def tratar_faltantes(df):
    """Aplica as regras de valores faltantes (remover ou preencher)."""
    df = remover_linhas_incompletas(df)
    df = preencher_categoria(df)
    df = preencher_vendedor(df)
    return df


# ---------------------------------------------------------------------------
# Duplicatas
# ---------------------------------------------------------------------------

def remover_duplicatas(df, df_original):
    """Remove duplicatas, separando no log as exatas das "quase duplicatas".

    Como esta etapa roda depois da padronização, linhas que só diferiam por
    espaços, maiúsculas ou formato de data/preço já ficaram idênticas.

    Args:
        df: dados já padronizados.
        df_original: dados como vieram da planilha (mesmos índices de linha),
            usados para saber quais duplicatas já eram idênticas na origem.
    """
    duplicadas = df.duplicated(keep="first")
    if not duplicadas.any():
        return df

    # Duplicata exata = já era idêntica na planilha, antes de qualquer correção.
    exatas = df_original.loc[df.index].duplicated(keep="first").sum()
    quase = duplicadas.sum() - exatas
    log.info(
        "Duplicatas: %d removidas (%d exatas, %d quase duplicatas) (%s)",
        duplicadas.sum(), exatas, quase, contagem_por_filial(df[duplicadas]),
    )
    return df[~duplicadas]


# ---------------------------------------------------------------------------
# Mês (só no modo mensal)
# ---------------------------------------------------------------------------

def remover_fora_do_mes(df, mes):
    """Remove vendas cuja data não pertence ao mês do relatório.

    Args:
        df: dados com a coluna "data" já convertida.
        mes: período mensal do pandas (ex.: 2026-10).
    """
    fora = df["data"].dt.to_period("M") != mes
    if fora.any():
        log.warning(
            "Mês: %d vendas fora de %s removidas (%s)",
            fora.sum(), nome_do_mes(mes), contagem_por_filial(df[fora]),
        )
    return df[~fora]


# ---------------------------------------------------------------------------
# Faturamento e gravação
# ---------------------------------------------------------------------------

def calcular_faturamento(df):
    """Adiciona a coluna "faturamento" = quantidade x preço unitário."""
    df = df.copy()
    df["faturamento"] = (df["quantidade"] * df["preco_unitario"]).astype(float).round(2)
    return df


def limpar_dados(df, mes=None):
    """Executa todas as etapas de limpeza, na ordem certa.

    Args:
        df: dados brutos, como vieram de ler_todas_planilhas.
        mes: período mensal; se informado, descarta vendas de outros meses.
    """
    linhas_lidas = len(df)

    df = remover_linhas_vazias(df)
    df_original = df.copy()  # "foto" dos dados brutos, usada para classificar duplicatas

    df = converter_datas(df)
    df = converter_precos(df)
    df = converter_quantidades(df)
    df = padronizar_textos(df)
    df = tratar_faltantes(df)
    df = remover_duplicatas(df, df_original)
    if mes is not None:
        df = remover_fora_do_mes(df, mes)
    df = calcular_faturamento(df)

    df = df[ORDEM_COLUNAS_SAIDA].sort_values(["data", "filial"]).reset_index(drop=True)
    log.info(
        "Resumo: %d linhas lidas -> %d linhas na base limpa (%d removidas)",
        linhas_lidas, len(df), linhas_lidas - len(df),
    )
    log.info("Linhas finais por filial: %s", contagem_por_filial(df))
    return df


def salvar_base(df, caminho):
    """Salva a base limpa em CSV (UTF-8 com BOM, para o Excel ler os acentos)."""
    caminho.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(caminho, index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
    log.info("Base limpa salva em %s", caminho.relative_to(PASTA_PROJETO))


def executar(mes=None):
    """Lê as planilhas brutas, limpa os dados e salva a base consolidada.

    Args:
        mes: período mensal (ex.: 2026-10) ou None para o ano inteiro.
    """
    arquivo_log = abrir_log_arquivo()
    try:
        comando_gerar = ".venv/bin/python src/gerar_dados.py" + complemento_comando(mes)
        df = ler_todas_planilhas(pasta_raw(mes), comando_gerar)
        df = limpar_dados(df, mes)
        salvar_base(df, arquivo_base(mes))
    finally:
        # "finally" roda mesmo se der erro: o arquivo de log sempre é fechado.
        fechar_log_arquivo(arquivo_log)


def main():
    """Lê a opção --mes e executa a limpeza."""
    parser = argparse.ArgumentParser(description="Lê, limpa e consolida as planilhas de vendas.")
    adicionar_argumento_mes(parser)
    argumentos = parser.parse_args()
    configurar_log()
    executar(argumentos.mes)


if __name__ == "__main__":
    main()
