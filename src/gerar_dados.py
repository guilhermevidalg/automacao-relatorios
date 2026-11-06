"""Gera planilhas de vendas fictícias (e propositalmente "sujas") por filial.

Modo anual (sem --mes): gera 5 planilhas .xlsx em data/raw/ (uma por filial),
cada uma com cerca de 300 vendas de 2026.
Modo mensal (--mes AAAA-MM): gera as 5 planilhas só daquele mês, com cerca de
100 vendas por filial, em data/raw/AAAA-MM/. A semente depende do mês: cada mês
tem dados diferentes, mas sempre os mesmos se o script for rodado de novo.
Os problemas de qualidade (nomes de colunas, formatos de data, preços em texto,
duplicatas, valores faltando etc.) são diferentes em cada filial, para simular
dados reais vindos de equipes que não seguem um padrão.

Uso:
    .venv/bin/python src/gerar_dados.py               # ano de 2026 inteiro
    .venv/bin/python src/gerar_dados.py --mes 2026-10 # só outubro de 2026
"""

import argparse

import numpy as np
import pandas as pd

from periodo import adicionar_argumento_mes, nome_do_mes, pasta_raw

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

SEED = 42  # semente fixa: garante que os dados gerados sejam sempre os mesmos
ANO = 2026
VENDAS_POR_FILIAL = 300      # modo anual
VENDAS_POR_FILIAL_MES = 100  # modo mensal (o suficiente para a "sujeira" aparecer)

# Colunas "oficiais" usadas internamente antes de sujar os dados.
COLUNAS_PADRAO = ["data", "produto", "categoria", "quantidade", "preco_unitario", "vendedor"]

# (produto, categoria, preço base em reais)
CATALOGO = [
    ("Notebook", "Informática", 3500.00),
    ("Monitor 24 polegadas", "Informática", 899.90),
    ("Mouse sem fio", "Periféricos", 79.90),
    ("Teclado mecânico", "Periféricos", 349.00),
    ("Headset", "Periféricos", 229.90),
    ("Webcam Full HD", "Periféricos", 189.00),
    ("Cadeira gamer", "Móveis", 1249.00),
    ("Mesa para escritório", "Móveis", 699.00),
    ("Impressora multifuncional", "Informática", 1099.00),
    ("Pen drive 64GB", "Acessórios", 49.90),
    ("Cabo HDMI", "Acessórios", 29.90),
    ("Mochila para notebook", "Acessórios", 159.90),
]

VENDEDORES = {
    "Rio": ["Ana Souza", "Bruno Lima", "Carla Mendes"],
    "São Paulo": ["Diego Alves", "Elisa Rocha", "Fábio Nunes", "Gabriela Dias"],
    "BH": ["Helena Castro", "Igor Martins", "Júlia Ramos"],
    "Curitiba": ["Karina Lopes", "Lucas Pereira", "Marina Teixeira"],
    "Salvador": ["Nádia Santos", "Otávio Barbosa", "Paula Freitas", "Rafael Costa"],
}

MESES = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]

# Configuração da "sujeira" de cada filial. Cada filial tem problemas diferentes.
CONFIG_FILIAIS = {
    "Rio": {
        "arquivo": "vendas_rio.xlsx",
        "colunas": ["Data", "Produto", "Categoria", "Quantidade", "Preço Unitário", "Vendedor"],
        "formato_data": "dd/mm/aaaa",
        "preco_texto": False,
        "proporcao_produtos_sujos": 0.30,
        "proporcao_duplicatas": 0.04,
        "proporcao_faltantes": 0.0,
        "linhas_em_branco": 0,
        "linha_cabecalho": 1,
    },
    "São Paulo": {
        "arquivo": "vendas_sao_paulo.xlsx",
        "colunas": ["data_venda", "produto", "categoria", "qtd", "preco_unitario", "vendedor"],
        "formato_data": "aaaa-mm-dd",
        "preco_texto": True,
        "proporcao_produtos_sujos": 0.0,
        "proporcao_duplicatas": 0.0,
        "proporcao_faltantes": 0.05,
        "linhas_em_branco": 0,
        "linha_cabecalho": 1,
    },
    "BH": {
        "arquivo": "vendas_bh.xlsx",
        "colunas": ["DT VENDA", "PRODUTO", "CATEGORIA", "QTDE", "VL UNIT", "VENDEDOR"],
        "formato_data": "texto",
        "preco_texto": False,
        "proporcao_produtos_sujos": 0.40,
        "proporcao_duplicatas": 0.0,
        "proporcao_faltantes": 0.0,
        "linhas_em_branco": 8,
        "linha_cabecalho": 1,
    },
    "Curitiba": {
        "arquivo": "vendas_curitiba.xlsx",
        "colunas": ["Data", "Produto", "Categoria", "Qtd.", "Preço", "Vendedor"],
        "formato_data": "excel",
        "preco_texto": True,
        "proporcao_produtos_sujos": 0.0,
        "proporcao_duplicatas": 0.0,
        "proporcao_faltantes": 0.0,
        "linhas_em_branco": 0,
        "linha_cabecalho": 3,  # título na linha 1, linha 2 vazia, cabeçalho na 3
    },
    "Salvador": {
        "arquivo": "vendas_salvador.xlsx",
        "colunas": ["data", "Produto ", "CATEGORIA", "quantidade", "preço unit.", "vendedor"],
        "formato_data": "misto",
        "preco_texto": True,
        "proporcao_produtos_sujos": 0.20,
        "proporcao_duplicatas": 0.03,
        "proporcao_faltantes": 0.03,
        "linhas_em_branco": 5,
        "linha_cabecalho": 1,
    },
}


# ---------------------------------------------------------------------------
# Geração dos dados limpos
# ---------------------------------------------------------------------------

def gerar_vendas_limpas(filial, rng, quantidade, data_inicial, numero_dias):
    """Gera um DataFrame com vendas corretas (sem sujeira) de uma filial.

    Args:
        filial: nome da filial (chave de VENDEDORES).
        rng: gerador de números aleatórios do numpy.
        quantidade: número de vendas a gerar.
        data_inicial: primeiro dia do período (pd.Timestamp).
        numero_dias: quantos dias o período tem (365 para o ano, 31 para outubro...).

    Returns:
        DataFrame com as colunas de COLUNAS_PADRAO, ordenado por data.
    """
    datas = data_inicial + pd.to_timedelta(rng.integers(0, numero_dias, quantidade), unit="D")

    indices_produtos = rng.integers(0, len(CATALOGO), quantidade)
    produtos = [CATALOGO[i][0] for i in indices_produtos]
    categorias = [CATALOGO[i][1] for i in indices_produtos]

    # Preço varia até 10% para cima ou para baixo do preço base (promoções, reajustes).
    precos_base = np.array([CATALOGO[i][2] for i in indices_produtos])
    precos = np.round(precos_base * rng.uniform(0.9, 1.1, quantidade), 2)

    df = pd.DataFrame({
        "data": datas,
        "produto": produtos,
        "categoria": categorias,
        "quantidade": rng.integers(1, 11, quantidade),
        "preco_unitario": precos,
        "vendedor": rng.choice(VENDEDORES[filial], quantidade),
    })
    return df.sort_values("data").reset_index(drop=True)


# ---------------------------------------------------------------------------
# Funções de "sujeira"
# ---------------------------------------------------------------------------

def sujar_texto(texto, rng):
    """Aplica uma alteração aleatória de espaços ou maiúsculas/minúsculas a um texto."""
    alteracoes = [
        lambda t: t.upper(),
        lambda t: t.lower(),
        lambda t: t.title(),
        lambda t: f"  {t}",
        lambda t: f"{t}   ",
        lambda t: t.replace(" ", "  "),
    ]
    alteracao = alteracoes[rng.integers(0, len(alteracoes))]
    return alteracao(texto)


def sujar_produtos(df, rng, proporcao):
    """Suja os nomes de uma parte dos produtos (espaços extras e caixa inconsistente)."""
    df = df.copy()
    quantidade = int(len(df) * proporcao)
    linhas = rng.choice(len(df), quantidade, replace=False)
    for linha in linhas:
        df.loc[linha, "produto"] = sujar_texto(df.loc[linha, "produto"], rng)
    return df


def formatar_uma_data(data, formato):
    """Converte uma data (Timestamp) para o formato indicado.

    Formatos: "dd/mm/aaaa", "aaaa-mm-dd", "texto" (ex.: "15 de março de 2026")
    ou "excel" (mantém como data de verdade, que o Excel reconhece).
    """
    if formato == "dd/mm/aaaa":
        return data.strftime("%d/%m/%Y")
    if formato == "aaaa-mm-dd":
        return data.strftime("%Y-%m-%d")
    if formato == "texto":
        return f"{data.day} de {MESES[data.month - 1]} de {data.year}"
    if formato == "excel":
        return data
    raise ValueError(f"Formato de data desconhecido: {formato}")


def formatar_datas(df, formato, rng):
    """Converte a coluna de datas para um formato; "misto" sorteia um formato por linha."""
    df = df.copy()
    if formato == "misto":
        opcoes = ["dd/mm/aaaa", "aaaa-mm-dd", "texto", "excel"]
        df["data"] = [formatar_uma_data(d, opcoes[rng.integers(0, len(opcoes))]) for d in df["data"]]
    else:
        df["data"] = [formatar_uma_data(d, formato) for d in df["data"]]
    df["data"] = df["data"].astype(object)  # coluna pode misturar texto e datas
    return df


def formatar_preco_texto(df):
    """Converte o preço para texto no padrão brasileiro, ex.: "R$ 1.234,56"."""
    df = df.copy()

    def para_reais(valor):
        # f"{1234.56:,.2f}" gera "1,234.56"; trocamos vírgula e ponto de lugar.
        texto = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {texto}"

    df["preco_unitario"] = df["preco_unitario"].apply(para_reais)
    return df


def duplicar_linhas(df, rng, proporcao):
    """Duplica uma parte das linhas; a cópia aparece logo abaixo da original."""
    quantidade = int(len(df) * proporcao)
    if quantidade == 0:
        return df
    linhas = rng.choice(len(df), quantidade, replace=False)
    # As cópias mantêm o mesmo índice da original; a ordenação estável as deixa juntas.
    resultado = pd.concat([df, df.iloc[linhas]]).sort_index(kind="stable")
    return resultado.reset_index(drop=True)


def apagar_valores(df, rng, proporcao):
    """Apaga valores aleatórios (deixa células vazias) em todas as colunas."""
    df = df.copy()
    quantidade = int(len(df) * proporcao)
    if quantidade == 0:
        return df
    for coluna in df.columns:
        linhas = rng.choice(len(df), quantidade, replace=False)
        df[coluna] = df[coluna].astype(object)  # permite misturar valores e vazios
        df.loc[linhas, coluna] = None
    return df


def inserir_linhas_em_branco(df, rng, quantidade):
    """Insere linhas totalmente vazias em posições aleatórias da planilha."""
    if quantidade == 0:
        return df
    posicoes = rng.choice(len(df), quantidade, replace=False)
    # Índice "posição + 0.5" coloca cada linha vazia entre duas linhas existentes.
    # Tudo como "object" (tipo genérico) para o pandas juntar as tabelas sem
    # tentar adivinhar tipos a partir das linhas vazias.
    vazias = pd.DataFrame(None, index=posicoes + 0.5, columns=df.columns, dtype=object)
    resultado = pd.concat([df.astype(object), vazias]).sort_index()
    return resultado.reset_index(drop=True)


def aplicar_sujeira(df, config, rng):
    """Aplica à planilha de uma filial todas as sujeiras definidas na sua configuração."""
    # A formatação vem antes de criar vazios/duplicatas, para não formatar células vazias.
    df = sujar_produtos(df, rng, config["proporcao_produtos_sujos"])
    df = formatar_datas(df, config["formato_data"], rng)
    if config["preco_texto"]:
        df = formatar_preco_texto(df)
    df = duplicar_linhas(df, rng, config["proporcao_duplicatas"])
    df = apagar_valores(df, rng, config["proporcao_faltantes"])
    df = inserir_linhas_em_branco(df, rng, config["linhas_em_branco"])
    df.columns = config["colunas"]
    return df


# ---------------------------------------------------------------------------
# Gravação
# ---------------------------------------------------------------------------

def salvar_planilha(df, caminho, linha_cabecalho=1, titulo=None):
    """Salva o DataFrame em .xlsx, com o cabeçalho na linha indicada.

    Se linha_cabecalho for maior que 1, escreve o título na célula A1
    (simulando relatórios exportados por sistemas com cabeçalho decorativo).
    """
    with pd.ExcelWriter(
        caminho,
        engine="openpyxl",
        date_format="DD/MM/YYYY",
        datetime_format="DD/MM/YYYY",
    ) as escritor:
        df.to_excel(escritor, sheet_name="Vendas", index=False, startrow=linha_cabecalho - 1)
        if linha_cabecalho > 1 and titulo:
            escritor.sheets["Vendas"]["A1"] = titulo


def parametros_periodo(mes=None):
    """Devolve (data inicial, nº de dias, vendas por filial, texto do período).

    mes=None: ano de 2026 inteiro; senão, só o mês indicado (pd.Period).
    """
    if mes is None:
        inicio = pd.Timestamp(f"{ANO}-01-01")
        dias = 366 if inicio.is_leap_year else 365
        return inicio, dias, VENDAS_POR_FILIAL, str(ANO)
    return mes.start_time, mes.days_in_month, VENDAS_POR_FILIAL_MES, nome_do_mes(mes)


def criar_gerador(indice_filial, mes=None):
    """Cria o gerador de números aleatórios de uma filial.

    Modo anual: semente SEED + índice (os mesmos dados de sempre).
    Modo mensal: semente [ano, mês, índice]. O numpy aceita uma lista como
    semente, e cada combinação gera números diferentes. Somar os números
    (ex.: 202610 + índice) faria meses e filiais diferentes colidirem.
    """
    if mes is None:
        return np.random.default_rng(SEED + indice_filial)
    return np.random.default_rng([mes.year, mes.month, indice_filial])


def executar(mes=None):
    """Gera e salva as planilhas sujas de todas as filiais (do ano ou de um mês)."""
    pasta_saida = pasta_raw(mes)
    pasta_saida.mkdir(parents=True, exist_ok=True)
    inicio, dias, vendas_por_filial, texto_periodo = parametros_periodo(mes)

    for indice, (filial, config) in enumerate(CONFIG_FILIAIS.items()):
        # Um gerador por filial: cada uma tem dados próprios, mas sempre reproduzíveis.
        rng = criar_gerador(indice, mes)

        df = gerar_vendas_limpas(filial, rng, vendas_por_filial, inicio, dias)
        df = aplicar_sujeira(df, config, rng)

        caminho = pasta_saida / config["arquivo"]
        titulo = f"Relatório de Vendas – Filial {filial} – {texto_periodo}"
        salvar_planilha(df, caminho, config["linha_cabecalho"], titulo)
        print(f"{config['arquivo']}: {len(df)} linhas")

    print(f"Planilhas salvas em {pasta_saida}")


def main():
    """Lê a opção --mes e gera as planilhas."""
    parser = argparse.ArgumentParser(description="Gera planilhas de vendas fictícias e sujas.")
    adicionar_argumento_mes(parser)
    executar(parser.parse_args().mes)


if __name__ == "__main__":
    main()
