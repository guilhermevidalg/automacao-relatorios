"""Geração do relatório de vendas em Excel a partir da base limpa.

Modo anual (sem --mes): lê data/processed/vendas_consolidadas.csv e cria
output/relatorio_vendas_2026.xlsx.
Modo mensal (--mes AAAA-MM): lê data/processed/vendas_AAAA-MM.csv e cria
output/relatorio_vendas_AAAA-MM.xlsx, com comparação com o mês anterior.

O relatório tem quatro abas:
1. Resumo: indicadores principais (KPIs);
2. Por filial: faturamento por filial e evolução mês a mês (ou semana a
   semana, no modo mensal), com gráficos;
3. Por produto: top 10 produtos por faturamento e por quantidade;
4. Base consolidada: os dados limpos, com filtro automático.

Divisão de tarefas: o pandas faz as contas e o openpyxl escreve as células,
aplica a formatação e cria os gráficos nativos do Excel.

Uso:
    .venv/bin/python src/relatorio.py
    .venv/bin/python src/relatorio.py --mes 2026-10
"""

import argparse
import logging
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from periodo import (
    MESES_ABREVIADOS,
    PASTA_PROJETO,
    adicionar_argumento_mes,
    arquivo_base,
    arquivo_relatorio,
    complemento_comando,
    mes_anterior,
    nome_do_mes,
)

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

# Formatos de exibição do Excel. A célula guarda o número puro (ex.: 1234.56)
# e o formato só muda a aparência. Num Excel em português, "R$" #,##0.00
# aparece como "R$ 1.234,56" (o Excel troca ponto e vírgula conforme o idioma).
FORMATO_REAIS = '"R$" #,##0.00'
FORMATO_REAIS_EIXO = '"R$" #,##0'  # sem centavos, para os eixos dos gráficos
FORMATO_INTEIRO = "#,##0"
FORMATO_PERCENTUAL = "0.0%"
# Variação com sinal sempre visível: +8,3% / -2,1% / 0,0%.
FORMATO_VARIACAO = "+0.0%;-0.0%;0.0%"
SEM_COMPARACAO = "sem comparação"
FORMATO_DATA = "dd/mm/yyyy"

# Formato de cada coluna, pelo nome que aparece no cabeçalho.
FORMATOS_COLUNAS = {
    "Data": FORMATO_DATA,
    "Quantidade": FORMATO_INTEIRO,
    "Preço unitário": FORMATO_REAIS,
    "Faturamento": FORMATO_REAIS,
    "Nº de vendas": FORMATO_INTEIRO,
    "% do total": FORMATO_PERCENTUAL,
}

# Nomes das colunas da base limpa como aparecem no relatório.
NOMES_COLUNAS_BASE = {
    "data": "Data",
    "filial": "Filial",
    "produto": "Produto",
    "categoria": "Categoria",
    "vendedor": "Vendedor",
    "quantidade": "Quantidade",
    "preco_unitario": "Preço unitário",
    "faturamento": "Faturamento",
}

# Visual: cabeçalho azul-escuro com texto branco; títulos em azul.
COR_PRINCIPAL = "1F4E78"
PREENCHIMENTO_CABECALHO = PatternFill(fill_type="solid", start_color=COR_PRINCIPAL)
FONTE_CABECALHO = Font(bold=True, color="FFFFFF")
FONTE_TITULO = Font(bold=True, size=14, color=COR_PRINCIPAL)

# Layout comum a todas as abas: título na linha 1, cabeçalho na linha 2.
LINHA_TITULO = 1
LINHA_CABECALHO = 2

LARGURA_MINIMA = 10
LARGURA_MAXIMA = 45
FOLGA_SETA_FILTRO = 4  # caracteres extras no cabeçalho para a setinha do filtro


# ---------------------------------------------------------------------------
# Log e leitura
# ---------------------------------------------------------------------------

def configurar_log():
    """Configura o log para mostrar mensagens curtas no terminal."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-7s | %(message)s",
        datefmt="%H:%M:%S",
    )


def ler_csv_base(caminho):
    """Lê um CSV de base limpa, com a coluna "data" já como data."""
    return pd.read_csv(caminho, parse_dates=["data"], encoding="utf-8-sig")


def carregar_base(mes=None):
    """Lê a base limpa do ano ou do mês. Se ela não existir, explica o que fazer e encerra."""
    caminho = arquivo_base(mes)
    if not caminho.exists():
        log.error(
            "Base limpa não encontrada em %s. Rode primeiro: .venv/bin/python src/limpeza.py%s",
            caminho.relative_to(PASTA_PROJETO), complemento_comando(mes),
        )
        sys.exit(1)  # código 1 = terminou com erro

    df = ler_csv_base(caminho)
    log.info("Base limpa lida: %d vendas", len(df))
    return df


# ---------------------------------------------------------------------------
# Cálculos (somente pandas)
# ---------------------------------------------------------------------------

def descrever_periodo(kpis, mes=None):
    """Texto do período, ex.: "01/10/2026 a 31/10/2026".

    No modo mensal usa o mês inteiro; no anual, da primeira à última venda.
    """
    if mes is None:
        inicio, fim = kpis["data_inicial"], kpis["data_final"]
    else:
        inicio, fim = mes.start_time, mes.end_time
    return f"{inicio:%d/%m/%Y} a {fim:%d/%m/%Y}"


def calcular_variacao(atual, anterior):
    """Variação percentual entre dois valores, como fração (0.083 = +8,3%).

    Devolve None se o valor anterior for zero (não dá para dividir por zero).
    """
    if not anterior:
        return None
    return (atual - anterior) / anterior


def comparar_com_mes_anterior(faturamento_atual, mes):
    """Compara o faturamento do mês com o do mês anterior.

    Só compara se a base limpa do mês anterior já existir (ela não é gerada
    automaticamente: num sistema real, inventar dados não é uma opção).

    Returns:
        Dicionário com "mes_anterior", "faturamento_anterior" e "variacao";
        os dois últimos ficam None quando não há dados para comparar.
    """
    anterior = mes_anterior(mes)
    caminho = arquivo_base(anterior)
    if not caminho.exists():
        log.warning(
            "Comparação: base de %s não encontrada (%s); relatório fica sem comparação",
            nome_do_mes(anterior), caminho.relative_to(PASTA_PROJETO),
        )
        return {"mes_anterior": anterior, "faturamento_anterior": None, "variacao": None}

    faturamento_anterior = round(ler_csv_base(caminho)["faturamento"].sum(), 2)
    variacao = calcular_variacao(faturamento_atual, faturamento_anterior)
    if variacao is not None:
        log.info("Comparação: faturamento %+.1f%% vs. %s", variacao * 100, nome_do_mes(anterior))
    return {
        "mes_anterior": anterior,
        "faturamento_anterior": faturamento_anterior,
        "variacao": variacao,
    }


def calcular_kpis(df):
    """Calcula os indicadores do Resumo. Cada linha da base é uma venda."""
    faturamento_filiais = df.groupby("filial")["faturamento"].sum()
    quantidade_produtos = df.groupby("produto")["quantidade"].sum()

    faturamento_total = round(df["faturamento"].sum(), 2)
    numero_vendas = len(df)
    return {
        "faturamento_total": faturamento_total,
        "numero_vendas": numero_vendas,
        "ticket_medio": round(faturamento_total / numero_vendas, 2),
        "melhor_filial": faturamento_filiais.idxmax(),
        "faturamento_melhor_filial": round(faturamento_filiais.max(), 2),
        "produto_mais_vendido": quantidade_produtos.idxmax(),
        "unidades_produto_mais_vendido": quantidade_produtos.max(),
        "data_inicial": df["data"].min(),
        "data_final": df["data"].max(),
    }


def faturamento_por_filial(df):
    """Faturamento, nº de vendas e participação de cada filial, da maior para a menor."""
    tabela = (
        df.groupby("filial")
        .agg(faturamento=("faturamento", "sum"), vendas=("faturamento", "size"))
        .sort_values("faturamento", ascending=False)
        .reset_index()
    )
    tabela["participacao"] = tabela["faturamento"] / tabela["faturamento"].sum()
    # Somas de números com centavos podem gerar "sujeira" (ex.: 1234.5600000001).
    tabela["faturamento"] = tabela["faturamento"].round(2)
    return tabela.rename(columns={
        "filial": "Filial",
        "faturamento": "Faturamento",
        "vendas": "Nº de vendas",
        "participacao": "% do total",
    })


def tabela_evolucao(df, agrupamento, ordem_filiais):
    """Faturamento por período (linhas) e filial (colunas), mais a coluna Total.

    Args:
        df: base limpa.
        agrupamento: Series com o período de cada venda (mês ou semana).
        ordem_filiais: ordem das colunas (a mesma do ranking de filiais).
    """
    tabela = df.pivot_table(
        index=agrupamento,
        columns="filial",
        values="faturamento",
        aggfunc="sum",
        fill_value=0,
    )
    tabela = tabela[ordem_filiais]
    tabela["Total"] = tabela.sum(axis=1)
    tabela.columns.name = None
    return tabela.round(2)


def evolucao_mensal(df, ordem_filiais):
    """Faturamento mês a mês: meses nas linhas, filiais nas colunas, mais o total."""
    tabela = tabela_evolucao(df, df["data"].dt.to_period("M"), ordem_filiais)
    tabela.index = [nome_do_mes(periodo) for periodo in tabela.index]
    return tabela.rename_axis("Mês").reset_index()


def rotulo_semana(numero_semana, mes):
    """Rótulo de uma semana do mês, ex.: semana 0 de out/2026 -> "01–07/out"."""
    inicio = numero_semana * 7 + 1
    fim = min(inicio + 6, mes.days_in_month)
    return f"{inicio:02d}–{fim:02d}/{MESES_ABREVIADOS[mes.month]}"


def evolucao_semanal(df, ordem_filiais, mes):
    """Faturamento semana a semana dentro do mês (dias 1–7, 8–14, 15–21, 22–28, 29–fim).

    Usada no modo mensal, em que a evolução mês a mês teria um ponto só.
    """
    semanas = (df["data"].dt.day - 1) // 7  # dia 1 a 7 -> 0, dia 8 a 14 -> 1...
    tabela = tabela_evolucao(df, semanas, ordem_filiais)
    # Garante todas as semanas do mês, mesmo uma semana sem nenhuma venda.
    numero_semanas = (mes.days_in_month + 6) // 7
    tabela = tabela.reindex(range(numero_semanas), fill_value=0)
    tabela.index = [rotulo_semana(semana, mes) for semana in tabela.index]
    return tabela.rename_axis("Semana").reset_index()


def top_produtos(df, coluna, n=10):
    """Os n produtos com maior valor em `coluna` ("faturamento" ou "quantidade")."""
    tabela = (
        df.groupby("produto")
        .agg(
            categoria=("categoria", "first"),
            faturamento=("faturamento", "sum"),
            quantidade=("quantidade", "sum"),
        )
        .sort_values(coluna, ascending=False)
        .head(n)
        .round(2)
        .reset_index()
    )
    # A coluna usada no ranking aparece logo depois do nome e da categoria.
    outra = "quantidade" if coluna == "faturamento" else "faturamento"
    tabela = tabela[["produto", "categoria", coluna, outra]]
    tabela.insert(0, "Posição", range(1, len(tabela) + 1))
    return tabela.rename(columns=NOMES_COLUNAS_BASE)


# ---------------------------------------------------------------------------
# Escrita e formatação (openpyxl)
# ---------------------------------------------------------------------------

def valor_para_excel(valor):
    """Converte tipos do pandas/numpy em tipos simples do Python, que o openpyxl entende."""
    if isinstance(valor, pd.Timestamp):
        return valor.to_pydatetime()
    if hasattr(valor, "item"):  # números do numpy (ex.: numpy.int64)
        return valor.item()
    return valor


def escrever_titulo(aba, texto, coluna=1):
    """Escreve o título de uma tabela na linha 1, em negrito e maior."""
    celula = aba.cell(row=LINHA_TITULO, column=coluna, value=texto)
    celula.font = FONTE_TITULO


def formatar_cabecalho(aba, linha, coluna_inicial, coluna_final):
    """Pinta o cabeçalho de uma tabela: fundo azul-escuro e texto branco em negrito."""
    for coluna in range(coluna_inicial, coluna_final + 1):
        celula = aba.cell(row=linha, column=coluna)
        celula.fill = PREENCHIMENTO_CABECALHO
        celula.font = FONTE_CABECALHO
        celula.alignment = Alignment(horizontal="center", vertical="center")


def escrever_tabela(aba, df, titulo, coluna_inicial=1, formatos=None):
    """Escreve um DataFrame com título (linha 1) e cabeçalho colorido (linha 2).

    Args:
        aba: aba do Excel onde a tabela será escrita.
        df: dados a escrever.
        titulo: texto que aparece acima da tabela.
        coluna_inicial: número da coluna onde a tabela começa (1 = coluna A).
        formatos: dicionário {nome da coluna: formato}; por padrão, FORMATOS_COLUNAS.

    Returns:
        Número da última linha escrita (útil para montar os gráficos).
    """
    formatos = FORMATOS_COLUNAS if formatos is None else formatos
    escrever_titulo(aba, titulo, coluna_inicial)

    for deslocamento, nome in enumerate(df.columns):
        aba.cell(row=LINHA_CABECALHO, column=coluna_inicial + deslocamento, value=nome)
    formatar_cabecalho(aba, LINHA_CABECALHO, coluna_inicial, coluna_inicial + len(df.columns) - 1)

    for numero_linha, linha in enumerate(df.itertuples(index=False), start=LINHA_CABECALHO + 1):
        for deslocamento, (nome, valor) in enumerate(zip(df.columns, linha)):
            celula = aba.cell(
                row=numero_linha,
                column=coluna_inicial + deslocamento,
                value=valor_para_excel(valor),
            )
            if nome in formatos:
                celula.number_format = formatos[nome]

    return LINHA_CABECALHO + len(df)


def tamanho_exibido(valor):
    """Estima quantos caracteres o valor ocupa na tela, já formatado."""
    if valor is None:
        return 0
    if isinstance(valor, float):
        return len(f"R$ {valor:,.2f}")  # pior caso: valor em reais
    if hasattr(valor, "strftime"):  # datas aparecem como dd/mm/aaaa
        return 10
    return len(str(valor))


def ajustar_largura_colunas(aba):
    """Ajusta a largura de cada coluna ao maior conteúdo (ignorando os títulos).

    Se a aba tem filtro automático, o cabeçalho ganha uma folga extra para
    caber o nome inteiro mais a setinha do filtro.
    """
    folga_cabecalho = FOLGA_SETA_FILTRO if aba.auto_filter.ref else 0
    for colunas in aba.iter_cols(min_row=LINHA_CABECALHO):
        cabecalho, *dados = colunas  # a primeira célula é o cabeçalho
        tamanho_cabecalho = tamanho_exibido(cabecalho.value)
        if tamanho_cabecalho:
            tamanho_cabecalho += folga_cabecalho
        maior = max([tamanho_cabecalho] + [tamanho_exibido(celula.value) for celula in dados])
        if maior == 0:  # coluna vazia (espaço entre tabelas)
            continue
        largura = min(max(maior + 2, LARGURA_MINIMA), LARGURA_MAXIMA)
        aba.column_dimensions[get_column_letter(colunas[0].column)].width = largura


def finalizar_aba(aba):
    """Ajusta as larguras e congela o painel abaixo do cabeçalho."""
    ajustar_largura_colunas(aba)
    # Congelar em A3 deixa as linhas 1 e 2 (título e cabeçalho) sempre visíveis.
    aba.freeze_panes = aba.cell(row=LINHA_CABECALHO + 1, column=1)


# ---------------------------------------------------------------------------
# Gráficos nativos do Excel
# ---------------------------------------------------------------------------

def mostrar_eixos(grafico):
    """Garante que os eixos apareçam (no openpyxl 3.1 eles vêm ocultos por padrão)."""
    grafico.x_axis.delete = False
    grafico.y_axis.delete = False


def definir_area_grafico(grafico, x, y, largura, altura):
    """Define onde fica a área do desenho (barras/linhas), em frações do gráfico.

    0 = topo/esquerda e 1 = base/direita. Ex.: y=0.15 faz o desenho começar
    a 15% do topo, deixando folga abaixo do título.
    Usa grafico.layout porque o openpyxl copia este valor para a área do
    gráfico ao salvar (um ajuste direto em plot_area.layout seria perdido).
    """
    grafico.layout = Layout(
        manualLayout=ManualLayout(
            layoutTarget="inner", xMode="edge", yMode="edge",
            x=x, y=y, w=largura, h=altura,
        )
    )


def grafico_barras_filiais(aba, ultima_linha, coluna_filial, coluna_faturamento):
    """Cria um gráfico de colunas com o faturamento de cada filial."""
    grafico = BarChart()
    grafico.type = "col"  # barras em pé
    grafico.title = "Faturamento por filial"
    grafico.y_axis.number_format = FORMATO_REAIS_EIXO
    grafico.legend = None  # só há uma série; a legenda seria redundante
    grafico.width, grafico.height = 16, 8  # em centímetros

    dados = Reference(aba, min_col=coluna_faturamento, min_row=LINHA_CABECALHO, max_row=ultima_linha)
    categorias = Reference(aba, min_col=coluna_filial, min_row=LINHA_CABECALHO + 1, max_row=ultima_linha)
    grafico.add_data(dados, titles_from_data=True)
    grafico.set_categories(categorias)
    mostrar_eixos(grafico)

    # Uma cor para cada barra (filial). Sem esta linha, o resultado dependeria
    # do padrão de cada programa (Excel, LibreOffice...).
    grafico.varyColors = True

    # Folga abaixo do título. Sem legenda, a área das barras pode ocupar mais altura que no gráfico de linha.
    grafico.title.overlay = False
    definir_area_grafico(grafico, x=0.12, y=0.15, largura=0.85, altura=0.72)
    return grafico


def grafico_linha_evolucao(aba, ultima_linha, coluna_periodo, primeira_filial, ultima_filial, titulo):
    """Cria um gráfico de linhas com a evolução por período (mês ou semana), uma linha por filial."""
    grafico = LineChart()
    grafico.title = titulo
    grafico.y_axis.number_format = FORMATO_REAIS_EIXO
    grafico.width, grafico.height = 24, 9

    dados = Reference(
        aba, min_col=primeira_filial, max_col=ultima_filial,
        min_row=LINHA_CABECALHO, max_row=ultima_linha,
    )
    periodos = Reference(aba, min_col=coluna_periodo, min_row=LINHA_CABECALHO + 1, max_row=ultima_linha)
    grafico.add_data(dados, titles_from_data=True)  # nome de cada filial vira o nome da linha
    grafico.set_categories(periodos)
    mostrar_eixos(grafico)

    # Os períodos são texto ("jan/2026", "01–07/out"), então o eixo precisa de uma
    # referência de texto (StrRef); o set_categories cria uma referência numérica
    # por padrão, e o Excel poderia mostrar 1, 2, 3... no lugar dos nomes.
    referencia_periodos = StrRef(f=str(periodos))

    # Cada filial é uma "série": linhas retas, com uma bolinha em cada período.
    for serie in grafico.series:
        serie.cat = AxDataSource(strRef=referencia_periodos)
        serie.smooth = False
        serie.marker.symbol = "circle"
        serie.marker.size = 6

    # Períodos sempre na base do gráfico, mesmo que algum valor fique perto de zero.
    grafico.x_axis.tickLblPos = "low"

    # Legenda embaixo, num espaço próprio (overlay=False: não fica por cima das linhas).
    grafico.legend.position = "b"
    grafico.legend.overlay = False
    grafico.title.overlay = False

    # Área das linhas: começa a 15% do topo (folga abaixo do título) e ocupa
    # 60% da altura, deixando espaço embaixo para os meses e a legenda.
    definir_area_grafico(grafico, x=0.10, y=0.15, largura=0.87, altura=0.60)
    return grafico


# ---------------------------------------------------------------------------
# Abas
# ---------------------------------------------------------------------------

def linhas_comparacao(comparacao):
    """Linhas do Resumo com a comparação com o mês anterior (indicador, valor, formato)."""
    nome_anterior = nome_do_mes(comparacao["mes_anterior"])
    if comparacao["variacao"] is None:
        return [(f"Variação vs. {nome_anterior}", SEM_COMPARACAO, None)]
    return [
        (f"Faturamento de {nome_anterior}", comparacao["faturamento_anterior"], FORMATO_REAIS),
        (f"Variação vs. {nome_anterior}", comparacao["variacao"], FORMATO_VARIACAO),
    ]


def criar_aba_resumo(planilha, kpis, mes=None, comparacao=None):
    """Aba "Resumo": tabela de indicadores com o formato certo para cada valor.

    No modo mensal (mes informado), o título traz o mês e a tabela ganha
    a comparação com o mês anterior.
    """
    aba = planilha.create_sheet("Resumo")
    periodo = descrever_periodo(kpis, mes)

    # (indicador, valor, formato de exibição)
    linhas = [
        ("Faturamento total", kpis["faturamento_total"], FORMATO_REAIS),
        ("Nº de vendas", kpis["numero_vendas"], FORMATO_INTEIRO),
        ("Ticket médio", kpis["ticket_medio"], FORMATO_REAIS),
        ("Melhor filial (faturamento)", kpis["melhor_filial"], None),
        ("Faturamento da melhor filial", kpis["faturamento_melhor_filial"], FORMATO_REAIS),
        ("Produto mais vendido (unidades)", kpis["produto_mais_vendido"], None),
        ("Unidades vendidas desse produto", kpis["unidades_produto_mais_vendido"], FORMATO_INTEIRO),
        ("Período analisado", periodo, None),
    ]
    if comparacao is not None:
        # Logo abaixo do faturamento total, que é o número comparado.
        linhas[1:1] = linhas_comparacao(comparacao)

    rotulo = "2026" if mes is None else nome_do_mes(mes)
    escrever_titulo(aba, f"Relatório de Vendas {rotulo} — Resumo")
    aba.cell(row=LINHA_CABECALHO, column=1, value="Indicador")
    aba.cell(row=LINHA_CABECALHO, column=2, value="Valor")
    formatar_cabecalho(aba, LINHA_CABECALHO, 1, 2)

    for numero_linha, (indicador, valor, formato) in enumerate(linhas, start=LINHA_CABECALHO + 1):
        aba.cell(row=numero_linha, column=1, value=indicador).font = Font(bold=True)
        celula = aba.cell(row=numero_linha, column=2, value=valor_para_excel(valor))
        celula.alignment = Alignment(horizontal="left")  # números e textos alinhados
        if formato:
            celula.number_format = formato

    finalizar_aba(aba)


def criar_aba_filiais(planilha, por_filial, evolucao, nome_evolucao):
    """Aba "Por filial": ranking das filiais, evolução no tempo e os dois gráficos.

    Args:
        planilha: arquivo Excel (Workbook) em construção.
        por_filial: tabela de faturamento_por_filial.
        evolucao: tabela de evolucao_mensal ou evolucao_semanal (1ª coluna = período).
        nome_evolucao: "mensal" ou "semanal", usado nos títulos.
    """
    aba = planilha.create_sheet("Por filial")

    # Tabela 1 nas colunas A-D; tabela 2 a partir da F (a coluna E fica de espaço).
    ultima_linha_filiais = escrever_tabela(aba, por_filial, "Faturamento por filial", coluna_inicial=1)

    coluna_periodo = len(por_filial.columns) + 2
    coluna_nome_periodo = evolucao.columns[0]  # "Mês" ou "Semana"
    formatos_evolucao = {nome: FORMATO_REAIS for nome in evolucao.columns if nome != coluna_nome_periodo}
    ultima_linha_evolucao = escrever_tabela(
        aba, evolucao, f"Evolução {nome_evolucao} do faturamento",
        coluna_inicial=coluna_periodo, formatos=formatos_evolucao,
    )

    # Gráficos abaixo das tabelas. A coluna "Total" fica fora do gráfico de
    # linhas, pois achataria as linhas das filiais.
    barras = grafico_barras_filiais(
        aba, ultima_linha_filiais,
        coluna_filial=1,
        coluna_faturamento=por_filial.columns.get_loc("Faturamento") + 1,
    )
    linhas = grafico_linha_evolucao(
        aba, ultima_linha_evolucao,
        coluna_periodo=coluna_periodo,
        primeira_filial=coluna_periodo + 1,
        ultima_filial=coluna_periodo + len(evolucao.columns) - 2,  # sem o período e o "Total"
        titulo=f"Evolução {nome_evolucao} do faturamento por filial",
    )
    linha_graficos = ultima_linha_evolucao + 3
    aba.add_chart(barras, f"A{linha_graficos}")
    aba.add_chart(linhas, f"A{linha_graficos + 18}")

    finalizar_aba(aba)


def criar_aba_produtos(planilha, top_faturamento, top_quantidade):
    """Aba "Por produto": top 10 por faturamento e top 10 por quantidade, lado a lado."""
    aba = planilha.create_sheet("Por produto")
    escrever_tabela(aba, top_faturamento, "Top 10 por faturamento", coluna_inicial=1)
    escrever_tabela(
        aba, top_quantidade, "Top 10 por quantidade",
        coluna_inicial=len(top_faturamento.columns) + 2,
    )
    finalizar_aba(aba)


def criar_aba_base(planilha, df):
    """Aba "Base consolidada": todos os dados limpos, com filtro automático."""
    aba = planilha.create_sheet("Base consolidada")
    base = df.rename(columns=NOMES_COLUNAS_BASE)
    ultima_linha = escrever_tabela(aba, base, f"Base consolidada ({len(base):,} vendas)".replace(",", "."))

    # Filtro automático: setinhas no cabeçalho para filtrar/ordenar cada coluna.
    ultima_coluna = get_column_letter(len(base.columns))
    aba.auto_filter.ref = f"A{LINHA_CABECALHO}:{ultima_coluna}{ultima_linha}"
    finalizar_aba(aba)


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------

def gerar_relatorio(df, caminho, mes=None):
    """Monta as quatro abas e salva o arquivo Excel.

    Args:
        df: base limpa.
        caminho: arquivo .xlsx de saída.
        mes: período mensal (modo mensal) ou None (ano inteiro).
    """
    planilha = Workbook()
    planilha.remove(planilha.active)  # remove a aba vazia que vem por padrão

    kpis = calcular_kpis(df)
    por_filial = faturamento_por_filial(df)
    ordem_filiais = por_filial["Filial"].tolist()

    if mes is None:
        comparacao = None
        evolucao, nome_evolucao = evolucao_mensal(df, ordem_filiais), "mensal"
    else:
        comparacao = comparar_com_mes_anterior(kpis["faturamento_total"], mes)
        evolucao, nome_evolucao = evolucao_semanal(df, ordem_filiais, mes), "semanal"

    criar_aba_resumo(planilha, kpis, mes, comparacao)
    criar_aba_filiais(planilha, por_filial, evolucao, nome_evolucao)
    criar_aba_produtos(planilha, top_produtos(df, "faturamento"), top_produtos(df, "quantidade"))
    criar_aba_base(planilha, df)
    log.info("Abas criadas: %s", ", ".join(planilha.sheetnames))

    caminho.parent.mkdir(parents=True, exist_ok=True)
    planilha.save(caminho)
    log.info("Relatório salvo em %s", caminho.relative_to(PASTA_PROJETO))


def executar(mes=None):
    """Lê a base limpa (do ano ou do mês) e gera o relatório em Excel."""
    df = carregar_base(mes)
    gerar_relatorio(df, arquivo_relatorio(mes), mes)


def main():
    """Lê a opção --mes e gera o relatório."""
    parser = argparse.ArgumentParser(description="Gera o relatório de vendas em Excel.")
    adicionar_argumento_mes(parser)
    argumentos = parser.parse_args()
    configurar_log()
    executar(argumentos.mes)


if __name__ == "__main__":
    main()
