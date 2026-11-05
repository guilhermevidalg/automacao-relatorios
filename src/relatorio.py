"""Geração do relatório de vendas em Excel a partir da base limpa.

Lê data/processed/vendas_consolidadas.csv (gerada por src/limpeza.py) e cria
output/relatorio_vendas_2026.xlsx com quatro abas:

1. Resumo: indicadores principais (KPIs);
2. Por filial: faturamento por filial e evolução mês a mês, com gráficos;
3. Por produto: top 10 produtos por faturamento e por quantidade;
4. Base consolidada: os dados limpos, com filtro automático.

Divisão de tarefas: o pandas faz as contas e o openpyxl escreve as células,
aplica a formatação e cria os gráficos nativos do Excel.

Uso:
    .venv/bin/python src/relatorio.py
"""

import logging
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

PASTA_PROJETO = Path(__file__).resolve().parent.parent
ARQUIVO_BASE = PASTA_PROJETO / "data" / "processed" / "vendas_consolidadas.csv"
ARQUIVO_RELATORIO = PASTA_PROJETO / "output" / "relatorio_vendas_2026.xlsx"

# Formatos de exibição do Excel. A célula guarda o número puro (ex.: 1234.56)
# e o formato só muda a aparência. Num Excel em português, "R$" #,##0.00
# aparece como "R$ 1.234,56" (o Excel troca ponto e vírgula conforme o idioma).
FORMATO_REAIS = '"R$" #,##0.00'
FORMATO_REAIS_EIXO = '"R$" #,##0'  # sem centavos, para os eixos dos gráficos
FORMATO_INTEIRO = "#,##0"
FORMATO_PERCENTUAL = "0.0%"
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

# Meses em português definidos aqui, para não depender do idioma do computador.
MESES_ABREVIADOS = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez",
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


def carregar_base(caminho=ARQUIVO_BASE):
    """Lê a base limpa. Se ela não existir, explica o que fazer e encerra."""
    if not caminho.exists():
        log.error(
            "Base limpa não encontrada em %s. "
            "Rode primeiro: .venv/bin/python src/limpeza.py",
            caminho.relative_to(PASTA_PROJETO),
        )
        sys.exit(1)  # código 1 = terminou com erro

    df = pd.read_csv(caminho, parse_dates=["data"], encoding="utf-8-sig")
    log.info("Base limpa lida: %d vendas", len(df))
    return df


# ---------------------------------------------------------------------------
# Cálculos (somente pandas)
# ---------------------------------------------------------------------------

def nome_do_mes(periodo):
    """Converte um período mensal do pandas em texto, ex.: 2026-03 -> "mar/2026"."""
    return f"{MESES_ABREVIADOS[periodo.month]}/{periodo.year}"


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


def evolucao_mensal(df, ordem_filiais):
    """Faturamento mês a mês: meses nas linhas, filiais nas colunas, mais o total."""
    tabela = df.pivot_table(
        index=df["data"].dt.to_period("M"),
        columns="filial",
        values="faturamento",
        aggfunc="sum",
        fill_value=0,
    )
    tabela = tabela[ordem_filiais]  # mesma ordem do ranking de filiais
    tabela["Total"] = tabela.sum(axis=1)
    tabela = tabela.round(2)
    tabela.index = [nome_do_mes(periodo) for periodo in tabela.index]
    tabela.columns.name = None
    return tabela.rename_axis("Mês").reset_index()


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


def grafico_linha_mensal(aba, ultima_linha, coluna_mes, primeira_filial, ultima_filial):
    """Cria um gráfico de linhas com a evolução mensal, uma linha por filial."""
    grafico = LineChart()
    grafico.title = "Evolução mensal do faturamento por filial"
    grafico.y_axis.number_format = FORMATO_REAIS_EIXO
    grafico.width, grafico.height = 24, 9

    dados = Reference(
        aba, min_col=primeira_filial, max_col=ultima_filial,
        min_row=LINHA_CABECALHO, max_row=ultima_linha,
    )
    meses = Reference(aba, min_col=coluna_mes, min_row=LINHA_CABECALHO + 1, max_row=ultima_linha)
    grafico.add_data(dados, titles_from_data=True)  # nome de cada filial vira o nome da linha
    grafico.set_categories(meses)
    mostrar_eixos(grafico)

    # Os meses são texto ("jan/2026"), então o eixo precisa de uma referência de
    # texto (StrRef); o set_categories cria uma referência numérica por padrão,
    # e o Excel poderia mostrar 1, 2, 3... no lugar dos nomes.
    referencia_meses = StrRef(f=str(meses))

    # Cada filial é uma "série": linhas retas, com uma bolinha em cada mês.
    for serie in grafico.series:
        serie.cat = AxDataSource(strRef=referencia_meses)
        serie.smooth = False
        serie.marker.symbol = "circle"
        serie.marker.size = 6

    # Meses sempre na base do gráfico, mesmo que algum valor fique perto de zero.
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

def criar_aba_resumo(planilha, kpis):
    """Aba "Resumo": tabela de indicadores com o formato certo para cada valor."""
    aba = planilha.create_sheet("Resumo")
    periodo = f"{kpis['data_inicial']:%d/%m/%Y} a {kpis['data_final']:%d/%m/%Y}"

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

    escrever_titulo(aba, "Relatório de Vendas 2026 — Resumo")
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


def criar_aba_filiais(planilha, por_filial, mensal):
    """Aba "Por filial": ranking das filiais, evolução mensal e os dois gráficos."""
    aba = planilha.create_sheet("Por filial")

    # Tabela 1 nas colunas A-D; tabela 2 a partir da F (a coluna E fica de espaço).
    ultima_linha_filiais = escrever_tabela(aba, por_filial, "Faturamento por filial", coluna_inicial=1)

    coluna_mes = len(por_filial.columns) + 2
    formatos_mensal = {nome: FORMATO_REAIS for nome in mensal.columns if nome != "Mês"}
    ultima_linha_mensal = escrever_tabela(
        aba, mensal, "Evolução mensal do faturamento",
        coluna_inicial=coluna_mes, formatos=formatos_mensal,
    )

    # Gráficos abaixo das tabelas. A coluna "Total" fica fora do gráfico de
    # linhas, pois achataria as linhas das filiais.
    barras = grafico_barras_filiais(
        aba, ultima_linha_filiais,
        coluna_filial=1,
        coluna_faturamento=por_filial.columns.get_loc("Faturamento") + 1,
    )
    linhas = grafico_linha_mensal(
        aba, ultima_linha_mensal,
        coluna_mes=coluna_mes,
        primeira_filial=coluna_mes + 1,
        ultima_filial=coluna_mes + len(mensal.columns) - 2,  # sem "Mês" e "Total"
    )
    linha_graficos = ultima_linha_mensal + 3
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

def gerar_relatorio(df, caminho=ARQUIVO_RELATORIO):
    """Monta as quatro abas e salva o arquivo Excel."""
    planilha = Workbook()
    planilha.remove(planilha.active)  # remove a aba vazia que vem por padrão

    por_filial = faturamento_por_filial(df)
    criar_aba_resumo(planilha, calcular_kpis(df))
    criar_aba_filiais(planilha, por_filial, evolucao_mensal(df, por_filial["Filial"].tolist()))
    criar_aba_produtos(planilha, top_produtos(df, "faturamento"), top_produtos(df, "quantidade"))
    criar_aba_base(planilha, df)
    log.info("Abas criadas: %s", ", ".join(planilha.sheetnames))

    caminho.parent.mkdir(parents=True, exist_ok=True)
    planilha.save(caminho)
    log.info("Relatório salvo em %s", caminho.relative_to(PASTA_PROJETO))


def main():
    """Lê a base limpa e gera o relatório em Excel."""
    configurar_log()
    df = carregar_base()
    gerar_relatorio(df)


if __name__ == "__main__":
    main()
