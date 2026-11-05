# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Sobre o projeto
Automação de relatórios de vendas em Python. É um projeto de portfólio para
candidatura a estágio, então o código deve ser claro e fácil de ler para quem
avaliar.

## Regras do projeto
- Código (nomes de variáveis, funções etc.) e comentários em português
- Bibliotecas principais: pandas (manipulação de dados), openpyxl (leitura e
  escrita de Excel) e matplotlib (gráficos)
- Funções pequenas, cada uma com uma responsabilidade, e sempre com docstring
- Nunca commitar o arquivo `.env`; ele deve estar no `.gitignore`
- O autor é iniciante: explique as decisões antes de implementar
- Tudo que o código gera (dados, logs, gráficos, relatórios) vai para o
  `.gitignore`; só o código e a documentação vão para o GitHub

## Como rodar
As dependências estão no ambiente virtual `.venv/` (o `python3` do sistema não
tem pandas). Use `.venv/bin/python` ou ative antes com `source .venv/bin/activate`.

```bash
.venv/bin/python src/gerar_dados.py   # Fase 1: gera as planilhas sujas em data/raw/
.venv/bin/python src/limpeza.py       # Fase 2: lê, limpa e salva a base consolidada
```

## Estado atual
- **Fase 1 (concluída)** — `src/gerar_dados.py`: gera 5 planilhas fictícias
  (uma por filial: Rio, São Paulo, BH, Curitiba, Salvador) com problemas de
  qualidade diferentes em cada uma. Semente fixa, então os dados são sempre iguais.
- **Fase 2 (concluída)** — leitura e limpeza:
  - `src/leitura.py`: acha a linha do cabeçalho, traduz nomes de colunas para o
    padrão, adiciona `filial` e junta as planilhas
  - `src/limpeza.py`: converte tipos, padroniza textos, trata faltantes, remove
    duplicatas, calcula `faturamento`, salva
    `data/processed/vendas_consolidadas.csv` e registra tudo em `logs/limpeza.log`
  - Resultado atual: 1.534 linhas lidas → 1.414 na base limpa

Colunas da base limpa: `data, filial, produto, categoria, vendedor, quantidade,
preco_unitario, faturamento`.

## Decisões da Fase 2 (limpeza)
- **Cabeçalho**: detectado automaticamente (primeira linha, entre as 10 primeiras,
  com pelo menos 3 nomes de coluna conhecidos), sem número de linha fixo no código
- **Nomes de colunas**: normalizados (minúsculas, sem acento/pontuação) e
  traduzidos pelo dicionário `SINONIMOS_COLUNAS` em `leitura.py`; nome novo de
  coluna → acrescentar ali
- **Filial**: vem do nome do arquivo (dicionário `NOMES_FILIAIS`)
- **Datas**: formatos explícitos (data do Excel, número serial, `dd/mm/aaaa`,
  `aaaa-mm-dd`, por extenso "15 de março de 2026"); nunca deixar o pandas
  adivinhar, para não confundir dia com mês
- **Preços**: `"R$ 1.234,56"` → `1234.56` (ponto de milhar removido, vírgula vira ponto)
- **Textos**: espaços extras removidos; para maiúsculas/minúsculas adota-se a
  **grafia mais frequente** de cada nome (não usar `.title()`, que estraga
  "Pen drive 64GB")
- **Valores faltando — estratégia mista** (escolha do autor):
  - falta data, produto, quantidade ou preço → linha removida (sem eles não há faturamento)
  - falta categoria → preenchida com a categoria mais comum do mesmo produto
  - falta vendedor → `"Não informado"`
- **Duplicatas**: removidas **depois** da padronização, para pegar também as
  "quase duplicatas" (diferem só em espaços, maiúsculas ou formato). O log
  separa exatas de quase duplicatas. Limitação aceita: uma cópia que teve o
  vendedor apagado vira "Não informado" e não é reconhecida como duplicata
- **Saída**: CSV em `utf-8-sig` (para o Excel abrir os acentos), datas `aaaa-mm-dd`
- **Log**: módulo `logging`, no terminal e em `logs/limpeza.log` (sobrescrito a
  cada execução); cada etapa diz quantas linhas removeu/corrigiu, por filial
