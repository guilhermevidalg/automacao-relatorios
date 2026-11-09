# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Sobre o projeto
Automação de relatórios de vendas em Python. É um projeto de portfólio para
candidatura a estágio, então o código deve ser claro e fácil de ler para quem
avaliar.

## Regras do projeto
- Código (nomes de variáveis, funções etc.) e comentários em português
- Bibliotecas principais: pandas (manipulação de dados) e openpyxl (leitura e
  escrita de Excel, inclusive os gráficos nativos). O matplotlib não é usado
  (foi tirado do `requirements.txt`)
- Funções pequenas, cada uma com uma responsabilidade, e sempre com docstring
- Nunca commitar o arquivo `.env`; ele deve estar no `.gitignore`
- O autor é iniciante: explique as decisões antes de implementar
- Tudo que o código gera (dados, logs, gráficos, relatórios) vai para o
  `.gitignore`; só o código e a documentação vão para o GitHub

## Como rodar
As dependências estão no ambiente virtual `.venv/` (o `python3` do sistema não
tem pandas). Use `.venv/bin/python` ou ative antes com `source .venv/bin/activate`.

```bash
# Pipeline mensal (Fase 5A): tudo com um comando só
.venv/bin/python main.py                          # mês anterior ao atual, envia o e-mail
.venv/bin/python main.py --mes 2026-10 --dry-run  # e-mail só em prévia
.venv/bin/python main.py --mes 2026-10 --sem-email  # para no relatório
.venv/bin/python -m pytest -v                     # testes

# Cada etapa sozinha (sem --mes = ano de 2026 inteiro; com --mes AAAA-MM = um mês)
.venv/bin/python src/gerar_dados.py   # Fase 1: gera as planilhas sujas em data/raw/
.venv/bin/python src/limpeza.py       # Fase 2: lê, limpa e salva a base consolidada
.venv/bin/python src/relatorio.py     # Fase 3: gera output/relatorio_vendas_2026.xlsx
.venv/bin/python src/email_sender.py --dry-run  # Fase 4: mostra o e-mail sem enviar
.venv/bin/python src/email_sender.py            # Fase 4: envia o relatório pelo Gmail

# Fase 6: disparar o workflow do GitHub Actions manualmente (mes é opcional)
gh workflow run relatorio-mensal.yml -f mes=2026-10
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
- **Fase 3 (concluída)** — `src/relatorio.py`: lê a base limpa e gera
  `output/relatorio_vendas_2026.xlsx` com as abas Resumo, Por filial,
  Por produto e Base consolidada
- **Fase 4 (concluída)** — `src/email_sender.py`: envia o
  relatório como anexo pelo Gmail, com resumo de KPIs em HTML; `--dry-run` só mostra
- **Fase 5A (concluída)** — pipeline mensal: `main.py` roda gerar dados → limpeza →
  relatório → e-mail para um mês (`--mes`, `--dry-run`, `--sem-email`), com
  comparação com o mês anterior, `logs/pipeline.log` e testes em `tests/`
- **Fase 5B (concluída)** — `README.md` para o GitHub (público: recrutadores de
  estágio em dados): problema, diagrama Mermaid do pipeline, imagens de
  `docs/img/` (`antes_planilha_suja.png`, `graficos.png`, `email_recebido.png`),
  tabela antes vs. depois com números dos logs, como rodar, estrutura de pastas,
  decisões técnicas, testes e próximos passos. Se mudar comandos, números da
  limpeza ou a quantidade de testes, atualizar o README também
- **Fase 6 (concluída)** — `.github/workflows/relatorio-mensal.yml`: GitHub Actions
  roda o pipeline todo dia 1º (e manualmente), sempre em `--dry-run`, sem Secrets,
  e salva relatório, prévia do e-mail e logs como artifacts. README com badge e
  seção "Agendamento automático". 118 testes (novo `tests/test_email_sender.py`)

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

## Decisões da Fase 3 (relatório)
- **pandas calcula, openpyxl formata**: contas em funções só de pandas; escrita,
  estilos e gráficos com openpyxl
- **Gráficos nativos do Excel** (`openpyxl.chart`), não imagens do matplotlib:
  barras (faturamento por filial) e linhas (evolução mensal, sem a coluna Total).
  No openpyxl 3.1 os eixos vêm ocultos: usar `eixo.delete = False`.
  Gráfico de barras: uma cor por filial (`varyColors = True`, explícito) e valores
  só no eixo Y (sem rótulos em cima das barras, por escolha do autor).
  Os dois gráficos usam `definir_area_grafico` para dar folga abaixo do título.
  Gráfico de linha: sem suavização, marcadores por mês, legenda embaixo
  (`overlay=False`). Pegadinhas do openpyxl: categorias de texto (meses) precisam
  de `StrRef` (o `set_categories` cria referência numérica), e o layout manual vai
  em `grafico.layout` (o `plot_area.layout` é sobrescrito ao salvar)
- **Valores como número** + formato de exibição `"R$" #,##0.00` (o Excel em
  português mostra `R$ 1.234,56`); somas arredondadas para 2 casas
- **KPIs**: cada linha = 1 venda; ticket médio = faturamento ÷ nº de vendas;
  produto mais vendido = mais **unidades**; melhor filial = maior faturamento
- **Layout das abas**: título na linha 1, cabeçalho azul na linha 2, painel
  congelado em A3; tabelas lado a lado (uma coluna vazia entre elas)
- **Meses** em português por dicionário próprio (`jan/2026`), sem `locale`
- **Base ausente**: mensagem pedindo para rodar `src/limpeza.py` e `sys.exit(1)`

## Decisões da Fase 4 (e-mail)
- **SMTP do Gmail** com `smtplib.SMTP_SSL` na porta 465 (criptografado desde o início)
- **Credenciais** só no `.env` (modelo em `.env.example`), lidas com `python-dotenv`
  e `override=True` (o `.env` prevalece sobre variáveis do terminal). Variáveis:
  `GMAIL_REMETENTE`, `GMAIL_SENHA_APP` (senha de app; espaços removidos),
  `EMAIL_DESTINATARIOS` (separados por vírgula)
- **KPIs** reaproveitados de `relatorio.calcular_kpis` (mesmos números da aba Resumo)
- **Corpo**: texto simples + HTML alternativo (`EmailMessage.add_alternative`),
  estilos inline (programas de e-mail ignoram `<style>`); valores em formato
  brasileiro (`R$ 1.234,56`)
- **`--dry-run`**: não exige a senha nem o `.env` (desde a Fase 6); remetente ou
  destinatário vazios viram `remetente@example.com` / `destinatario@example.com`
  (`preencher_com_exemplos`, com aviso no log). Mostra a prévia no terminal e salva
  o HTML em `output/previa_email.html`. O envio real continua exigindo tudo
- **Erros**: cada caso (sem `.env`, variável vazia, relatório ausente, login
  recusado, destinatário recusado, sem conexão) tem mensagem própria e `sys.exit(1)`

## Decisões da Fase 5A (pipeline mensal)
- **Dois modos**: sem `--mes` cada script faz exatamente o que fazia antes (ano de
  2026, `data/raw/`, `vendas_consolidadas.csv`, `relatorio_vendas_2026.xlsx`,
  1.534 → 1.414 linhas). Com `--mes AAAA-MM`: `data/raw/AAAA-MM/`,
  `data/processed/vendas_AAAA-MM.csv`, `output/relatorio_vendas_AAAA-MM.xlsx`
- **`src/periodo.py`**: tudo de mês (validar `AAAA-MM`, mês anterior, nomes em
  português) e todos os caminhos de arquivo (`pasta_raw`, `arquivo_base`,
  `arquivo_relatorio`, com `mes=None` = modo anual)
- **Cada script tem `executar(mes=None, ...)`**; o `main()` do script só lê os
  argumentos e chama `executar`. O `main.py` importa os módulos (não usa
  subprocessos) e chama `executar` de cada etapa dentro de `try/except` que pega
  `Exception` e `SystemExit`; falha → mensagem com a etapa e `sys.exit(1)`
- **Padrão de `--mes` no `main.py`**: o mês anterior ao atual. `--dry-run` e
  `--sem-email` são mutuamente exclusivos (argparse)
- **Dados mensais**: 100 vendas por filial; semente `[ano, mês, índice da filial]`
  (lista, para meses e filiais nunca colidirem); modo anual mantém `SEED + índice`
- **Limpeza mensal**: etapa extra `remover_fora_do_mes` (vendas de outro mês são
  removidas e registradas no log)
- **Comparação com o mês anterior** (escolha do autor): só se
  `data/processed/vendas_<mês anterior>.csv` existir; senão "sem comparação" +
  aviso no log. Nada é gerado automaticamente (num sistema real não há gerador).
  Variação = (atual − anterior) ÷ anterior; formato Excel `+0.0%;-0.0%;0.0%`; no
  e-mail `+8,3%`. Resumo e e-mail usam a mesma função `comparar_com_mes_anterior`
- **Relatório mensal**: título com o mês (`out/2026`), período = mês inteiro, e a
  aba Por filial mostra evolução **semanal** (blocos 01–07, 08–14, 15–21, 22–28,
  29–fim), pois a mensal teria um ponto só
- **Logs**: `logs/pipeline.log` em modo append, só com as mensagens do logger
  `"pipeline"` (início, fim e duração de cada etapa, falhas com traceback).
  `logs/limpeza.log` continua: o `FileHandler` é ligado dentro de
  `limpeza.executar()` e desligado no fim (`basicConfig` só funciona uma vez por processo)
- **Testes**: pytest, `pytest.ini` com `pythonpath = src`. `tests/test_limpeza.py`,
  `tests/test_leitura.py` (inclui: todo nome de coluna do gerador é reconhecido
  pelo leitor), `tests/test_periodo.py` e `tests/test_email_sender.py` (Fase 6:
  prévia sem `.env`, envio real exigindo tudo)

## Decisões da Fase 6 (GitHub Actions)
- **Sempre `--dry-run` na nuvem, sem Secrets** (escolha do autor): nenhuma senha ou
  e-mail cadastrado no GitHub. O envio real é só local, com o `.env`
- **Cron em UTC**: `0 11 1 * *` = dia 1º às 8h de Brasília (sem horário de verão).
  O GitHub pode atrasar alguns minutos e desliga agendamentos de repositórios sem
  atividade há 60 dias
- **Gatilho manual** (`workflow_dispatch`) com o campo `mes` opcional (vazio = mês
  anterior). O valor entra por variável de ambiente (`MES_INFORMADO`), nunca com
  `${{ }}` direto no `run`, para evitar injeção de comando; o formato é validado
  no passo. `MES` e `MES_ANTERIOR` vão para `$GITHUB_ENV` (calculados com `date -d`)
- **Comparação na nuvem** (escolha do autor): a máquina começa vazia, então o
  workflow roda antes `main.py --mes <anterior> --sem-email` só para criar a base,
  e depois `main.py --mes <mês> --dry-run`
- **Passos**: checkout → Python 3.12 com `cache: pip` → `pip install -r
  requirements.txt` → pytest (falhou, para) → meses → mês anterior → pipeline →
  `upload-artifact` com `if: always()` (relatório do mês, `previa_email.html`,
  `logs/`), artifact `relatorio-AAAA-MM`
- **`permissions: contents: read`**: o workflow só lê o repositório
