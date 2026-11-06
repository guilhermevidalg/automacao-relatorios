"""Pipeline mensal de relatórios de vendas: tudo com um comando só.

Roda as etapas em ordem, para um mês:
    1. gerar dados  -> data/raw/AAAA-MM/
    2. limpeza      -> data/processed/vendas_AAAA-MM.csv
    3. relatório    -> output/relatorio_vendas_AAAA-MM.xlsx
    4. e-mail       -> envia o relatório (ou só gera a prévia, com --dry-run)

Se uma etapa falhar, o pipeline para, diz qual etapa falhou e termina com
código de erro 1. Início, fim e duração de cada etapa ficam registrados em
logs/pipeline.log (o arquivo acumula o histórico de todas as execuções).

Uso:
    .venv/bin/python main.py                       # mês anterior ao atual, envia o e-mail
    .venv/bin/python main.py --mes 2026-10         # um mês específico
    .venv/bin/python main.py --mes 2026-10 --dry-run    # e-mail só em prévia
    .venv/bin/python main.py --mes 2026-10 --sem-email  # para no relatório
"""

import argparse
import logging
import sys
import time
from pathlib import Path

# Os módulos do projeto estão em src/. sys.path é a lista de pastas onde o
# Python procura módulos; acrescentamos src/ para poder importá-los daqui.
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

import email_sender  # noqa: E402 (import depois de ajustar o sys.path, de propósito)
import gerar_dados  # noqa: E402
import limpeza  # noqa: E402
import relatorio  # noqa: E402
from periodo import PASTA_PROJETO, ler_mes, mes_padrao, nome_do_mes  # noqa: E402

log = logging.getLogger("pipeline")

ARQUIVO_LOG_PIPELINE = PASTA_PROJETO / "logs" / "pipeline.log"
FORMATO_LOG = "%(asctime)s | %(levelname)-7s | %(message)s"


# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

def ler_argumentos(argumentos=None):
    """Lê as opções da linha de comando (--mes, --dry-run, --sem-email)."""
    parser = argparse.ArgumentParser(
        description="Pipeline mensal: gera os dados, limpa, monta o relatório e envia por e-mail.",
    )
    parser.add_argument(
        "--mes",
        type=ler_mes,
        default=mes_padrao(),
        help="mês no formato AAAA-MM (padrão: o mês anterior ao atual)",
    )
    # Grupo mutuamente exclusivo: o argparse recusa as duas opções juntas.
    modo_email = parser.add_mutually_exclusive_group()
    modo_email.add_argument(
        "--dry-run",
        action="store_true",
        help="roda tudo normalmente, mas o e-mail só gera a prévia (não envia)",
    )
    modo_email.add_argument(
        "--sem-email",
        action="store_true",
        help="para no relatório: não envia o e-mail nem gera a prévia",
    )
    return parser.parse_args(argumentos)


def configurar_log():
    """Mostra tudo no terminal e grava as mensagens do pipeline em logs/pipeline.log.

    Só o logger "pipeline" grava no arquivo: assim ele guarda apenas o resumo
    de cada execução (etapas e durações). Os detalhes de cada etapa aparecem
    no terminal (e os da limpeza também em logs/limpeza.log).
    """
    logging.basicConfig(level=logging.INFO, format=FORMATO_LOG, datefmt="%H:%M:%S")

    ARQUIVO_LOG_PIPELINE.parent.mkdir(parents=True, exist_ok=True)
    # mode="a" (append): acrescenta ao fim do arquivo, sem apagar execuções anteriores.
    arquivo = logging.FileHandler(ARQUIVO_LOG_PIPELINE, mode="a", encoding="utf-8")
    arquivo.setFormatter(logging.Formatter(FORMATO_LOG, datefmt="%Y-%m-%d %H:%M:%S"))
    log.addHandler(arquivo)


def descrever_modo(argumentos):
    """Texto curto sobre o que acontece com o e-mail nesta execução."""
    if argumentos.sem_email:
        return "sem e-mail"
    if argumentos.dry_run:
        return "e-mail só em prévia (dry-run)"
    return "envio do e-mail"


def montar_etapas(argumentos):
    """Lista de (nome da etapa, função a executar), na ordem certa.

    `lambda` cria uma função curta "embrulhando" a chamada com os argumentos,
    para que todas as etapas possam ser chamadas do mesmo jeito: etapa().
    """
    mes = argumentos.mes
    etapas = [
        ("Gerar dados", lambda: gerar_dados.executar(mes)),
        ("Limpeza", lambda: limpeza.executar(mes)),
        ("Relatório", lambda: relatorio.executar(mes)),
    ]
    if not argumentos.sem_email:
        etapas.append(("E-mail", lambda: email_sender.executar(mes, dry_run=argumentos.dry_run)))
    return etapas


# ---------------------------------------------------------------------------
# Execução
# ---------------------------------------------------------------------------

def formatar_duracao(segundos):
    """Formata uma duração em segundos com vírgula decimal: 2.345 -> "2,3 s"."""
    return f"{segundos:.1f} s".replace(".", ",")


def executar_etapa(numero, total, nome, funcao):
    """Executa uma etapa, registrando início, fim e duração.

    Returns:
        True se a etapa terminou bem; False se falhou.
    """
    rotulo = f"Etapa {numero}/{total} ({nome})"
    log.info("%s: início", rotulo)
    inicio = time.perf_counter()  # relógio de alta precisão, próprio para medir durações

    try:
        funcao()
    except SystemExit as erro:
        # Os scripts usam sys.exit(1) nos erros que já conhecem (ex.: .env ausente),
        # depois de explicar o problema no log. Código 0 ou None = saiu sem erro.
        if erro.code not in (0, None):
            log.error("%s falhou após %s (detalhes nas mensagens acima)",
                      rotulo, formatar_duracao(time.perf_counter() - inicio))
            return False
    except Exception as erro:  # erro inesperado: registra o tipo, a mensagem e o rastro
        log.error("%s falhou após %s: %s: %s",
                  rotulo, formatar_duracao(time.perf_counter() - inicio),
                  type(erro).__name__, erro, exc_info=True)
        return False

    log.info("%s: concluída em %s", rotulo, formatar_duracao(time.perf_counter() - inicio))
    return True


def main(argumentos=None):
    """Roda todas as etapas para o mês escolhido; para na primeira que falhar."""
    argumentos = ler_argumentos(argumentos)
    configurar_log()

    etapas = montar_etapas(argumentos)
    log.info("===== Pipeline iniciado: %s (%s) =====",
             nome_do_mes(argumentos.mes), descrever_modo(argumentos))
    inicio = time.perf_counter()

    for numero, (nome, funcao) in enumerate(etapas, start=1):
        if not executar_etapa(numero, len(etapas), nome, funcao):
            log.error("===== Pipeline interrompido na etapa %d/%d (%s) após %s =====",
                      numero, len(etapas), nome, formatar_duracao(time.perf_counter() - inicio))
            sys.exit(1)

    log.info("===== Pipeline concluído em %s =====", formatar_duracao(time.perf_counter() - inicio))


if __name__ == "__main__":
    main()
