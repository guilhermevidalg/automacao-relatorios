"""Envio do relatório de vendas por e-mail, pelo Gmail.

Anexa o relatório em Excel (gerado por src/relatorio.py) e escreve no corpo do
e-mail, em HTML, um resumo dos principais indicadores (KPIs). No modo mensal
(--mes), o assunto e o texto citam o mês e o resumo traz a variação do
faturamento em relação ao mês anterior.

As credenciais ficam no arquivo .env (nunca no código). Para criá-lo:
    cp .env.example .env
e preencha os valores. O Gmail exige uma "senha de app", não a senha normal.

Uso:
    .venv/bin/python src/email_sender.py                          # envia (ano inteiro)
    .venv/bin/python src/email_sender.py --dry-run                # só mostra o que seria enviado
    .venv/bin/python src/email_sender.py --mes 2026-10 --dry-run  # relatório de outubro
"""

import argparse
import html
import logging
import os
import smtplib
import sys
from email.message import EmailMessage

from dotenv import load_dotenv

from periodo import (
    PASTA_PROJETO,
    adicionar_argumento_mes,
    arquivo_relatorio,
    complemento_comando,
    mes_por_extenso,
    nome_do_mes,
)

# Reaproveita o relatório: mesma cor e os mesmos cálculos (KPIs e comparação),
# para o e-mail mostrar exatamente os números da aba Resumo.
from relatorio import (
    COR_PRINCIPAL,
    SEM_COMPARACAO,
    calcular_kpis,
    carregar_base,
    comparar_com_mes_anterior,
    configurar_log,
    descrever_periodo,
)

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

ARQUIVO_ENV = PASTA_PROJETO / ".env"
ARQUIVO_PREVIA = PASTA_PROJETO / "output" / "previa_email.html"

# Porta 465 = conexão já criptografada desde o início (SSL).
SERVIDOR_SMTP = "smtp.gmail.com"
PORTA_SMTP = 465
TEMPO_LIMITE_SEGUNDOS = 30

# Tipo oficial (MIME) de um arquivo .xlsx, para o anexo abrir no Excel.
TIPO_ANEXO_XLSX = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Variáveis esperadas no .env.
VAR_REMETENTE = "GMAIL_REMETENTE"
VAR_SENHA = "GMAIL_SENHA_APP"
VAR_DESTINATARIOS = "EMAIL_DESTINATARIOS"


# ---------------------------------------------------------------------------
# Configuração e verificações
# ---------------------------------------------------------------------------

def encerrar_com_erro(mensagem, *argumentos):
    """Mostra a mensagem de erro no log e encerra o programa com código 1."""
    log.error(mensagem, *argumentos)
    sys.exit(1)


def ler_configuracao(arquivo_env, exigir_senha=True):
    """Lê remetente, senha e destinatários do arquivo .env.

    Args:
        arquivo_env: caminho do arquivo .env.
        exigir_senha: no modo --dry-run é False, pois nada é enviado.

    Returns:
        Dicionário com "remetente", "senha" e "destinatarios" (lista).
    """
    if not arquivo_env.exists():
        encerrar_com_erro(
            "Arquivo .env não encontrado. Crie-o a partir do modelo com "
            "'cp .env.example .env' e preencha remetente, senha de app e destinatários."
        )

    # Coloca os valores do .env nas variáveis de ambiente. override=True: o .env
    # prevalece sobre variáveis de mesmo nome que já existam no terminal.
    load_dotenv(arquivo_env, override=True)

    obrigatorias = [VAR_REMETENTE, VAR_DESTINATARIOS] + ([VAR_SENHA] if exigir_senha else [])
    faltando = [nome for nome in obrigatorias if not os.getenv(nome, "").strip()]
    if faltando:
        encerrar_com_erro("Preencha no arquivo .env: %s", ", ".join(faltando))

    destinatarios = os.getenv(VAR_DESTINATARIOS).split(",")
    return {
        "remetente": os.getenv(VAR_REMETENTE).strip(),
        # O Google mostra a senha de app em blocos ("abcd efgh ..."); os espaços não fazem parte dela.
        "senha": os.getenv(VAR_SENHA, "").replace(" ", ""),
        "destinatarios": [email.strip() for email in destinatarios if email.strip()],
    }


def verificar_relatorio(caminho, mes=None):
    """Confere se o relatório em Excel existe; se não existir, explica o que fazer e encerra."""
    if not caminho.exists():
        encerrar_com_erro(
            "Relatório não encontrado em %s. Rode primeiro: .venv/bin/python src/relatorio.py%s",
            caminho.relative_to(PASTA_PROJETO), complemento_comando(mes),
        )


# ---------------------------------------------------------------------------
# Conteúdo do e-mail
# ---------------------------------------------------------------------------

def formatar_reais(valor):
    """Formata no padrão brasileiro: 1234.56 -> "R$ 1.234,56"."""
    texto = f"{valor:,.2f}"  # padrão americano: 1,234.56
    # Troca vírgula por ponto e vice-versa, usando "X" como posição temporária.
    return "R$ " + texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_inteiro(numero):
    """Formata com ponto de milhar: 1414 -> "1.414"."""
    return f"{numero:,}".replace(",", ".")


def formatar_variacao(variacao):
    """Formata uma variação com sinal e vírgula decimal: 0.0834 -> "+8,3%"."""
    return f"{variacao * 100:+.1f}%".replace(".", ",")


def linhas_resumo(kpis, comparacao=None):
    """Lista de (indicador, valor já formatado) usada nas versões texto e HTML."""
    linhas = [
        ("Faturamento total", formatar_reais(kpis["faturamento_total"])),
        ("Nº de vendas", formatar_inteiro(kpis["numero_vendas"])),
        ("Ticket médio", formatar_reais(kpis["ticket_medio"])),
        ("Melhor filial", f"{kpis['melhor_filial']} ({formatar_reais(kpis['faturamento_melhor_filial'])})"),
    ]
    if comparacao is not None:
        variacao = comparacao["variacao"]
        texto = SEM_COMPARACAO if variacao is None else formatar_variacao(variacao)
        # Logo abaixo do faturamento total, que é o número comparado.
        linhas.insert(1, (f"Variação vs. {nome_do_mes(comparacao['mes_anterior'])}", texto))
    return linhas


def descrever_relatorio(kpis, mes=None):
    """Trecho da frase de abertura, ex.: "de outubro de 2026" ou "do período de 01/01/2026 a ..."."""
    if mes is None:
        return f"do período de {descrever_periodo(kpis)}"
    return f"de {mes_por_extenso(mes)}"


def montar_assunto(kpis, mes=None):
    """Assunto do e-mail, com o mês (modo mensal) ou o período analisado (modo anual)."""
    if mes is None:
        return f"Relatório de Vendas 2026 — {descrever_periodo(kpis)}"
    return f"Relatório de Vendas — {nome_do_mes(mes)}"


def montar_corpo_texto(kpis, mes=None, comparacao=None):
    """Versão em texto simples do corpo, para programas de e-mail sem HTML."""
    linhas = [f"  - {indicador}: {valor}" for indicador, valor in linhas_resumo(kpis, comparacao)]
    return (
        "Olá,\n\n"
        f"Segue em anexo o relatório de vendas {descrever_relatorio(kpis, mes)}.\n\n"
        "Resumo:\n" + "\n".join(linhas) + "\n\n"
        "Os detalhes por filial e por produto estão nas abas da planilha anexa.\n\n"
        "E-mail gerado automaticamente. Não é necessário responder.\n"
    )


def montar_corpo_html(kpis, mes=None, comparacao=None):
    """Versão em HTML do corpo, com a tabela de KPIs.

    Os estilos ficam dentro de cada tag (style="..."), porque muitos programas
    de e-mail ignoram blocos <style> separados.
    """
    estilo_celula = "padding: 8px 14px; border: 1px solid #d0d7de;"
    # html.escape protege contra caracteres especiais (<, >, &) nos textos.
    linhas_tabela = "\n".join(
        f'<tr><td style="{estilo_celula} font-weight: bold;">{html.escape(indicador)}</td>'
        f'<td style="{estilo_celula}">{html.escape(valor)}</td></tr>'
        for indicador, valor in linhas_resumo(kpis, comparacao)
    )
    titulo = "Relatório de Vendas 2026" if mes is None else f"Relatório de Vendas — {nome_do_mes(mes)}"
    return f"""\
<html>
  <body style="font-family: Arial, Helvetica, sans-serif; color: #333333;">
    <h2 style="color: #{COR_PRINCIPAL};">{titulo}</h2>
    <p>Olá,</p>
    <p>Segue em anexo o relatório de vendas
       <strong>{descrever_relatorio(kpis, mes)}</strong>.</p>
    <table style="border-collapse: collapse;">
      <tr>
        <th style="{estilo_celula} background: #{COR_PRINCIPAL}; color: #ffffff;">Indicador</th>
        <th style="{estilo_celula} background: #{COR_PRINCIPAL}; color: #ffffff;">Valor</th>
      </tr>
{linhas_tabela}
    </table>
    <p>Os detalhes por filial e por produto estão nas abas da planilha anexa.</p>
    <p style="color: #888888; font-size: 12px;">
      E-mail gerado automaticamente. Não é necessário responder.
    </p>
  </body>
</html>
"""


def montar_email(config, kpis, caminho_relatorio, mes=None, comparacao=None):
    """Monta a mensagem completa: cabeçalhos, corpo (texto + HTML) e anexo."""
    mensagem = EmailMessage()
    mensagem["From"] = config["remetente"]
    mensagem["To"] = ", ".join(config["destinatarios"])
    mensagem["Subject"] = montar_assunto(kpis, mes)

    # Primeiro o texto simples; o HTML entra como "alternativa" preferida.
    mensagem.set_content(montar_corpo_texto(kpis, mes, comparacao))
    mensagem.add_alternative(montar_corpo_html(kpis, mes, comparacao), subtype="html")

    tipo_principal, subtipo = TIPO_ANEXO_XLSX
    mensagem.add_attachment(
        caminho_relatorio.read_bytes(),
        maintype=tipo_principal,
        subtype=subtipo,
        filename=caminho_relatorio.name,
    )
    return mensagem


# ---------------------------------------------------------------------------
# Envio ou prévia
# ---------------------------------------------------------------------------

def mostrar_previa(mensagem, caminho_relatorio):
    """Modo --dry-run: mostra no terminal o que seria enviado e salva o HTML para conferência."""
    # A mensagem tem duas versões do corpo; pegamos cada uma pelo tipo.
    corpo_texto = mensagem.get_body(preferencelist=("plain",)).get_content()
    corpo_html = mensagem.get_body(preferencelist=("html",)).get_content()
    ARQUIVO_PREVIA.parent.mkdir(parents=True, exist_ok=True)
    ARQUIVO_PREVIA.write_text(corpo_html, encoding="utf-8")
    tamanho_kb = caminho_relatorio.stat().st_size / 1024

    separador = "=" * 70
    print(separador)
    print("PRÉVIA DO E-MAIL (modo --dry-run: nada foi enviado)")
    print(separador)
    print(f"De:      {mensagem['From']}")
    print(f"Para:    {mensagem['To']}")
    print(f"Assunto: {mensagem['Subject']}")
    print(f"Anexo:   {caminho_relatorio.name} ({formatar_inteiro(round(tamanho_kb))} KB)")
    print("-" * 70)
    print(corpo_texto)
    print(separador)
    print(f"Versão HTML salva em {ARQUIVO_PREVIA.relative_to(PASTA_PROJETO)} "
          "(abra no navegador para ver como fica).")


def enviar_email(mensagem, config):
    """Conecta no Gmail, faz login e envia. Cada falha comum tem uma mensagem própria."""
    try:
        with smtplib.SMTP_SSL(SERVIDOR_SMTP, PORTA_SMTP, timeout=TEMPO_LIMITE_SEGUNDOS) as servidor:
            servidor.login(config["remetente"], config["senha"])
            servidor.send_message(mensagem)
    except smtplib.SMTPAuthenticationError:
        encerrar_com_erro(
            "O Gmail recusou o login de %s. Confira no .env o remetente e a %s: ela deve "
            "ser uma senha de app de 16 letras (criada em "
            "https://myaccount.google.com/apppasswords, com a verificação em duas etapas "
            "ativada), não a senha normal da conta.",
            config["remetente"], VAR_SENHA,
        )
    except smtplib.SMTPRecipientsRefused as erro:
        encerrar_com_erro(
            "Destinatário(s) recusado(s): %s. Confira %s no .env.",
            ", ".join(erro.recipients), VAR_DESTINATARIOS,
        )
    except smtplib.SMTPException as erro:
        encerrar_com_erro("O servidor do Gmail recusou o envio: %s", erro)
    except OSError as erro:  # sem internet, servidor fora do ar, tempo esgotado...
        encerrar_com_erro(
            "Não foi possível conectar a %s:%d (%s). Verifique sua conexão com a internet.",
            SERVIDOR_SMTP, PORTA_SMTP, erro,
        )

    log.info("E-mail enviado para: %s", ", ".join(config["destinatarios"]))


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------

def ler_argumentos():
    """Lê as opções da linha de comando: --mes e --dry-run."""
    parser = argparse.ArgumentParser(description="Envia o relatório de vendas por e-mail (Gmail).")
    adicionar_argumento_mes(parser)
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="monta o e-mail e mostra o que seria enviado, sem enviar",
    )
    return parser.parse_args()


def executar(mes=None, dry_run=False):
    """Verifica a configuração, monta o e-mail e envia (ou só mostra, com dry_run=True).

    Args:
        mes: período mensal (ex.: 2026-10) ou None para o relatório anual.
        dry_run: se True, não envia; só mostra a prévia e salva o HTML.
    """
    config = ler_configuracao(ARQUIVO_ENV, exigir_senha=not dry_run)
    caminho_relatorio = arquivo_relatorio(mes)
    verificar_relatorio(caminho_relatorio, mes)

    kpis = calcular_kpis(carregar_base(mes))
    comparacao = None if mes is None else comparar_com_mes_anterior(kpis["faturamento_total"], mes)
    mensagem = montar_email(config, kpis, caminho_relatorio, mes, comparacao)

    if dry_run:
        mostrar_previa(mensagem, caminho_relatorio)
    else:
        enviar_email(mensagem, config)


def main():
    """Lê as opções da linha de comando e executa o envio."""
    argumentos = ler_argumentos()
    configurar_log()
    executar(argumentos.mes, argumentos.dry_run)


if __name__ == "__main__":
    main()
