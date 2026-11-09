"""Testes da leitura da configuração do e-mail (src/email_sender.py).

Garantem que a prévia (--dry-run) funciona sem .env, como no GitHub Actions,
e que o envio real continua exigindo todas as configurações.
"""

import pytest

from email_sender import (
    DESTINATARIO_EXEMPLO,
    REMETENTE_EXEMPLO,
    VAR_DESTINATARIOS,
    VAR_REMETENTE,
    VAR_SENHA,
    ler_configuracao,
)


@pytest.fixture(autouse=True)
def ambiente_limpo(monkeypatch):
    """Apaga as variáveis do e-mail antes de cada teste e as restaura depois.

    Assim os testes não dependem do .env de quem roda nem de outro teste.
    """
    for nome in (VAR_REMETENTE, VAR_SENHA, VAR_DESTINATARIOS):
        monkeypatch.delenv(nome, raising=False)


def criar_env(pasta, conteudo):
    """Cria um arquivo .env temporário com o conteúdo dado e devolve o caminho."""
    arquivo = pasta / ".env"
    arquivo.write_text(conteudo, encoding="utf-8")
    return arquivo


def test_previa_sem_env_usa_enderecos_de_exemplo(tmp_path):
    config = ler_configuracao(tmp_path / ".env", exigir_senha=False)
    assert config["remetente"] == REMETENTE_EXEMPLO
    assert config["destinatarios"] == [DESTINATARIO_EXEMPLO]


def test_previa_com_env_preenchido_usa_os_valores_dele(tmp_path):
    arquivo = criar_env(
        tmp_path,
        f"{VAR_REMETENTE}=eu@empresa.com\n{VAR_DESTINATARIOS}=a@empresa.com, b@empresa.com\n",
    )
    config = ler_configuracao(arquivo, exigir_senha=False)
    assert config["remetente"] == "eu@empresa.com"
    assert config["destinatarios"] == ["a@empresa.com", "b@empresa.com"]


def test_previa_com_env_vazio_completa_so_o_que_falta(tmp_path):
    arquivo = criar_env(tmp_path, f"{VAR_REMETENTE}=eu@empresa.com\n{VAR_DESTINATARIOS}=\n")
    config = ler_configuracao(arquivo, exigir_senha=False)
    assert config["remetente"] == "eu@empresa.com"
    assert config["destinatarios"] == [DESTINATARIO_EXEMPLO]


def test_envio_real_sem_env_encerra(tmp_path):
    with pytest.raises(SystemExit):
        ler_configuracao(tmp_path / ".env", exigir_senha=True)


def test_envio_real_sem_senha_encerra(tmp_path):
    arquivo = criar_env(
        tmp_path, f"{VAR_REMETENTE}=eu@empresa.com\n{VAR_DESTINATARIOS}=a@empresa.com\n"
    )
    with pytest.raises(SystemExit):
        ler_configuracao(arquivo, exigir_senha=True)
