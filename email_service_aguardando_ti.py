import json
import os
from pathlib import Path

import msal
import requests


# ============================================================
# CONFIGURAÇÃO
# FASE 2.0B2-A
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

CONFIG_FILE = (
    BASE_DIR
    / "email_config.json"
)

TOKEN_CACHE_FILE = (
    BASE_DIR
    / ".email_token_cache.json"
)

GRAPH_BASE_URL = (
    "https://graph.microsoft.com/v1.0"
)

GRAPH_SENDMAIL_ENDPOINT = (
    f"{GRAPH_BASE_URL}/me/sendMail"
)

GRAPH_ME_ENDPOINT = (
    f"{GRAPH_BASE_URL}/me"
)

DEFAULT_SCOPES = [
    "Mail.Send",
    "User.Read",
]


# ============================================================
# EXCEÇÕES
# ============================================================

class EmailServiceError(
    Exception
):
    """Erro base do serviço de e-mail."""


class EmailConfigurationError(
    EmailServiceError
):
    """Erro de configuração do Microsoft Graph."""


class EmailAuthenticationError(
    EmailServiceError
):
    """Erro de autenticação no Microsoft Entra ID."""


class EmailSendError(
    EmailServiceError
):
    """Erro durante o envio pelo Microsoft Graph."""


# ============================================================
# CONFIGURAÇÃO
# ============================================================

def _texto(
    valor
):

    if valor is None:
        return ""

    return str(
        valor
    ).strip()


def _validar_email(
    email
):

    email = (
        _texto(
            email
        )
        .lower()
    )

    if not email:

        raise ValueError(
            "Informe o e-mail."
        )

    if (
        " " in email
        or "@" not in email
    ):

        raise ValueError(
            f"E-mail inválido: {email}"
        )

    local, dominio = (
        email.rsplit(
            "@",
            1
        )
    )

    if (
        not local
        or not dominio
        or "." not in dominio
        or dominio.startswith(".")
        or dominio.endswith(".")
    ):

        raise ValueError(
            f"E-mail inválido: {email}"
        )

    return email


def _normalizar_destinatarios(
    destinatarios
):

    """
    Aceita:

        [
            "usuario@empresa.com",
            {
                "nome": "Nome",
                "email": "usuario@empresa.com"
            }
        ]

    Retorna lista sem duplicidade.
    """

    if not destinatarios:

        return []

    resultado = []

    emails_vistos = set()

    for item in destinatarios:

        if isinstance(
            item,
            str
        ):

            nome = ""

            email = item

        elif isinstance(
            item,
            dict
        ):

            nome = _texto(
                item.get(
                    "nome"
                )
            )

            email = (
                item.get(
                    "email"
                )
                or item.get(
                    "address"
                )
            )

        else:

            raise ValueError(
                "Destinatário inválido."
            )

        email = (
            _validar_email(
                email
            )
        )

        if email in emails_vistos:

            continue

        emails_vistos.add(
            email
        )

        resultado.append(
            {
                "nome": nome,
                "email": email,
            }
        )

    return resultado


def carregar_configuracao():

    """
    Ordem de leitura:

    1. Variáveis de ambiente
    2. email_config.json

    Variáveis suportadas:

        FISCAL_TRACKER_CLIENT_ID
        FISCAL_TRACKER_TENANT_ID

    Nenhuma senha do Outlook é necessária.
    """

    client_id = _texto(
        os.getenv(
            "FISCAL_TRACKER_CLIENT_ID"
        )
    )

    tenant_id = _texto(
        os.getenv(
            "FISCAL_TRACKER_TENANT_ID"
        )
    )

    config_arquivo = {}

    if CONFIG_FILE.exists():

        try:

            with open(
                CONFIG_FILE,
                "r",
                encoding="utf-8"
            ) as arquivo:

                config_arquivo = (
                    json.load(
                        arquivo
                    )
                )

        except json.JSONDecodeError as erro:

            raise EmailConfigurationError(
                "O arquivo email_config.json "
                "possui JSON inválido."
            ) from erro

        except OSError as erro:

            raise EmailConfigurationError(
                "Não foi possível ler "
                "email_config.json."
            ) from erro

    if not client_id:

        client_id = _texto(
            config_arquivo.get(
                "client_id"
            )
        )

    if not tenant_id:

        tenant_id = _texto(
            config_arquivo.get(
                "tenant_id"
            )
        )

    if not client_id:

        raise EmailConfigurationError(
            "Microsoft Graph ainda não configurado. "
            "Informe FISCAL_TRACKER_CLIENT_ID "
            "ou client_id em email_config.json."
        )

    if not tenant_id:

        raise EmailConfigurationError(
            "Microsoft Graph ainda não configurado. "
            "Informe FISCAL_TRACKER_TENANT_ID "
            "ou tenant_id em email_config.json."
        )

    authority = (
        "https://login.microsoftonline.com/"
        f"{tenant_id}"
    )

    return {
        "client_id": client_id,
        "tenant_id": tenant_id,
        "authority": authority,
        "scopes": DEFAULT_SCOPES.copy(),
    }


def criar_modelo_configuracao():

    """
    Cria um arquivo modelo somente quando
    email_config.json ainda não existe.

    Não sobrescreve configuração existente.
    """

    if CONFIG_FILE.exists():

        return False

    modelo = {
        "client_id": (
            "COLE_AQUI_O_APPLICATION_CLIENT_ID"
        ),
        "tenant_id": (
            "COLE_AQUI_O_DIRECTORY_TENANT_ID"
        ),
    }

    with open(
        CONFIG_FILE,
        "w",
        encoding="utf-8"
    ) as arquivo:

        json.dump(
            modelo,
            arquivo,
            indent=4,
            ensure_ascii=False,
        )

    return True


# ============================================================
# CACHE DE TOKEN
# ============================================================

def _carregar_cache_token():

    cache = (
        msal.SerializableTokenCache()
    )

    if TOKEN_CACHE_FILE.exists():

        try:

            conteudo = (
                TOKEN_CACHE_FILE.read_text(
                    encoding="utf-8"
                )
            )

            if conteudo.strip():

                cache.deserialize(
                    conteudo
                )

        except Exception:

            # Cache corrompido não deve impedir
            # o funcionamento da aplicação.
            pass

    return cache


def _salvar_cache_token(
    cache
):

    if not cache.has_state_changed:

        return

    try:

        TOKEN_CACHE_FILE.write_text(
            cache.serialize(),
            encoding="utf-8",
        )

    except OSError:

        # Falha no cache não invalida
        # um token que já foi obtido.
        pass


# ============================================================
# APLICAÇÃO MSAL
# ============================================================

def _criar_aplicacao_msal():

    config = (
        carregar_configuracao()
    )

    cache = (
        _carregar_cache_token()
    )

    app = (
        msal.PublicClientApplication(
            client_id=(
                config[
                    "client_id"
                ]
            ),
            authority=(
                config[
                    "authority"
                ]
            ),
            token_cache=cache,
        )
    )

    return (
        app,
        cache,
        config,
    )


# ============================================================
# AUTENTICAÇÃO
# ============================================================

def obter_token_silencioso():

    """
    Tenta utilizar uma sessão/token já armazenado.

    Não abre navegador e não solicita interação.
    """

    (
        app,
        cache,
        config,
    ) = _criar_aplicacao_msal()

    contas = (
        app.get_accounts()
    )

    if not contas:

        return None

    resultado = (
        app.acquire_token_silent(
            scopes=(
                config[
                    "scopes"
                ]
            ),
            account=contas[0],
        )
    )

    _salvar_cache_token(
        cache
    )

    if (
        resultado
        and "access_token"
        in resultado
    ):

        return resultado[
            "access_token"
        ]

    return None


def autenticar_interativamente():

    """
    Abre o navegador local para login corporativo.

    Requer que o App Registration permita
    aplicações públicas / desktop e tenha
    redirect URI http://localhost.
    """

    (
        app,
        cache,
        config,
    ) = _criar_aplicacao_msal()

    try:

        resultado = (
            app.acquire_token_interactive(
                scopes=(
                    config[
                        "scopes"
                    ]
                ),
                redirect_uri=(
                    "http://localhost"
                ),
                prompt="select_account",
            )
        )

    except Exception as erro:

        raise EmailAuthenticationError(
            "Não foi possível iniciar "
            "a autenticação no Microsoft 365."
        ) from erro

    _salvar_cache_token(
        cache
    )

    if (
        resultado
        and "access_token"
        in resultado
    ):

        return resultado[
            "access_token"
        ]

    descricao = (
        resultado.get(
            "error_description"
        )
        if isinstance(
            resultado,
            dict
        )
        else None
    )

    raise EmailAuthenticationError(
        descricao
        or (
            "A autenticação no Microsoft 365 "
            "não foi concluída."
        )
    )


def autenticar_device_code():

    """
    Fluxo alternativo.

    Retorna inicialmente a mensagem/código
    e aguarda o usuário concluir o login.

    Útil quando a abertura automática
    do navegador não funciona.
    """

    (
        app,
        cache,
        config,
    ) = _criar_aplicacao_msal()

    flow = (
        app.initiate_device_flow(
            scopes=(
                config[
                    "scopes"
                ]
            )
        )
    )

    if (
        not flow
        or "user_code"
        not in flow
    ):

        raise EmailAuthenticationError(
            "Não foi possível iniciar "
            "o Device Code Flow."
        )

    print()
    print(
        flow.get(
            "message",
            (
                "Conclua a autenticação "
                "no Microsoft 365."
            ),
        )
    )
    print()

    resultado = (
        app.acquire_token_by_device_flow(
            flow
        )
    )

    _salvar_cache_token(
        cache
    )

    if (
        resultado
        and "access_token"
        in resultado
    ):

        return resultado[
            "access_token"
        ]

    descricao = (
        resultado.get(
            "error_description"
        )
        if isinstance(
            resultado,
            dict
        )
        else None
    )

    raise EmailAuthenticationError(
        descricao
        or (
            "A autenticação pelo código "
            "do dispositivo não foi concluída."
        )
    )


def obter_token(
    permitir_interacao=True
):

    """
    Primeiro tenta token silencioso.

    Se não houver sessão válida e
    permitir_interacao=True, abre login.
    """

    token = (
        obter_token_silencioso()
    )

    if token:

        return token

    if not permitir_interacao:

        raise EmailAuthenticationError(
            "Não existe uma sessão autenticada "
            "no Microsoft 365."
        )

    return (
        autenticar_interativamente()
    )


# ============================================================
# PERFIL DO USUÁRIO
# ============================================================

def obter_usuario_logado(
    token=None
):

    if not token:

        token = (
            obter_token()
        )

    headers = {
        "Authorization": (
            f"Bearer {token}"
        ),
        "Accept": (
            "application/json"
        ),
    }

    try:

        resposta = (
            requests.get(
                GRAPH_ME_ENDPOINT,
                headers=headers,
                timeout=30,
            )
        )

    except requests.RequestException as erro:

        raise EmailServiceError(
            "Não foi possível consultar "
            "o usuário conectado ao Microsoft 365."
        ) from erro

    if resposta.status_code != 200:

        raise EmailServiceError(
            _mensagem_erro_graph(
                resposta
            )
        )

    dados = resposta.json()

    return {
        "id": dados.get(
            "id"
        ),
        "nome": (
            dados.get(
                "displayName"
            )
            or ""
        ),
        "email": (
            dados.get(
                "mail"
            )
            or dados.get(
                "userPrincipalName"
            )
            or ""
        ),
        "user_principal_name": (
            dados.get(
                "userPrincipalName"
            )
            or ""
        ),
    }


# ============================================================
# GRAPH HELPERS
# ============================================================

def _mensagem_erro_graph(
    resposta
):

    status = (
        resposta.status_code
    )

    try:

        dados = (
            resposta.json()
        )

        erro_graph = (
            dados.get(
                "error",
                {}
            )
        )

        codigo = (
            erro_graph.get(
                "code"
            )
        )

        mensagem = (
            erro_graph.get(
                "message"
            )
        )

    except Exception:

        codigo = None

        mensagem = None

    if codigo and mensagem:

        return (
            f"Microsoft Graph retornou "
            f"{status} - {codigo}: "
            f"{mensagem}"
        )

    if mensagem:

        return (
            f"Microsoft Graph retornou "
            f"{status}: {mensagem}"
        )

    return (
        f"Microsoft Graph retornou "
        f"HTTP {status}."
    )


def _converter_destinatarios_graph(
    destinatarios
):

    destinatarios = (
        _normalizar_destinatarios(
            destinatarios
        )
    )

    return [
        {
            "emailAddress": {
                "address": (
                    item[
                        "email"
                    ]
                ),
                **(
                    {
                        "name": (
                            item[
                                "nome"
                            ]
                        )
                    }
                    if item[
                        "nome"
                    ]
                    else {}
                ),
            }
        }
        for item
        in destinatarios
    ]


# ============================================================
# ENVIO
# ============================================================

def enviar_email(
    destinatarios,
    assunto,
    mensagem,
    cc=None,
    bcc=None,
    html=False,
    salvar_em_enviados=True,
    token=None,
):

    """
    Envia um e-mail pelo Outlook / Microsoft Graph.

    Retorno:

        {
            "ok": True,
            "status_code": 202,
            "mensagem": "...",
            "destinatarios": [...]
        }

    Em caso de erro, levanta EmailSendError.
    """

    assunto = (
        _texto(
            assunto
        )
    )

    mensagem = (
        _texto(
            mensagem
        )
    )

    if not assunto:

        raise ValueError(
            "Informe o assunto do e-mail."
        )

    if not mensagem:

        raise ValueError(
            "Informe a mensagem do e-mail."
        )

    destinatarios_normalizados = (
        _normalizar_destinatarios(
            destinatarios
        )
    )

    if not destinatarios_normalizados:

        raise ValueError(
            "Informe pelo menos um destinatário."
        )

    if not token:

        token = (
            obter_token()
        )

    mensagem_graph = {
        "subject": assunto,

        "body": {
            "contentType": (
                "HTML"
                if html
                else "Text"
            ),
            "content": mensagem,
        },

        "toRecipients": (
            _converter_destinatarios_graph(
                destinatarios_normalizados
            )
        ),
    }

    if cc:

        cc_graph = (
            _converter_destinatarios_graph(
                cc
            )
        )

        if cc_graph:

            mensagem_graph[
                "ccRecipients"
            ] = cc_graph

    if bcc:

        bcc_graph = (
            _converter_destinatarios_graph(
                bcc
            )
        )

        if bcc_graph:

            mensagem_graph[
                "bccRecipients"
            ] = bcc_graph

    payload = {
        "message": (
            mensagem_graph
        ),
        "saveToSentItems": bool(
            salvar_em_enviados
        ),
    }

    headers = {
        "Authorization": (
            f"Bearer {token}"
        ),
        "Content-Type": (
            "application/json"
        ),
    }

    try:

        resposta = (
            requests.post(
                GRAPH_SENDMAIL_ENDPOINT,
                headers=headers,
                json=payload,
                timeout=30,
            )
        )

    except requests.Timeout as erro:

        raise EmailSendError(
            "O Microsoft Graph demorou "
            "demais para responder."
        ) from erro

    except requests.RequestException as erro:

        raise EmailSendError(
            "Não foi possível conectar "
            "ao Microsoft Graph."
        ) from erro

    if resposta.status_code != 202:

        raise EmailSendError(
            _mensagem_erro_graph(
                resposta
            )
        )

    return {
        "ok": True,

        "status_code": (
            resposta.status_code
        ),

        "mensagem": (
            "E-mail aceito para envio "
            "pelo Microsoft 365."
        ),

        "destinatarios": [
            item[
                "email"
            ]
            for item
            in destinatarios_normalizados
        ],

        "quantidade_destinatarios": (
            len(
                destinatarios_normalizados
            )
        ),
    }


# ============================================================
# TESTE DE CONFIGURAÇÃO
# ============================================================

def testar_configuracao(
    autenticar=False
):

    resultado = {
        "configuracao": False,
        "autenticacao": False,
        "usuario": None,
        "mensagem": "",
    }

    try:

        config = (
            carregar_configuracao()
        )

        resultado[
            "configuracao"
        ] = True

        resultado[
            "mensagem"
        ] = (
            "Configuração do Microsoft Graph encontrada."
        )

    except EmailConfigurationError as erro:

        resultado[
            "mensagem"
        ] = str(
            erro
        )

        return resultado

    if not autenticar:

        return resultado

    try:

        token = (
            obter_token()
        )

        usuario = (
            obter_usuario_logado(
                token
            )
        )

        resultado[
            "autenticacao"
        ] = True

        resultado[
            "usuario"
        ] = usuario

        resultado[
            "mensagem"
        ] = (
            "Autenticação com Microsoft 365 concluída."
        )

    except EmailServiceError as erro:

        resultado[
            "mensagem"
        ] = str(
            erro
        )

    return resultado


# ============================================================
# LIMPEZA DA SESSÃO
# ============================================================

def limpar_cache_autenticacao():

    if TOKEN_CACHE_FILE.exists():

        try:

            TOKEN_CACHE_FILE.unlink()

        except OSError as erro:

            raise EmailServiceError(
                "Não foi possível limpar "
                "o cache de autenticação."
            ) from erro

    return True


# ============================================================
# EXECUÇÃO LOCAL
# ============================================================

if __name__ == "__main__":

    print()
    print(
        "========================================"
    )
    print(
        " FISCAL TRACKER"
    )
    print(
        " MICROSOFT 365 EMAIL SERVICE"
    )
    print(
        " FASE 2.0B2-A"
    )
    print(
        "========================================"
    )
    print()

    if not CONFIG_FILE.exists():

        criado = (
            criar_modelo_configuracao()
        )

        if criado:

            print(
                "Arquivo email_config.json "
                "criado."
            )

            print()
            print(
                "Preencha client_id e tenant_id "
                "antes de testar a autenticação."
            )

            print()

    resultado = (
        testar_configuracao(
            autenticar=False
        )
    )

    print(
        "Configuracao:",
        (
            "OK"
            if resultado[
                "configuracao"
            ]
            else "PENDENTE"
        ),
    )

    print(
        resultado[
            "mensagem"
        ]
    )

    print()
    print(
        "Nenhum e-mail foi enviado."
    )
    print()

