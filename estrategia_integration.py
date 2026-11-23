# -*- coding: utf-8 -*-
"""
Modulo de Integracao com o Estratégia OS para o sistema Enriquecimento Play.

Frentes cobertas:
A. Micro-API de dados (leitura) - rotas seguras com X-API-Key, rate limiting,
   despacho estrito de metricas e catalogo OpenAPI para o Cockpit IA.
B. Single Sign-On (SSO) - endpoint /api/auth/sso com validacao criptografica RS256,
   cache da chave publica do Estrategia OS e mapeamento de contas de usuario.
C. Seguranca de Iframe - cabecalhos Content-Security-Policy: frame-ancestors
   permitindo renderizacao embutida no Estrategia OS.
"""

from __future__ import annotations

import hmac
import json
import os
import re
import ssl

import time
import urllib.request
from collections import defaultdict
from functools import wraps
from typing import Any, Callable, Dict, List, Optional, Tuple

import jwt
from flask import Blueprint, current_app, jsonify, request, session


integracao_bp = Blueprint("integracao_estrategia", __name__)

# Configuracoes de ambiente
DEFAULT_ESTRATEGIA_URL = "https://estrategiaos.bookplay.com.br"
LEGACY_ESTRATEGIA_URL = "https://estrategia-os.vercel.app"  # plano de volta durante a transicao
ESTRATEGIA_OS_URL = os.environ.get("ESTRATEGIA_OS_URL", DEFAULT_ESTRATEGIA_URL).strip().rstrip("/")
# Origens que podem embutir este sistema em iframe e enviar AUTH_SESSION (lista separada por virgula).
RAW_ALLOWED_ORIGINS = os.environ.get(
    "ESTRATEGIA_OS_ALLOWED_ORIGINS",
    f"{DEFAULT_ESTRATEGIA_URL},{LEGACY_ESTRATEGIA_URL},http://localhost:47821"
)
ALLOWED_ORIGINS = set(
    origin.strip().rstrip("/")
    for origin in RAW_ALLOWED_ORIGINS.split(",")
    if origin.strip()
)
ALLOWED_ORIGINS.add(ESTRATEGIA_OS_URL)
ALLOWED_ORIGINS.add("http://localhost:47821")

# Cache em memoria da chave publica RSA do Estrategia OS
_PUBLIC_KEY_CACHE: Optional[str] = None
_PUBLIC_KEY_EXPIRES_AT: float = 0.0
_PUBLIC_KEY_CACHE_TTL = 6 * 3600  # 6 horas


# Rate Limiting em memoria (janela deslizante por chave/IP)
_RATE_LIMIT_BUCKETS: Dict[str, List[float]] = defaultdict(list)
_RATE_LIMIT_WINDOW_SECONDS = 10.0
_RATE_LIMIT_MAX_REQUESTS = 30


def get_integration_api_key() -> str:
    """Recupera a chave de integracao configurada no servidor."""
    return os.environ.get("INTEGRATION_API_KEY", "").strip()


def check_rate_limit(identifier: str) -> bool:
    """Verifica se o identificador excedeu o limite de requisicoes."""
    now = time.time()
    cutoff = now - _RATE_LIMIT_WINDOW_SECONDS
    bucket = _RATE_LIMIT_BUCKETS[identifier]
    # Limpa timestamps antigos
    _RATE_LIMIT_BUCKETS[identifier] = [ts for ts in bucket if ts > cutoff]
    if len(_RATE_LIMIT_BUCKETS[identifier]) >= _RATE_LIMIT_MAX_REQUESTS:
        return False
    _RATE_LIMIT_BUCKETS[identifier].append(now)
    return True


def require_integration_api_key(f: Callable) -> Callable:
    """
    Decorator de autenticacao fail-closed para a micro-API de dados.
    Verifica o cabecalho X-API-Key com comparacao de tempo constante.
    """
    @wraps(f)
    def decorated_function(*args: Any, **kwargs: Any) -> Any:
        expected_key = get_integration_api_key()
        # Modo fail-closed: se a chave nao estiver configurada, recusa imediatamente
        if not expected_key:
            return jsonify({
                "ok": False,
                "error": "Chave de integracao nao configurada no servidor."
            }), 401

        received_key = (
            request.headers.get("X-API-Key")
            or request.headers.get("x-api-key")
            or request.args.get("api_key")
            or ""
        ).strip()

        if not received_key or not hmac.compare_digest(received_key, expected_key):
            return jsonify({
                "ok": False,
                "error": "Chave de integracao ausente ou invalida."
            }), 401

        # Controle de taxa de requisicoes
        client_id = f"{request.remote_addr}:{received_key[:12]}"
        if not check_rate_limit(client_id):
            return jsonify({
                "ok": False,
                "error": "Limite de requisicoes excedido (max 30 req / 10s). Aguarde para tentar novamente."
            }), 429

        return f(*args, **kwargs)

    return decorated_function


def fetch_estrategia_public_key() -> Optional[str]:
    """
    Busca a chave publica RS256 do Estrategia OS com cache em memoria.
    Prioriza variavel local, depois endpoint PEM, com fallback para endpoint JSON.
    """
    global _PUBLIC_KEY_CACHE, _PUBLIC_KEY_EXPIRES_AT

    now = time.time()
    if _PUBLIC_KEY_CACHE and now < _PUBLIC_KEY_EXPIRES_AT:
        return _PUBLIC_KEY_CACHE

    # Se fornecida via variavel de ambiente diretamente
    env_key = os.environ.get("ESTRATEGIA_OS_PUBLIC_KEY_PEM", "").strip()
    if env_key:
        _PUBLIC_KEY_CACHE = env_key
        _PUBLIC_KEY_EXPIRES_AT = now + _PUBLIC_KEY_CACHE_TTL
        return _PUBLIC_KEY_CACHE

    pem_url = os.environ.get("OS_PUBLIC_KEY_URL", "").strip() or f"{ESTRATEGIA_OS_URL}/api/auth/public-key.pem"
    json_url = f"{ESTRATEGIA_OS_URL}/api/auth/public-key.json"

    context = ssl._create_unverified_context()

    # Tentativa 1: endpoint PEM plano
    try:
        req = urllib.request.Request(pem_url, headers={"User-Agent": "EnriquecimentoPlay-Integration/1.0"})
        with urllib.request.urlopen(req, timeout=4.0, context=context) as resp:
            content = resp.read().decode("utf-8").strip()
            if "BEGIN PUBLIC KEY" in content:
                _PUBLIC_KEY_CACHE = content
                _PUBLIC_KEY_EXPIRES_AT = now + _PUBLIC_KEY_CACHE_TTL
                return _PUBLIC_KEY_CACHE
    except Exception:
        pass

    # Tentativa 2: endpoint JSON
    try:
        req = urllib.request.Request(json_url, headers={"User-Agent": "EnriquecimentoPlay-Integration/1.0"})
        with urllib.request.urlopen(req, timeout=4.0, context=context) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            public_key = data.get("publicKey") or data.get("public_key")
            if public_key and "BEGIN PUBLIC KEY" in public_key:
                _PUBLIC_KEY_CACHE = public_key.strip()
                _PUBLIC_KEY_EXPIRES_AT = now + _PUBLIC_KEY_CACHE_TTL
                return _PUBLIC_KEY_CACHE
    except Exception:
        pass

    return _PUBLIC_KEY_CACHE


def validate_sso_token(token: str) -> Tuple[bool, Dict[str, Any], str]:
    """
    Valida token JWT assinado com RS256 pelo Estrategia OS.
    Retorna (sucesso, claims, mensagem_erro).
    """
    if not token:
        return False, {}, "Token ausente."

    public_key = fetch_estrategia_public_key()
    if not public_key:
        return False, {}, "Chave publica do Estrategia OS indisponivel no momento."

    try:
        claims = jwt.decode(
            token,
            public_key,
            algorithms=["RS256"],
            issuer="estrategia-os",
            options={
                "verify_signature": True,
                "verify_exp": True,
                "verify_iss": True,
            }
        )
        return True, claims, ""
    except jwt.ExpiredSignatureError:
        return False, {}, "Token SSO expirado."
    except jwt.InvalidIssuerError:
        return False, {}, "Emissor do token invalido."
    except jwt.InvalidSignatureError:
        return False, {}, "Assinatura digital do token invalida."
    except Exception as e:
        return False, {}, f"Falha na verificacao do token SSO: {str(e)}"


# ==============================================================================
# FRENTE A: MICRO-API DE DADOS (LEITURA)
# ==============================================================================

def get_usuarios_metric() -> Dict[str, Any]:
    """
    Retorna a lista de contas ativas para a Central de Contas do Estrategia OS.
    Exibe apenas username e nome (sem expor senhas, hashes ou dados sensiveis).
    """
    users_dict = current_app.config.get("INTEGRATION_USERS", {})
    usuarios = []
    for username, info in users_dict.items():
        usuarios.append({
            "username": username,
            "nome": info.get("name") or username.capitalize(),
        })
    return {"usuarios": usuarios}


def get_metricas_gerais_metric() -> Dict[str, Any]:
    """Retorna metricas consolidadas de enriquecimento e filas."""
    state = current_app.config.get("INTEGRATION_STATE", {})
    records = state.get("records") or []
    queue = state.get("queue") or []

    total_lotes = len(records)
    total_itens_fila = len(queue)

    pendentes = sum(1 for q in queue if q.get("status") == "pendente")
    processando = sum(1 for q in queue if q.get("status") == "processando")
    concluidos = sum(1 for q in queue if q.get("status") == "concluido")
    recusados = sum(1 for q in queue if q.get("status") == "recusado")

    total_linhas_processadas = 0
    total_linhas_sucesso = 0
    total_linhas_erros = 0

    for q in queue:
        try:
            total_linhas_processadas += int(q.get("processed_count") or 0)
            total_linhas_sucesso += int(q.get("success_count") or 0)
            total_linhas_erros += int(q.get("error_count") or 0)
        except (ValueError, TypeError):
            pass

    taxa_sucesso = 0.0
    if total_linhas_processadas > 0:
        taxa_sucesso = round((total_linhas_sucesso / total_linhas_processadas) * 100, 1)

    return {
        "sistema": "Enriquecimento Play",
        "total_lotes_historico": total_lotes,
        "total_itens_fila": total_itens_fila,
        "fila_pendentes": pendentes,
        "fila_em_processamento": processando,
        "fila_concluidos": concluidos,
        "fila_recusados": recusados,
        "total_linhas_processadas": total_linhas_processadas,
        "total_linhas_sucesso": total_linhas_sucesso,
        "total_linhas_erros": total_linhas_erros,
        "taxa_sucesso_percentual": taxa_sucesso,
    }


def get_fila_metric(params: Dict[str, Any]) -> Dict[str, Any]:
    """Retorna a relacao de itens na fila com suporte a filtros basicos."""
    state = current_app.config.get("INTEGRATION_STATE", {})
    queue = state.get("queue") or []

    status_filter = (params.get("status") or "").strip().lower()
    limit_str = params.get("limit") or "20"
    try:
        limit = min(max(int(limit_str), 1), 100)
    except ValueError:
        limit = 20

    items_filtrados = []
    for q in reversed(queue):
        if status_filter and q.get("status", "").lower() != status_filter:
            continue
        items_filtrados.append({
            "id": q.get("id"),
            "queue_number": q.get("queue_number"),
            "requester_name": q.get("requester_name"),
            "filename": q.get("filename"),
            "status": q.get("status"),
            "total_rows": q.get("total_rows"),
            "processed_count": q.get("processed_count"),
            "success_count": q.get("success_count"),
            "error_count": q.get("error_count"),
            "request_time": q.get("request_time"),
            "completed_at": q.get("completed_at"),
            "mode": q.get("mode"),
        })
        if len(items_filtrados) >= limit:
            break

    return {
        "total_encontrados": len(items_filtrados),
        "filtro_status": status_filter or "todos",
        "itens": items_filtrados,
    }


def get_lotes_recentes_metric(params: Dict[str, Any]) -> Dict[str, Any]:
    """Retorna os lotes de arquivos importados e enriquecidos recentemente."""
    state = current_app.config.get("INTEGRATION_STATE", {})
    records = state.get("records") or []

    limit_str = params.get("limit") or "15"
    try:
        limit = min(max(int(limit_str), 1), 50)
    except ValueError:
        limit = 15

    lotes = []
    for r in reversed(records[-limit:]):
        lotes.append({
            "id": r.get("id"),
            "filename": r.get("filename"),
            "type": r.get("type"),
            "total_in": r.get("total_in"),
            "total_enriched": r.get("total_enriched"),
            "invalid_phone_count": r.get("invalid_phone_count"),
            "columns": r.get("columns", [])[:8],
        })

    return {
        "total_lotes": len(lotes),
        "lotes": lotes,
    }


def get_consumo_api_metric() -> Dict[str, Any]:
    """Retorna o resumo de chamadas operacionais registradas no sistema."""
    audit_func = current_app.config.get("INTEGRATION_AUDIT_FUNC")
    events = audit_func(limit=200) if callable(audit_func) else []

    contagem_acoes: Dict[str, int] = defaultdict(int)
    logins_sucesso = 0
    enriquecimentos = 0
    exportacoes = 0

    for ev in events:
        action = ev.get("action", "")
        contagem_acoes[action] += 1
        if "LOGIN_SUCCESS" in action:
            logins_sucesso += 1
        elif "ENRICH" in action or "PROCESS" in action:
            enriquecimentos += 1
        elif "EXPORT" in action or "DOWNLOAD" in action:
            exportacoes += 1

    return {
        "total_eventos_auditados": len(events),
        "logins_recentes": logins_sucesso,
        "operacoes_enriquecimento": enriquecimentos,
        "exportacoes_geradas": exportacoes,
        "acoes_distribuicao": dict(contagem_acoes),
    }


@integracao_bp.route("/api/integracao/consultar", methods=["GET", "POST"])
@require_integration_api_key
def consultar():
    """
    Rota unica de leitura para o Estratégia OS.
    Parametro 'metrica' despacha para as consultas permitidas.
    """
    if request.method == "POST":
        data = request.get_json(silent=True) or {}
    else:
        data = request.args.to_dict()

    metrica = (data.get("metrica") or "").strip().lower()
    if not metrica:
        return jsonify({
            "ok": False,
            "error": "Parametro 'metrica' obrigatorio. Opcoes: usuarios, metricas_gerais, fila, lotes_recentes, consumo_api."
        }), 400

    try:
        if metrica == "usuarios":
            return jsonify({"ok": True, "data": get_usuarios_metric()}), 200
        elif metrica in ("metricas_gerais", "resumo", "metricas"):
            return jsonify({"ok": True, "data": get_metricas_gerais_metric()}), 200
        elif metrica in ("fila", "status_fila"):
            return jsonify({"ok": True, "data": get_fila_metric(data)}), 200
        elif metrica in ("lotes_recentes", "lotes", "historico_lotes"):
            return jsonify({"ok": True, "data": get_lotes_recentes_metric(data)}), 200
        elif metrica in ("consumo_api", "consultas_api", "auditoria"):
            return jsonify({"ok": True, "data": get_consumo_api_metric()}), 200
        else:
            return jsonify({
                "ok": False,
                "error": f"Metrica desconhecida: '{metrica}'. Opcoes validas: usuarios, metricas_gerais, fila, lotes_recentes, consumo_api."
            }), 422
    except Exception as e:
        return jsonify({
            "ok": False,
            "error": f"Erro interno ao consultar metrica '{metrica}': {str(e)}"
        }), 500


@integracao_bp.route("/api/integracao/openapi.json", methods=["GET"])
@require_integration_api_key
def openapi_schema():
    """
    Schema OpenAPI / Tool catalog para o Cockpit IA do Estratégia OS.
    Descreve a ferramenta 'consultar_enriquecimento' e seus parametros.
    """
    schema = {
        "openapi": "3.0.0",
        "info": {
            "title": "API de Integracao - Enriquecimento Play",
            "version": "1.0.0",
            "description": "Micro-API de leitura de metricas e dados de enriquecimento para o Estrategia OS e Cockpit IA."
        },
        "servers": [
            {"url": request.host_url.rstrip("/")}
        ],
        "paths": {
            "/api/integracao/consultar": {
                "post": {
                    "summary": "Consulta metricas do sistema de enriquecimento",
                    "description": "Retorna informacoes consolidadas sobre usuarios, filas de processamento, lotes e taxas de sucesso.",
                    "operationId": "consultar_enriquecimento",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {
                                    "type": "object",
                                    "required": ["metrica"],
                                    "properties": {
                                        "metrica": {
                                            "type": "string",
                                            "enum": ["usuarios", "metricas_gerais", "fila", "lotes_recentes", "consumo_api"],
                                            "description": "Chave da metrica desejada. 'usuarios' lista contas ativas; 'metricas_gerais' traz totais consolidados e taxas; 'fila' lista jobs em andamento ou pendentes; 'lotes_recentes' exibe ultimos arquivos; 'consumo_api' exibe estatisticas de uso."
                                        },
                                        "status": {
                                            "type": "string",
                                            "enum": ["pendente", "processando", "concluido", "recusado"],
                                            "description": "Filtro opcional para a metrica 'fila'."
                                        },
                                        "limit": {
                                            "type": "integer",
                                            "description": "Quantidade maxima de itens retornados nas metricas de lista (padrao 20, max 100)."
                                        }
                                    }
                                }
                            }
                        }
                    },
                    "responses": {
                        "200": {
                            "description": "Consulta executada com sucesso",
                            "content": {
                                "application/json": {
                                    "schema": {
                                        "type": "object",
                                        "properties": {
                                            "ok": {"type": "boolean", "example": True},
                                            "data": {"type": "object"}
                                        }
                                    }
                                }
                            }
                        },
                        "400": {"description": "Parametro invalido ou ausente"},
                        "401": {"description": "Chave X-API-Key ausente ou incorreta"},
                        "429": {"description": "Limite de requisicoes excedido"}
                    }
                }
            }
        },
        "tools": [
            {
                "name": "consultar_enriquecimento",
                "description": "Consulta metricas em tempo real do sistema de enriquecimento de dados (Enriquecimento Play), incluindo fila de processamento, lotes concluidos, taxas de enriquecimento e usuarios.",
                "parameters": {
                    "type": "object",
                    "required": ["metrica"],
                    "properties": {
                        "metrica": {
                            "type": "string",
                            "enum": ["usuarios", "metricas_gerais", "fila", "lotes_recentes", "consumo_api"],
                            "description": "Identificador da metrica a ser consultada."
                        },
                        "status": {
                            "type": "string",
                            "description": "Filtro opcional de status para a consulta de fila (ex: pendente, processando, concluido)."
                        },
                        "limit": {
                            "type": "integer",
                            "description": "Limite de registros a retornar."
                        }
                    }
                }
            }
        ]
    }
    return jsonify(schema), 200


# ==============================================================================
# FRENTE B: SINGLE SIGN-ON (SSO)
# ==============================================================================

@integracao_bp.route("/api/auth/sso", methods=["POST"])
def auth_sso():
    """
    Endpoint de Single Sign-On consumido apos postMessage AUTH_SESSION no iframe.
    Valida token RS256, localiza conta local e inicializa sessao web.
    """
    data = request.get_json(silent=True) or {}
    token = (data.get("token") or "").strip()

    # Suporte a Bearer token no cabecalho Authorization
    if not token and request.headers.get("Authorization"):
        auth_header = request.headers.get("Authorization", "")
        if auth_header.startswith("Bearer "):
            token = auth_header.split(" ", 1)[1].strip()

    if not token:
        return jsonify({"ok": False, "error": "Token de autenticacao ausente."}), 400

    valido, claims, erro = validate_sso_token(token)
    if not valido:
        return jsonify({"ok": False, "error": erro}), 401

    sub = str(claims.get("sub") or "").strip().lower()
    if not sub:
        return jsonify({"ok": False, "error": "Token sem identificador de usuario (claim 'sub')."}), 400

    users_dict = current_app.config.get("INTEGRATION_USERS", {})
    user = users_dict.get(sub)

    # Caso nao encontre pelo username exato, verifica se o claim 'role' e master
    # e mapeia para a conta admin se configurado
    if not user and claims.get("role") == "master" and "admin" in users_dict:
        sub = "admin"
        user = users_dict.get("admin")

    if not user:
        return jsonify({
            "ok": False,
            "error": f"Usuario '{sub}' autenticado pelo Estrategia OS nao esta vinculado neste sistema."
        }), 401

    # Inicializa sessao Flask oficial da aplicacao
    session["user_id"] = sub
    session["role"] = user.get("role", "user1")
    session["user_name"] = user.get("name") or sub.capitalize()

    audit_func = current_app.config.get("INTEGRATION_AUDIT_FUNC")
    if callable(audit_func):
        try:
            audit_func(
                "SSO_LOGIN_SUCCESS",
                user.get("name") or sub,
                {"username": sub, "role": user.get("role"), "origin": "estrategia-os"}
            )
        except Exception:
            pass

    return jsonify({
        "ok": True,
        "user": {
            "username": sub,
            "name": user.get("name"),
            "role": user.get("role"),
        },
        "redirect": "/",
    }), 200


# ==============================================================================
# MIDDLEWARE DE CABECALHOS HTTP (CSP / IFRAME)
# ==============================================================================

def apply_iframe_security_headers(response: Any) -> Any:
    """
    Injeta cabecalhos de seguranca permitindo incorporacao em iframes do Estrategia OS.
    Remove X-Frame-Options restritivo, aplica Content-Security-Policy com frame-ancestors
    e formata cookies Set-Cookie com SameSite=None, Secure e Partitioned (CHIPS).
    """
    # Remove X-Frame-Options que bloquearia a exibicao dentro do iframe do Estrategia OS
    response.headers.pop("X-Frame-Options", None)

    ancestors = " ".join(sorted(ALLOWED_ORIGINS))
    csp_directive = f"frame-ancestors 'self' {ancestors};"

    existing_csp = response.headers.get("Content-Security-Policy")
    if existing_csp:
        if "frame-ancestors" not in existing_csp:
            response.headers["Content-Security-Policy"] = f"{existing_csp}; {csp_directive}"
    else:
        response.headers["Content-Security-Policy"] = csp_directive

    # Garante que cookies de sessao sejam aceitos em contexto cross-origin de iframes (Chrome CHIPS / Partitioned)
    set_cookies = response.headers.getlist("Set-Cookie")
    if set_cookies:
        response.headers.remove("Set-Cookie")
        for sc in set_cookies:
            # Garante Secure
            if not re.search(r";\s*Secure", sc, flags=re.IGNORECASE):
                sc += "; Secure"
            # Garante SameSite=None
            if re.search(r"SameSite=[^;]+", sc, flags=re.IGNORECASE):
                sc = re.sub(r"SameSite=[^;]+", "SameSite=None", sc, flags=re.IGNORECASE)
            else:
                sc += "; SameSite=None"
            # Garante Partitioned (CHIPS para iframes cross-origin em navegadores modernos)
            if not re.search(r";\s*Partitioned", sc, flags=re.IGNORECASE):
                sc += "; Partitioned"
            response.headers.add("Set-Cookie", sc)

    return response


from flask.sessions import SecureCookieSessionInterface

class PartitionedSessionInterface(SecureCookieSessionInterface):
    """
    SessionInterface customizada para garantir cookies particionados (CHIPS),
    SameSite=None e Secure para operacao transparente dentro de iframes cross-origin.
    """
    def save_session(self, app: Any, session: Any, response: Any) -> None:
        super().save_session(app, session, response)
        cookies = response.headers.getlist("Set-Cookie")
        if cookies:
            response.headers.remove("Set-Cookie")
            cookie_name = getattr(app, "session_cookie_name", "session")
            for c in cookies:
                if f"{cookie_name}=" in c or "session=" in c:
                    if "; Partitioned" not in c and "; partitioned" not in c:
                        c += "; Partitioned"
                    if "; Secure" not in c and "; secure" not in c:
                        c += "; Secure"
                    if "SameSite=" in c:
                        c = re.sub(r"SameSite=[^;]+", "SameSite=None", c, flags=re.IGNORECASE)
                    else:
                        c += "; SameSite=None"
                response.headers.add("Set-Cookie", c)

