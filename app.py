from __future__ import annotations

import sys
from pathlib import Path

# Garantir carregamento das bibliotecas do ambiente virtual local
_venv_site = Path(__file__).resolve().parent / ".venv" / "Lib" / "site-packages"
if _venv_site.exists() and str(_venv_site) not in sys.path:
    sys.path.insert(0, str(_venv_site))

import asyncio
import base64
import csv
import html
import io
import json
import math
import os
import queue
import random
import re
import ssl
import shutil
import threading
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import uuid
import xml.etree.ElementTree as ET
from datetime import datetime, date, time as dt_time
from functools import wraps
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
from flask import Flask, jsonify, redirect, render_template, request, Response, send_file, send_from_directory, session, url_for, has_request_context
from openpyxl import Workbook, load_workbook
try:
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
except ImportError:
    ILLEGAL_CHARACTERS_RE = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")
from werkzeug.security import check_password_hash, generate_password_hash
import zipfile


def load_env_file(env_path: Path) -> None:
    """Carrega variáveis do arquivo .env de forma nativa e resiliente."""
    if not env_path.exists():
        return
    try:
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k, v = k.strip(), v.strip()
                if (v.startswith('"') and v.endswith('"')) or (v.startswith("'") and v.endswith("'")):
                    v = v[1:-1]
                if k and k not in os.environ:
                    os.environ[k] = v
    except Exception as e:
        pass

BASE_DIR = Path(__file__).resolve().parent
try:
    from dotenv import load_dotenv
    load_dotenv(dotenv_path=BASE_DIR / ".env")
except ImportError:
    load_env_file(BASE_DIR / ".env")

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or __import__("secrets").token_hex(32)
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0
app.config["MAX_CONTENT_LENGTH"] = 1024 * 1024 * 1024  # 1 GB
app.config["MAX_FORM_MEMORY_SIZE"] = 1024 * 1024 * 1024  # 1 GB

# Configuracao de sessao para operacao embutida em iframe (cross-origin e CHIPS)
app.config["SESSION_COOKIE_SAMESITE"] = os.environ.get("SESSION_COOKIE_SAMESITE", "None")
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("SESSION_COOKIE_SECURE", "1") == "1"




@app.context_processor
def override_url_for():
    return dict(url_for=dated_url_for)

def dated_url_for(endpoint, **values):
    if endpoint == 'static':
        filename = values.get('filename', None)
        if filename:
            file_path = os.path.join(app.root_path, 'static', filename)
            if os.path.exists(file_path):
                values['v'] = int(os.stat(file_path).st_mtime)
    return url_for(endpoint, **values)

def _required_secret(name: str) -> str:
    """Le um segredo do ambiente/.env. Sem valor, gera um aleatorio (valido so nesta execucao)."""
    value = os.environ.get(name, "").strip()
    if value:
        return value
    import secrets as _secrets
    value = _secrets.token_urlsafe(18)
    print(f"[AVISO] {name} nao definido no .env; usando valor aleatorio temporario: {value}", flush=True)
    return value


ADMIN_PASSWORD = _required_secret("ADMIN_PASSWORD")
SAF_PASSWORD = _required_secret("SAF_PASSWORD")
ESTRATEGIA_PASSWORD = _required_secret("ESTRATEGIA_PASSWORD")

USERS = {
    "admin": {
        "password_hash": generate_password_hash(ADMIN_PASSWORD),
        "role": "admin",
        "name": "Estratégia - Admin",
    },
    "saf": {
        "password_hash": generate_password_hash(SAF_PASSWORD),
        "role": "user1",
        "name": "Usuário SAF",
    },
    "estrategia": {
        "password_hash": generate_password_hash(ESTRATEGIA_PASSWORD),
        "role": "user1",
        "name": "Estratégia",
    },
}

SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://yaresjmqcrpuiorpvbck.supabase.co/rest/v1/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "sb_publishable_k7ISy3EG7LO6aR7Htsx7_g_1DIWf7OP")

import tempfile

is_vercel = os.environ.get("VERCEL") == "1" or os.environ.get("VERCEL_ENV") is not None
TEMP_PARENT = Path(tempfile.gettempdir()) if is_vercel else BASE_DIR

UPLOAD_DIR = TEMP_PARENT / "uploads"
EXPORT_DIR = TEMP_PARENT / "exports"
LOG_PATH = TEMP_PARENT / "api_debug.log"
AUDIT_LOG_FILE = TEMP_PARENT / "audit_logs.jsonl"
UPLOAD_DIR.mkdir(exist_ok=True)
EXPORT_DIR.mkdir(exist_ok=True)

def get_client_ip() -> str:
    """Extrai o IP real do cliente, considerando proxies reversos e cabecalhos de encaminhamento."""
    if not request:
        return "127.0.0.1"
    cf_ip = request.headers.get("CF-Connecting-IP")
    if cf_ip:
        return cf_ip.strip()
    xff = request.headers.get("X-Forwarded-For")
    if xff:
        client_ip = xff.split(",")[0].strip()
        if client_ip:
            return client_ip
    x_real_ip = request.headers.get("X-Real-IP")
    if x_real_ip:
        return x_real_ip.strip()
    return request.remote_addr or "127.0.0.1"


def db_save_audit_event(action: str, user_name: str = "", details: Any = None) -> Dict[str, Any]:
    """Salva evento no registro de auditoria com protecao de contexto."""
    user_id_val = "system"
    role_val = "user1"
    resolved_name = user_name or "Sistema"
    client_ip_val = ""

    if has_request_context():
        try:
            resolved_name = user_name or session.get("user_name") or "Sistema"
            user_id_val = session.get("user_id") or "system"
            role_val = session.get("role") or "user1"
            client_ip_val = get_client_ip()
        except Exception:
            pass

    event = {
        "id": str(uuid.uuid4()),
        "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "user_name": resolved_name,
        "user_id": user_id_val,
        "role": role_val,
        "action": action,
        "details": details or {},
        "ip": client_ip_val,
    }
    try:
        with open(AUDIT_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps(event, default=json_serial_fallback, ensure_ascii=False) + "\n")
    except Exception as e:
        log_api_debug(f"Erro ao salvar audit log: {e}")
    return event

def db_get_audit_events(limit: int = 150) -> List[Dict[str, Any]]:
    """Recupera os últimos eventos de auditoria ordenados do mais recente para o mais antigo."""
    events = []
    if AUDIT_LOG_FILE.exists():
        try:
            with open(AUDIT_LOG_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line:
                        try:
                            events.append(json.loads(line))
                        except Exception:
                            pass
        except Exception as e:
            log_api_debug(f"Erro ao ler audit log: {e}")
    events.reverse()
    return events[:limit]

@app.route('/favicon.ico')
@app.route('/favicon.png')
def favicon():
    return send_from_directory(os.path.join(app.root_path, 'static'), 'favicon.png', mimetype='image/png')

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("user_id"):
            if request.path.startswith("/api/"):
                return jsonify({"error": "Não autenticado."}), 401
            return redirect(url_for("login_page"))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("user_id"):
            if request.path.startswith("/api/"):
                return jsonify({"error": "Não autenticado."}), 401
            return redirect(url_for("login_page"))
        if session.get("role") != "admin":
            return jsonify({"error": "Acesso restrito para estratégia."}), 403
        return f(*args, **kwargs)
    return decorated_function

STATE: Dict[str, List[Dict[str, Any]]] = {
    "records": [],
    "queue": [],
    "notifications": [],
}

DATA_DIR = TEMP_PARENT / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
RECORDS_FILE = DATA_DIR / "records.json"
QUEUE_FILE = DATA_DIR / "queue.json"
NOTIFICATIONS_FILE = DATA_DIR / "notifications.json"

_STORAGE_LOCK = threading.Lock()

def _load_local_storage():
    """Carrega dados salvos localmente em disco ao iniciar."""
    with _STORAGE_LOCK:
        if RECORDS_FILE.exists():
            try:
                with open(RECORDS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        for rec in data:
                            if isinstance(rec, dict):
                                rec["rows"] = []
                        STATE["records"] = data
            except Exception as e:
                log_api_debug(f"Erro ao carregar records.json: {e}")


        if QUEUE_FILE.exists():
            try:
                with open(QUEUE_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        STATE["queue"] = data
            except Exception as e:
                log_api_debug(f"Erro ao carregar queue.json: {e}")

        if NOTIFICATIONS_FILE.exists():
            try:
                with open(NOTIFICATIONS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        STATE["notifications"] = data
            except Exception as e:
                log_api_debug(f"Erro ao carregar notifications.json: {e}")

_load_local_storage()

# Integracao com a plataforma Estrategia OS (Micro-API e SSO)
try:
    from estrategia_integration import integracao_bp, apply_iframe_security_headers, ALLOWED_ORIGINS, PartitionedSessionInterface
    app.session_interface = PartitionedSessionInterface()
    app.config["INTEGRATION_USERS"] = USERS
    app.config["INTEGRATION_STATE"] = STATE
    app.config["INTEGRATION_AUDIT_FUNC"] = db_get_audit_events
    app.register_blueprint(integracao_bp)

    import gzip

    @app.after_request
    def handle_post_response(response):
        response = apply_iframe_security_headers(response)
        try:
            if (
                response.status_code < 300
                and not getattr(response, "direct_passthrough", False)
                and response.content_length
                and response.content_length > 1024
                and request
                and "gzip" in request.headers.get("Accept-Encoding", "").lower()
                and not response.headers.get("Content-Encoding")
                and any(t in (response.mimetype or "") for t in ("json", "text", "javascript", "css", "xml"))
            ):
                compressed_data = gzip.compress(response.get_data(), compresslevel=6)
                response.set_data(compressed_data)
                response.headers["Content-Encoding"] = "gzip"
                response.headers["Content-Length"] = len(compressed_data)
        except Exception:
            pass
        return response
except Exception as _e_integ:
    log_api_debug(f"Alerta ao inicializar integracao Estrategia OS: {_e_integ}")



def _persist_local_state(key: str) -> None:
    """Grava o estado local em arquivo de forma segura e não-bloqueante."""
    def _save():
        with _STORAGE_LOCK:
            target_file = None
            if key == "records":
                target_file = RECORDS_FILE
            elif key == "queue":
                target_file = QUEUE_FILE
            elif key == "notifications":
                target_file = NOTIFICATIONS_FILE

            if target_file and key in STATE:
                try:
                    tmp_file = target_file.with_suffix(".tmp")
                    with open(tmp_file, "w", encoding="utf-8") as f:
                        json.dump(STATE[key], f, ensure_ascii=False, default=json_serial_fallback)
                    tmp_file.replace(target_file)
                except Exception as e:
                    log_api_debug(f"Erro ao persistir {target_file}: {e}")

    threading.Thread(target=_save, daemon=True).start()

def json_serial_fallback(obj: Any) -> str:
    if isinstance(obj, (datetime, date)):
        return obj.strftime("%d/%m/%Y")
    if isinstance(obj, dt_time):
        return obj.strftime("%H:%M:%S")
    return str(obj)

# Circuit Breaker para chamadas ao Supabase
_SUPABASE_OFFLINE_UNTIL = 0.0
_SUPABASE_CONSECUTIVE_FAILURES = 0

def make_supabase_request(table: str, method: str = "GET", data: Any = None, query: str = "") -> Any:
    global _SUPABASE_OFFLINE_UNTIL, _SUPABASE_CONSECUTIVE_FAILURES

    now = time.time()
    # Se o Supabase estiver marcado como offline, retorna None imediatamente (0ms)
    if now < _SUPABASE_OFFLINE_UNTIL:
        return None

    if not SUPABASE_URL or not SUPABASE_KEY:
        return None

    url = f"{SUPABASE_URL}{table}{query}"
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=representation"
    }
    
    encoded_data = None
    if data is not None:
        encoded_data = json.dumps(data, default=json_serial_fallback).encode("utf-8")
        
    req = urllib.request.Request(
        url,
        data=encoded_data,
        headers=headers,
        method=method
    )
    
    try:
        context = ssl._create_unverified_context()
        # Timeout reduzido para 1.5s para nunca travar a aplicação
        with urllib.request.urlopen(req, timeout=1.5, context=context) as resp:
            body = resp.read().decode("utf-8")
            _SUPABASE_CONSECUTIVE_FAILURES = 0
            if body:
                return json.loads(body)
            return []
    except Exception as e:
        _SUPABASE_CONSECUTIVE_FAILURES += 1
        if _SUPABASE_CONSECUTIVE_FAILURES >= 2:
            # Desativa tentativas ao Supabase por 10 minutos
            _SUPABASE_OFFLINE_UNTIL = time.time() + 600
            log_api_debug(f"Supabase inacessivel ({e}). Modo offline ativado por 10 minutos.")
        else:
            log_api_debug(f"Supabase request failed ({method} {url}): {e}")
        return None

def db_save_record(record: Dict[str, Any]) -> None:
    # Manter em memória local e persistir no records.json (sem bloqueio de rede)
    STATE["records"].append(record)
    if len(STATE["records"]) > 500:
        STATE["records"] = STATE["records"][-500:]
    _persist_local_state("records")

    def _sync_supa():
        record_db = {
            k: v for k, v in record.items()
            if k in {"id", "filename", "stored_file", "type", "available_fields", "preview", "total_in", "total_enriched", "columns", "invalid_phone_count"}
        }
        make_supabase_request("records", "POST", record_db)

    threading.Thread(target=_sync_supa, daemon=True).start()

def db_get_record(record_id: str) -> Dict[str, Any] | None:
    if not record_id:
        return None
    found = next((r for r in STATE["records"] if r.get("id") == record_id), None)
    if found:
        return found
    if RECORDS_FILE.exists():
        try:
            with open(RECORDS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                found = next((r for r in saved if r.get("id") == record_id), None)
                if found:
                    STATE["records"].append(found)
                    return found
        except Exception:
            pass
    return None


def extract_reject_reason(item: Dict[str, Any]) -> str:
    if item.get("reject_reason"):
        return str(item["reject_reason"]).strip()
    logs = item.get("logs") or []
    for log in logs:
        if isinstance(log, str) and "Solicitação recusada pela Estratégia. Motivo: " in log:
            return log.split("Solicitação recusada pela Estratégia. Motivo: ")[1].strip()
    return ""

def db_save_queue_item(item: Dict[str, Any]) -> None:
    if not any(q.get("id") == item["id"] for q in STATE["queue"]):
        STATE["queue"].append(item)
    else:
        for idx, q in enumerate(STATE["queue"]):
            if q.get("id") == item["id"]:
                STATE["queue"][idx] = item
                break
    _persist_local_state("queue")

    def _sync_q():
        item_db = {
            k: v for k, v in item.items()
            if k in {"id", "record_id", "queue_number", "requester_name", "observacoes", "filename", "status", "total_rows", "processed_count", "success_count", "error_count", "request_time", "completed_at", "summary_name", "selected_fields", "detected_type", "logs", "reject_reason", "mode"}
        }
        make_supabase_request("queue", "POST", item_db)

    threading.Thread(target=_sync_q, daemon=True).start()

def db_update_queue_item(item_id: str, updates: Dict[str, Any]) -> None:
    item = next((job for job in STATE["queue"] if job.get("id") == item_id), None)
    if item:
        item.update(updates)
    _persist_local_state("queue")

    updates_db = {
        k: v for k, v in updates.items()
        if k in {"id", "record_id", "queue_number", "requester_name", "observacoes", "filename", "status", "total_rows", "processed_count", "success_count", "error_count", "request_time", "completed_at", "summary_name", "selected_fields", "detected_type", "logs", "reject_reason", "mode"}
    }
    def _sync_patch():
        make_supabase_request("queue", "PATCH", data=updates_db, query=f"?id=eq.{item_id}")
    threading.Thread(target=_sync_patch, daemon=True).start()

def db_get_queue_item(item_id: str) -> Dict[str, Any] | None:
    item = next((job for job in STATE["queue"] if job.get("id") == item_id), None)
    if not item:
        res = make_supabase_request("queue", "GET", query=f"?id=eq.{item_id}")
        if res and len(res) > 0:
            item = res[0]
            if not any(q.get("id") == item.get("id") for q in STATE["queue"]):
                STATE["queue"].append(item)
                _persist_local_state("queue")
    if item:
        item["reject_reason"] = extract_reject_reason(item)
        if "queue_number" not in item:
            item["queue_number"] = ""
        item["reject_reason"] = extract_reject_reason(item)
    return item

def db_get_all_queue() -> List[Dict[str, Any]]:
    # Retorna o estado local instantaneamente sem bloquear a thread
    items = STATE["queue"]
    for item in items:
        if isinstance(item, dict):
            if "queue_number" not in item:
                item["queue_number"] = ""
            item["reject_reason"] = extract_reject_reason(item)

    # Sincronizacao não bloqueante em thread separada se o Supabase estiver disponivel
    if time.time() >= _SUPABASE_OFFLINE_UNTIL:
        def _bg_sync():
            try:
                res = make_supabase_request("queue", "GET")
                if res is not None and isinstance(res, list):
                    db_ids = {i.get("id") for i in res if isinstance(i, dict) and "id" in i}
                    for local_item in STATE["queue"]:
                        if isinstance(local_item, dict) and local_item.get("id") not in db_ids:
                            res.append(local_item)
                    STATE["queue"] = res
                    _persist_local_state("queue")
            except Exception:
                pass
        threading.Thread(target=_bg_sync, daemon=True).start()

    return items

def db_save_notification(notification: Dict[str, Any]) -> None:
    STATE["notifications"].append(notification)
    _persist_local_state("notifications")
    make_supabase_request("notifications", "POST", notification)

def db_get_notifications() -> List[Dict[str, Any]]:
    items = STATE["notifications"]
    if time.time() >= _SUPABASE_OFFLINE_UNTIL:
        res = make_supabase_request("notifications", "GET")
        if res is not None and isinstance(res, list):
            STATE["notifications"] = res
            _persist_local_state("notifications")
            return res
    return items

REQUIRED_FIELDS = ["nome", "telefone"]
EXTRA_FIELDS = [
    "cpf",
    "data_nascimento",
    "idade",
    "sexo",
    "email",
    "logradouro",
    "numero",
    "complemento",
    "bairro",
    "cidade",
    "estado",
    "cep",
    "renda_estimada",
    "cbo",
    "profissao",
    "score",
    "perfil",
    "obito",
    "whatsapp",
    "status",
    "similaridade_nome",
    "telefone_2",
    "telefone_3",
    "telefone_4",
    "telefone_5",
]
ALL_ENRICH_FIELDS = REQUIRED_FIELDS + EXTRA_FIELDS
ENRICHED_PREFIX = "enriquecido_"
PREFIXED_ENRICH_FIELDS = [f"{ENRICHED_PREFIX}{k}" for k in ALL_ENRICH_FIELDS]


def load_local_config() -> Dict[str, str]:
    config_file = Path(__file__).resolve().parent / "config.json"
    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    env_file = Path(__file__).resolve().parent / ".env"
    if env_file.exists():
        config = {}
        try:
            with open(env_file, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        config[k.strip()] = v.strip().strip('"').strip("'")
            return config
        except Exception:
            pass

    return {}

LOCAL_CONFIG = load_local_config()

NOVA_VIDA_BASE_URL = (LOCAL_CONFIG.get("NOVA_VIDA_BASE_URL") or os.getenv("NOVA_VIDA_BASE_URL") or "https://wsnv.novavidati.com.br").strip().rstrip("/")
NOVA_VIDA_TOKEN_URL = (LOCAL_CONFIG.get("NOVA_VIDA_TOKEN_URL") or os.getenv("NOVA_VIDA_TOKEN_URL") or f"{NOVA_VIDA_BASE_URL}/WSLocalizador.asmx/GerarToken").strip()
NOVA_VIDA_API_URL = (LOCAL_CONFIG.get("NOVA_VIDA_API_URL") or os.getenv("NOVA_VIDA_API_URL") or NOVA_VIDA_BASE_URL).strip().rstrip("/")
NOVA_VIDA_API_KEY = (LOCAL_CONFIG.get("NOVA_VIDA_API_KEY") or os.getenv("NOVA_VIDA_API_KEY") or os.getenv("NOVA_VIDA_TOKEN") or "").strip()
NOVA_VIDA_API_ENDPOINT = (LOCAL_CONFIG.get("NOVA_VIDA_API_ENDPOINT") or os.getenv("NOVA_VIDA_API_ENDPOINT") or "WSLocalizador.asmx/PesquisaAtributosTK").strip().lstrip("/")
NOVA_VIDA_USER = (LOCAL_CONFIG.get("NOVA_VIDA_USER") or os.getenv("NOVA_VIDA_USER") or "").strip()
NOVA_VIDA_PASSWORD = (LOCAL_CONFIG.get("NOVA_VIDA_PASSWORD") or os.getenv("NOVA_VIDA_PASSWORD") or "").strip()
NOVA_VIDA_CLIENT = (LOCAL_CONFIG.get("NOVA_VIDA_CLIENT") or os.getenv("NOVA_VIDA_CLIENT") or "").strip()
NOVA_VIDA_API_HEADER = (LOCAL_CONFIG.get("NOVA_VIDA_API_HEADER") or os.getenv("NOVA_VIDA_API_HEADER") or "Authorization").strip()
NOVA_VIDA_API_METHOD = (LOCAL_CONFIG.get("NOVA_VIDA_API_METHOD") or os.getenv("NOVA_VIDA_API_METHOD") or "POST").strip().upper()


_MOJIBAKE_TRIGGERS = ('Ã', 'Â', 'ï¿½', '\ufffd', '\xef\xbf\xbd')


def repair_mojibake_text(text: Any) -> str:
    if text is None:
        return ""
    if not isinstance(text, str):
        text = str(text)
    if not text:
        return ""

    # Fast path: 99.9% dos textos limpos pulam direto sem executar dezenas de regexes pesadas
    if not any(m in text for m in _MOJIBAKE_TRIGGERS):
        return text

    # 1. Tratar double UTF-8 / Mojibake comum (ex: Ã§ -> ç, Ã£ -> ã, etc)
    mojibake_markers = ["Ã§", "Ã£", "Ã©", "Ã¡", "Ã³", "Ãº", "Ãª", "Ã´", "Ã¢", "Ã ", "Â°", "Âª", "Âº", "Ã‡", "Ãƒ", "Ã‰", "Ã\x81", "Ã“", "Ãš", "ÃŠ", "Ã”", "Ã‚", "Ã€"]
    if any(m in text for m in mojibake_markers):
        try:
            repaired = text.encode("latin1").decode("utf-8")
            text = repaired
        except Exception:
            pass

    if not any(m in text for m in ("ï¿½", "\ufffd", "\xef\xbf\xbd")):
        return text

    bad = r"(?:[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]|ï¿½|\?{1,2})+"

    # 2. Casos de palavras terminadas em ção / ções
    # Ex: Vincula[bad]o -> Vinculação, Situa[bad]o -> Situação
    text = re.sub(r"([a-zA-Z\u00c0-\u00ff]+)" + bad + r"([oO]\b)", r"\g<1>çã\g<2>", text)
    text = re.sub(r"([a-zA-Z\u00c0-\u00ff]+)" + bad + r"([eE][sS]\b)", r"\g<1>çõ\g<2>", text)

    # Casos de ...ss[bad]o -> ...ssão (Profissão, Emissão, etc)
    text = re.sub(r"([a-zA-Z\u00c0-\u00ff]+[sS]{2})çã([oO]\b)", r"\g<1>ã\g<2>", text)
    text = re.sub(r"([a-zA-Z\u00c0-\u00ff]+(?:is|us|rs|ls))çã([oO]\b)", r"\g<1>ã\g<2>", text)
    text = re.sub(r"([a-zA-Z\u00c0-\u00ff]+[tT])çã([oO]\b)", r"\g<1>ã\g<2>", text)
    text = re.sub(r"([a-zA-Z\u00c0-\u00ff]+[rR])çã([oO]\b)", r"\g<1>ã\g<2>", text)
    text = re.sub(r"(\b[nN])çã([oO]\b)", r"\g<1>ã\g<2>", text)
    text = re.sub(r"(\b[sS])çã([oO]\b)", r"\g<1>ã\g<2>", text)

    # 3. Palavras com 'ç' no meio (Endereço, Serviço, Praça, Braço, etc)
    text = re.sub(r"endere" + bad + r"o\b", "endereço", text, flags=re.IGNORECASE)
    text = re.sub(r"servi" + bad + r"o\b", "serviço", text, flags=re.IGNORECASE)
    text = re.sub(r"pra" + bad + r"a\b", "praça", text, flags=re.IGNORECASE)
    text = re.sub(r"bra" + bad + r"o\b", "braço", text, flags=re.IGNORECASE)

    # 4. Palavras acentuadas pontuais muito comuns
    text = re.sub(r"\bn" + bad + r"mero\b", "número", text, flags=re.IGNORECASE)
    text = re.sub(r"\bm" + bad + r"s\b", "mês", text, flags=re.IGNORECASE)
    text = re.sub(r"\bsal" + bad + r"rio\b", "salário", text, flags=re.IGNORECASE)
    text = re.sub(r"\bhor" + bad + r"rio\b", "horário", text, flags=re.IGNORECASE)
    text = re.sub(r"\brelat" + bad + r"rio\b", "relatório", text, flags=re.IGNORECASE)
    text = re.sub(r"\busu" + bad + r"rio\b", "usuário", text, flags=re.IGNORECASE)
    text = re.sub(r"\bbenef" + bad + r"cio\b", "benefício", text, flags=re.IGNORECASE)
    text = re.sub(r"\bmunic" + bad + r"pio\b", "município", text, flags=re.IGNORECASE)
    text = re.sub(r"\bcr" + bad + r"dito\b", "crédito", text, flags=re.IGNORECASE)
    text = re.sub(r"\bd" + bad + r"bito\b", "débito", text, flags=re.IGNORECASE)
    text = re.sub(r"\bc" + bad + r"digo\b", "código", text, flags=re.IGNORECASE)
    text = re.sub(r"\bper" + bad + r"odo\b", "período", text, flags=re.IGNORECASE)
    text = re.sub(r"\bm" + bad + r"dico\b", "médico", text, flags=re.IGNORECASE)
    text = re.sub(r"\bsa" + bad + r"de\b", "saúde", text, flags=re.IGNORECASE)
    text = re.sub(r"\bv" + bad + r"nculo\b", "vínculo", text, flags=re.IGNORECASE)

    # 5. Se a palavra original estava toda em maiúsculas, manter em maiúsculas
    def fix_casing(match):
        w = match.group(0)
        uppers = sum(1 for c in w if c.isupper())
        if uppers > len(w) / 2:
            return w.upper()
        if w[0].isupper():
            return w[0].upper() + w[1:].lower()
        return w

    text = re.sub(r"\b[a-zA-Z\u00c0-\u00ff]+\b", fix_casing, text)

    # Limpeza final de qualquer resíduo inválido
    text = re.sub(r"[\u00ef\u00bf\u00bd\ufffd\xef\xbf\xbd]+", "", text)
    return text.strip()


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = repair_mojibake_text(str(value).strip())
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower().replace(" ", "_")
    text = re.sub(r"[^a-z0-9_]+", "_", text)
    return text.strip("_")


def labelize(value: Any, strip_prefix: bool = False) -> str:
    safe = (value or "").strip()
    if not safe:
        return "Campo"
    if safe.startswith("enriquecido_"):
        suffix = safe.replace("enriquecido_", "").strip()
        norm_s = normalize_text(suffix)
        base_label = ""
        if norm_s == "cpf":
            base_label = "CPF"
        elif norm_s == "cbo":
            base_label = "CBO"
        elif norm_s in ("uf", "estado"):
            base_label = "Estado"
        elif norm_s == "profissao":
            base_label = "Profissão"
        elif norm_s == "obito":
            base_label = "Óbito"
        elif norm_s in ("situacao", "status"):
            base_label = "Situação"
        elif norm_s in ("endereco", "logradouro"):
            base_label = "Endereço"
        elif norm_s in ("municipio", "cidade"):
            base_label = "Cidade"
        elif norm_s.startswith("telefone_"):
            num = norm_s.replace("telefone_", "")
            base_label = f"Telefone {num}"
        else:
            clean_s = suffix.replace("_", " ")
            base_label = clean_s.title()

        if strip_prefix:
            return base_label
        return f"Enriquecido - {base_label}"

    norm = normalize_text(safe)
    if norm == "cpf":
        return "CPF"
    if norm == "cbo":
        return "CBO"
    if norm in ("uf", "estado"):
        return "Estado"
    if norm == "profissao":
        return "Profissão" if safe.lower() in ("profissao", "profissão") else safe
    if norm == "obito":
        return "Óbito" if safe.lower() in ("obito", "óbito") else safe
    if norm.startswith("telefone_"):
        num = norm.replace("telefone_", "")
        return f"Telefone {num}"
    clean = safe.replace("_", " ")
    return clean.title() if safe.islower() or safe.isupper() else safe


def is_birthdate_field(field_name: Any) -> bool:
    norm = normalize_field_name(field_name)
    if not norm:
        return False
    return any(p in norm for p in ["nascimento", "nasc", "birth", "dt_nasc", "data_nasc", "data_de_nasc", "aniversario", "dt_nascimento"])


def format_to_brazilian_date(val: Any) -> str:
    if val is None or val == "":
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%d/%m/%Y")

    val_str = str(val).strip()
    if not val_str:
        return ""

    # DD/MM/YYYY ou DD/MM/YYYY HH:MM:SS
    match_dmy = re.match(r"^(\d{2})[-/](\d{2})[-/](\d{4})(?:\s+\d{2}:\d{2}(?::\d{2})?)?$", val_str)
    if match_dmy:
        d, m, y = match_dmy.groups()
        return f"{d}/{m}/{y}"

    # YYYY-MM-DD ou YYYY-MM-DD HH:MM:SS ou YYYY-MM-DDTHH:MM:SS
    match_ymd = re.match(r"^(\d{4})[-/](\d{2})[-/](\d{2})(?:[T\s]\d{2}:\d{2}(?::\d{2})?)?$", val_str)
    if match_ymd:
        y, m, d = match_ymd.groups()
        return f"{d}/{m}/{y}"

    # YYYYMMDD
    match_compact = re.match(r"^(\d{4})(\d{2})(\d{2})$", val_str)
    if match_compact:
        y, m, d = match_compact.groups()
        if int(m) <= 12 and int(d) <= 31:
            return f"{d}/{m}/{y}"

    # DDMMYYYY
    match_compact_dmy = re.match(r"^(\d{2})(\d{2})(\d{4})$", val_str)
    if match_compact_dmy:
        d, m, y = match_compact_dmy.groups()
        if int(m) <= 12 and int(d) <= 31:
            return f"{d}/{m}/{y}"

    return val_str


def normalize_email_value(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def normalize_renda_estimada(value: Any) -> str:
    if value is None or value == "":
        return ""
    text = str(value).strip()
    if not text:
        return ""

    if re.search(r"\d[\d\.,]*\s*(?:AT|A T)\s*\d[\d\.,]*", text, flags=re.IGNORECASE):
        return re.sub(r"\s+", " ", text).replace("  ", " ")

    if text.upper().startswith("R$") or text.upper().startswith("RS"):
        text = text[2:] if text.upper().startswith("R$") else text[2:]
    cleaned = re.sub(r"[^\d,\.\-\s]", "", text)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if not cleaned:
        return text
    return cleaned


def validate_cpf_checksum(cpf: Any) -> bool:
    """Valida os dígitos verificadores de um CPF de 11 dígitos."""
    if not cpf:
        return False
    digits = re.sub(r"\D", "", str(cpf))
    if len(digits) != 11 or not digits.isdigit():
        return False
    if len(set(digits)) == 1:
        return False

    s = sum(int(digits[i]) * (10 - i) for i in range(9))
    d1 = 11 - (s % 11)
    d1 = 0 if d1 >= 10 else d1
    if int(digits[9]) != d1:
        return False

    s = sum(int(digits[i]) * (11 - i) for i in range(10))
    d2 = 11 - (s % 11)
    d2 = 0 if d2 >= 10 else d2
    return int(digits[10]) == d2


def normalize_cpf_value(value: Any) -> str:
    """Normaliza e formata CPF para 11 dígitos com zfill para evitar perda de zeros à esquerda."""
    if value is None or value == "":
        return ""
    val_str = str(value).strip()
    digits = re.sub(r"\D", "", val_str)
    if not digits:
        return val_str
    if 9 <= len(digits) <= 11:
        return digits.zfill(11)
    return digits


VALID_BRAZILIAN_DDDS = {
    "11", "12", "13", "14", "15", "16", "17", "18", "19",
    "21", "22", "24", "27", "28",
    "31", "32", "33", "34", "35", "37", "38",
    "41", "42", "43", "44", "45", "46", "47", "48", "49",
    "51", "53", "54", "55",
    "61", "62", "63", "64", "65", "66", "67", "68", "69",
    "71", "73", "74", "75", "77", "79",
    "81", "82", "83", "84", "85", "86", "87", "88", "89",
    "91", "92", "93", "94", "95", "96", "97", "98", "99"
}


def normalize_phone_value(value: Any, ddd: str = "") -> str:
    """
    Normaliza qualquer telefone brasileiro (celular ou fixo), tratando com precisão:
    - Prefixos de tipo de contato (ex: 'p:', 'm:', 'c:', 't:', 'w:', 'tel:', 'cel:').
    - O DDI do Brasil (+55, 0055, +), distinguindo-o com precisão do DDD 55 (Santa Maria / RS).
    - Textos explicativos (ex: 'Tel:', 'WhatsApp:', 'Cel:').
    - Inserção do 9º dígito em celulares de 8 dígitos.
    - Associação de DDD fornecido quando o número for local (8 ou 9 dígitos).
    """
    if value is None or value == "":
        return ""
    val_str = str(value).strip()
    if not val_str:
        return ""

    # Remove prefixos de contato como 'p:', 'p:+55', 'tel:', 'cel:', 'm:', 'c:', 'w:', 'whats:', etc.
    val_str = re.sub(r"^\s*(?:(?:whatsapp|whats|wpp|telefone|telefones|celulares|celular|contatos|contato|mobile|phone|fone|fixo|fixos|ramal|voice|work|home|cell|mob|tel|cel|zap|fax|p|m|c|t|w|f|v)\s*[:.-]?\s*)+", "", val_str, flags=re.IGNORECASE).strip()
    if not val_str:
        return ""

    clean_ddd = re.sub(r"\D", "", str(ddd or ""))
    if clean_ddd.startswith("0") and len(clean_ddd) == 3:
        clean_ddd = clean_ddd[1:]
    clean_ddd = clean_ddd[:2]

    # 1. Se a string começa explicitamente com DDI internacional (+55, + 55, 0055)
    if re.search(r"^\s*(?:\+|00)\s*55\b", val_str):
        remainder = re.sub(r"^\s*(?:\+|00)\s*55[\s.-]*", "", val_str)
        rem_digits = re.sub(r"\D", "", remainder)
        if rem_digits.startswith("0") and len(rem_digits) in (11, 12):
            rem_digits = rem_digits[1:]

        if len(rem_digits) == 11 and rem_digits[2] == "9":
            return rem_digits
        elif len(rem_digits) == 10:
            if rem_digits[:2] in VALID_BRAZILIAN_DDDS and rem_digits[2] in "6789":
                return f"{rem_digits[:2]}9{rem_digits[2:]}"
            elif len(clean_ddd) == 2 and rem_digits[0] in "6789":
                return f"{clean_ddd}9{rem_digits}"
            return rem_digits
        elif len(rem_digits) == 9:
            if len(clean_ddd) == 2:
                return f"{clean_ddd}{rem_digits}"
            return f"55{rem_digits}"
        elif len(rem_digits) == 8:
            if rem_digits[0] in "6789":
                rem_digits = f"9{rem_digits}"
            if len(clean_ddd) == 2:
                return f"{clean_ddd}{rem_digits}"
            return f"55{rem_digits}"

    digits = re.sub(r"\D", "", val_str)
    if not digits:
        return val_str

    if digits.startswith("0055"):
        digits = digits[4:]
    elif digits.startswith("00"):
        digits = digits[2:]

    # 2. 13 ou 14 dígitos iniciando com DDI 55 (ex: 5555999998888 ou 5582996506901)
    if digits.startswith("55") and len(digits) in (13, 14):
        digits = digits[2:]
        if digits.startswith("0") and len(digits) == 12:
            digits = digits[1:]
        if len(digits) == 11 and digits[2] == "9":
            return digits

    # 3. 12 dígitos (ex: 55 + DDD + 8 dígitos móvel OU 0 + DDD + 9 dígitos móvel)
    if len(digits) == 12:
        if digits.startswith("0"):
            d11 = digits[1:]
            if len(d11) == 11 and d11[2] == "9":
                return d11
            elif len(d11) == 10 and d11[2] in "6789":
                return f"{d11[:2]}9{d11[2:]}"
        elif digits.startswith("55") and digits[2:4] in VALID_BRAZILIAN_DDDS:
            extracted_ddd = digits[2:4]
            local8 = digits[4:]
            if local8[0] in "6789":
                return f"{extracted_ddd}9{local8}"
            else:
                return f"{extracted_ddd}{local8}"
        elif digits[:2] in VALID_BRAZILIAN_DDDS and digits[2] == "0":
            return f"{digits[:2]}{digits[3:]}"

    # 4. 11 dígitos
    if len(digits) == 11:
        if digits.startswith("0"):
            d10 = digits[1:]
            if d10[:2] in VALID_BRAZILIAN_DDDS and d10[2] in "6789":
                return f"{d10[:2]}9{d10[2:]}"
            return d10
        
        # Número com DDD brasileiro válido (inclusive DDD 55!)
        if digits[:2] in VALID_BRAZILIAN_DDDS and digits[2] == "9":
            return digits
        
        if digits.startswith("55") and digits[2:4] in VALID_BRAZILIAN_DDDS:
            extracted_ddd = digits[2:4]
            local7 = digits[4:]
            return f"{extracted_ddd}9{local7}" if len(local7) == 7 else digits

        return digits

    # 5. 10 dígitos (DDD + 8 dígitos)
    if len(digits) == 10:
        if digits.startswith("0"):
            local9 = digits[1:]
            if len(clean_ddd) == 2:
                return f"{clean_ddd}{local9}"
            return local9
        
        extracted_ddd = digits[:2]
        local8 = digits[2:]
        if extracted_ddd in VALID_BRAZILIAN_DDDS:
            if local8[0] in "6789":
                return f"{extracted_ddd}9{local8}"
            else:
                return f"{extracted_ddd}{local8}"
        elif len(clean_ddd) == 2:
            if digits[0] == "9":
                return f"{clean_ddd}{digits[1:]}"

    # 6. 9 dígitos (móvel local)
    if len(digits) == 9:
        if len(clean_ddd) == 2:
            return f"{clean_ddd}{digits}"
        return digits

    # 7. 8 dígitos (local 8 dígitos)
    if len(digits) == 8:
        if digits[0] in "6789":
            digits = f"9{digits}"
        if len(clean_ddd) == 2:
            digits = f"{clean_ddd}{digits}"
        return digits

    # Fallback via regex estruturado
    m_std = re.search(
        r"\(?(?:0?([1-9][0-9]))\)?[\s.-]?(?:(9)[\s.-]?)?([0-9]{4})[\s.-]?([0-9]{4})",
        val_str
    )
    if m_std:
        extracted_ddd = m_std.group(1)
        has_nine = m_std.group(2)
        p1 = m_std.group(3)
        p2 = m_std.group(4)
        local_num = f"9{p1}{p2}" if has_nine or p1[0] in "6789" else f"{p1}{p2}"
        return f"{extracted_ddd}{local_num}"

    return digits


PHONE_ANNOTATION_WORDS = {
    "p", "m", "c", "t", "w", "f", "v",
    "tel", "telefone", "telefones", "cel", "celular", "celulares",
    "whats", "whatsapp", "zap", "fone", "fones", "contato", "contatos",
    "phone", "mobile", "cell", "work", "home", "voice", "fax",
    "fixo", "fixos", "ramal", "comercial", "residencial", "principal",
    "casa", "trabalho", "recado", "recados", "pessoal", "movel", "moveis",
    "wpp", "tel1", "tel2", "cel1", "cel2"
}


def is_annotation_only_text(text: str) -> bool:
    if not text:
        return True
    s = str(text).strip().lower()
    s = re.sub(r'^\s*(?:\+|00)?\s*55\b', '', s)
    s = re.sub(r'[\(\)\s:.\-\/\+\\]+', '', s).strip()
    if not s or s.isdigit():
        return True
    return s in PHONE_ANNOTATION_WORDS


def extract_phone_and_name_from_combined_text(val: Any) -> Tuple[Optional[str], Optional[str]]:
    """
    Detecta se um texto contém simultaneamente um Telefone (10 ou 11 dígitos) e um Nome,
    suportando prefixos como 'p:+55', '+55', '+', 'Tel:', 'Cel:', 'WhatsApp:', 'Zap:', 'Fone:', 'Contato:', etc.
    Exemplos:
      '82996506901 AMANDA KESIA CUNHA ROCHA' -> ('82996506901', 'AMANDA KESIA CUNHA ROCHA')
      '(82) 99650-6901 AMANDA KESIA CUNHA ROCHA' -> ('82996506901', 'AMANDA KESIA CUNHA ROCHA')
      '+55 82 99650-6901 AMANDA KESIA CUNHA ROCHA' -> ('82996506901', 'AMANDA KESIA CUNHA ROCHA')
      'p:+55 82 99650-6901 AMANDA KESIA' -> ('82996506901', 'AMANDA KESIA')
      'Tel: +55 (82) 99650-6901 AMANDA KESIA' -> ('82996506901', 'AMANDA KESIA')
      'AMANDA KESIA CUNHA ROCHA (Tel: +55 82 99650-6901)' -> ('82996506901', 'AMANDA KESIA CUNHA ROCHA')
    """
    if not val:
        return None, None
    s = str(val).strip()
    s_clean = re.sub(r'[\r\n\t]+', ' ', s).strip()
    s_single_space = re.sub(r'\s{2,}', ' ', s_clean)

    prefix_pattern = r'(?:(?:whatsapp|whats|wpp|telefone|telefones|celulares|celular|contatos|contato|mobile|phone|fone|fixo|fixos|ramal|voice|work|home|cell|mob|tel|cel|zap|fax|p|m|c|t|w|f|v)\s*[:.-]?\s*)?'
    phone_pattern = r'(?:(?:\+|00)?55[\s.-]?)?(?:\(?0?([1-9]{2})\)?[\s.-]?)?(?:9[\s.-]?)?(?:(9\d{4}|\d{4})[\s.-]?(\d{4}))'

    # 1. Telefone no início (com ou sem prefixos)
    m_start = re.match(
        rf'^{prefix_pattern}{phone_pattern}\s*[-–—/:]?\s*([A-Za-zÀ-ÿ].*)$',
        s_single_space,
        flags=re.IGNORECASE
    )
    if m_start:
        name_part = m_start.group(4).strip()
        name_part = re.sub(r'^[-–—/:\(\)\s]+|[-–—/:\(\)\s]+$', '', name_part)
        if is_annotation_only_text(name_part):
            return None, None
        phone_raw = s_single_space[:len(s_single_space) - len(name_part)]
        digits = normalize_phone_value(phone_raw)
        if len(digits) in (10, 11) and len(name_part) >= 3 and not name_part.isdigit():
            return digits, name_part

    # 2. Telefone no final (com ou sem prefixos ou parênteses)
    m_end = re.match(
        rf'^([A-Za-zÀ-ÿ].*?)\s*[-–—/:]?\s*\(?{prefix_pattern}{phone_pattern}\)?$',
        s_single_space,
        flags=re.IGNORECASE
    )
    if m_end:
        name_part = m_end.group(1).strip()
        name_part = re.sub(r'^[-–—/:\(\)\s]+|[-–—/:\(\)\s]+$', '', name_part)
        if is_annotation_only_text(name_part):
            return None, None
        phone_raw = s_single_space[len(name_part):]
        digits = normalize_phone_value(phone_raw)
        if len(digits) in (10, 11) and len(name_part) >= 3 and not name_part.isdigit():
            return digits, name_part

    # 3. Telefone numérico puro no início (ex: '82996506901 AMANDA KESIA' ou '+5582996506901 AMANDA' ou 'p:+5582996506901 AMANDA')
    m_digits_start = re.match(r'^(?:(?:whatsapp|whats|wpp|telefone|telefones|celulares|celular|contatos|contato|mobile|phone|fone|fixo|fixos|ramal|voice|work|home|cell|mob|tel|cel|zap|fax|p|m|c|t|w|f|v)\s*[:.-]?\s*)?(?:(?:\+|00)?55)?([1-9]{2}9?\d{8})\s+([A-Za-zÀ-ÿ].*)$', s_single_space, flags=re.IGNORECASE)
    if m_digits_start:
        digits = m_digits_start.group(1)
        if len(digits) == 10 and digits[2] in ("6", "7", "8", "9"):
            digits = f"{digits[:2]}9{digits[2:]}"
        name_part = m_digits_start.group(2).strip()
        name_part = re.sub(r'^[-–—/:\(\)\s]+|[-–—/:\(\)\s]+$', '', name_part)
        if is_annotation_only_text(name_part):
            return None, None
        if len(name_part) >= 3 and not name_part.isdigit():
            return digits, name_part

    return None, None


def is_cpf_field(field_name: Any) -> bool:
    norm = normalize_field_name(field_name)
    if not norm:
        return False
    return norm in {
        "cpf", "num_cpf", "nr_cpf", "documento", "doc_cpf", "cpf_cnpj",
        "numero_cpf", "cpf_cliente", "cpf_titular", "cpf_candidato",
        "doc_identificacao", "enriquecido_cpf", "nr_documento", "num_documento",
        "doc", "cpf_completo", "cpf_num", "cpf_numero", "inscricao_cpf", "cpf_aluno",
        "cpf_paciente", "cpf_associado", "cpf_funcionario", "cpf_motorista"
    } or norm.startswith("cpf_") or norm.endswith("_cpf") or norm == "cpf"



def normalize_field_name(field_name: Any) -> str:
    if field_name is None:
        return ""
    return normalize_text(str(field_name))


def is_required_name_field(field_name: Any) -> bool:
    norm = normalize_field_name(field_name)
    if not norm:
        return False
    # Evitar campos de status/resultado que contenham "candidato" ou "cliente" (ex: "situacao_candidato", "status_cliente")
    if any(p in norm for p in ["situacao", "status", "resultado", "parecer", "justificativa", "motivo", "vaga", "cargo"]):
        if not any(norm == f"nome_{p}" or norm.startswith("nome_") for p in ["candidato", "cliente", "titular"]):
            return False
    return (
        norm in {
            "nome", "nome_completo", "nome_cliente", "cliente", "razao_social", "razao",
            "contato", "titular", "pagador", "enriquecido_nome", "enriquecido_nome_completo",
            "candidato", "candidatos", "nome_do_candidato", "nome_candidato", "nome_da_pessoa",
            "nome_titular", "nome_favorecido", "nome_solicitante", "nome_proponente",
            "nome_aluno", "nome_paciente", "nome_associado", "nome_motorista",
            "nome_funcionario", "nome_colaborador", "nome_servidor", "pessoa", "favorecido",
            "proponente", "solicitante", "aluno", "paciente", "associado", "motorista",
            "funcionario", "colaborador", "servidor", "inscrito", "participante"
        }
        or norm.startswith("nome_")
        or "nome" in norm
        or "candidato" in norm
        or "cliente" in norm
        or "razao" in norm
        or "contato" in norm
        or "titular" in norm
        or "pagador" in norm
        or "solicitante" in norm
        or "proponente" in norm
        or "favorecido" in norm
        or "paciente" in norm
        or "associado" in norm
    )


def is_cellphone_field(field_name: Any) -> bool:
    norm = normalize_field_name(field_name)
    if not norm:
        return False
    if norm in {"ddd", "ddd_telefone", "ddd_celular", "ddd_fone", "ddd_tel"} or (norm.startswith("ddd") and not any(p in norm for p in ["cel", "tel", "fone"])):
        return False
    if norm in {"celular", "cel", "cel1", "cel_1", "cel2", "cel_2", "celular1", "celular2", "tel_celular", "telefone_celular", "whatsapp", "wpp", "whats", "movel", "celular_contato", "cel_contato"}:
        return True
    if any(p in norm for p in ["celular", "cel", "whatsapp", "wpp", "whats", "movel"]):
        return True
    return False


def is_required_phone_field(field_name: Any) -> bool:
    norm = normalize_field_name(field_name)
    if not norm:
        return False
    if norm in {"ddd", "ddd_telefone", "ddd_celular", "ddd_fone", "ddd_tel"} or (norm.startswith("ddd") and not any(p in norm for p in ["cel", "tel", "fone"])):
        return False
    if norm in {"telefone", "celular", "fone", "tel", "cel", "phone", "whatsapp", "wpp"}:
        return True
    if any(p in norm for p in ["telefone", "celular", "fone", "tel", "cel", "phone"]):
        return True
    return False


def get_field_value(row: Dict[str, Any], field: str) -> Any:
    if not row or field is None:
        return ""
    str_field = str(field).strip()
    if str_field in row and row[str_field] is not None and row[str_field] != "":
        return row[str_field]
    norm = normalize_field_name(str_field)
    if norm in row and row[norm] is not None and row[norm] != "":
        return row[norm]
    for k, v in row.items():
        if normalize_field_name(k) == norm and v is not None and v != "":
            return v
    if norm.startswith("enriquecido_"):
        suffix = norm.replace("enriquecido_", "", 1)
        if suffix in row and row[suffix] is not None and row[suffix] != "":
            return row[suffix]
        for k, v in row.items():
            if normalize_field_name(k) == suffix and v is not None and v != "":
                return v
    return ""


def normalize_selected_fields(selected_fields: List[str]) -> List[str]:
    ordered_fields: List[str] = []
    seen: set[str] = set()

    for field in selected_fields or []:
        if field is None:
            continue
        value = str(field).strip()
        if not value or value in seen:
            continue
        seen.add(value)
        ordered_fields.append(value)

    return ordered_fields




def extract_cpf_and_name_from_combined_text(val: Any) -> Tuple[Optional[str], Optional[str]]:
    """
    Detecta se um texto contém simultaneamente um CPF e um Nome real de pessoa.
    Exemplos:
      '123.456.789-00 AMANDA KESIA CUNHA ROCHA' -> ('12345678900', 'AMANDA KESIA CUNHA ROCHA')
      '12345678900 AMANDA KESIA CUNHA ROCHA' -> ('12345678900', 'AMANDA KESIA CUNHA ROCHA')
      'AMANDA KESIA CUNHA ROCHA 123.456.789-00' -> ('12345678900', 'AMANDA KESIA CUNHA ROCHA')
    """
    if not val:
        return None, None
    s = str(val).strip()
    s_clean = re.sub(r'[\r\n\t]+', ' ', s).strip()
    s_single_space = re.sub(r'\s{2,}', ' ', s_clean)

    # CPF no início
    m_cpf_start = re.match(r'^(\d{3}\.?\d{3}\.?\d{3}-?\d{2})\s*[-–—/:]?\s*([A-Za-zÀ-ÿ].*)$', s_single_space)
    if m_cpf_start:
        name_part = m_cpf_start.group(2).strip()
        name_part = re.sub(r'^[-–—/:\(\)\s]+|[-–—/:\(\)\s]+$', '', name_part)
        if is_annotation_only_text(name_part) or '+' in name_part or ':' in name_part:
            return None, None
        cpf_raw = m_cpf_start.group(1)
        digits = re.sub(r'\D', '', cpf_raw)
        if len(digits) == 11 and len(name_part) >= 3 and not name_part.isdigit():
            return digits, name_part

    # CPF no final
    m_cpf_end = re.match(r'^([A-Za-zÀ-ÿ].*?)\s*[-–—/:]?\s*(\d{3}\.?\d{3}\.?\d{3}-?\d{2})$', s_single_space)
    if m_cpf_end:
        name_part = m_cpf_end.group(1).strip()
        name_part = re.sub(r'^[-–—/:\(\)\s]+|[-–—/:\(\)\s]+$', '', name_part)
        if is_annotation_only_text(name_part) or '+' in name_part or ':' in name_part:
            return None, None
        cpf_raw = m_cpf_end.group(2)
        digits = re.sub(r'\D', '', cpf_raw)
        if len(digits) == 11 and len(name_part) >= 3 and not name_part.isdigit():
            return digits, name_part

    return None, None


def normalize_name_value(value: Any) -> str:
    """Higienização inteligente de Nome:
    - Corrige formato invertido (ex: 'SILVA, CARLOS EDUARDO' -> 'CARLOS EDUARDO SILVA')
    - Remove telefones ou CPFs acoplados no início/final do nome
    - Remove títulos e prefixos honoríficos (Dr., Dra., Prof., Sr., Sra., etc.)
    - Remove acentuações e quebras de texto/linha preservando espaçamento padrão.
    """
    if value is None:
        return ""
    text = ILLEGAL_CHARACTERS_RE.sub("", str(value))
    # 1. Substituir quebras de linha e tabulações por espaço
    text = re.sub(r"[\r\n\t\v\f]+", " ", text).strip()
    
    # 2. Remover telefone acoplado ao nome se presente
    _, clean_phone_name = extract_phone_and_name_from_combined_text(text)
    if clean_phone_name:
        text = clean_phone_name
    else:
        # Remover CPF acoplado ao nome se presente
        _, clean_cpf_name = extract_cpf_and_name_from_combined_text(text)
        if clean_cpf_name:
            text = clean_cpf_name

    # 3. Corrigir formato invertido com vírgula (ex: "SILVA, CARLOS EDUARDO" -> "CARLOS EDUARDO SILVA")
    if "," in text:
        parts = text.split(",", 1)
        part0 = parts[0].strip()
        part1 = parts[1].strip()
        if part0 and part1 and not re.search(r"^\d+$", part0) and not re.search(r"^\d+$", part1):
            text = f"{part1} {part0}"
            
    # 4. Remover títulos e prefixos de tratamento no início do nome
    text = re.sub(
        r"^(?:DR[AªOº]?\b|DR[AªOº]?\.|PROF[AªOº]?\b|PROF[AªOº]?\.|ME[AªOº]?\b|ME[AªOº]?\.|MS\b|MS\.|SR[AªOº]?\b|SR[AªOº]?\.|ENG[AªOº]?\b|ENG[AªOº]?\.|ILM[AªOº]?\b|ILM[AªOº]?\.|EXM[AªOº]?\b|EXM[AªOº]?\.|ADV[AªOº]?\b|ADV[AªOº]?\.)\s+",
        "",
        text,
        flags=re.IGNORECASE
    )
    
    # 5. Remover acentuações (À->A, é->e, ç->c, etc.)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    
    # 6. Remover números e caracteres especiais residuais agregados ao nome (ex: 'Eliana Silva31' -> 'Eliana Silva')
    text = re.sub(r"[^A-Za-z\s]", " ", text)
    
    # 7. Remover espaços múltiplos e nas extremidades
    text = re.sub(r"\s+", " ", text).strip()
    return text


def clean_name_for_api(vtexto: Any) -> str:
    """Equivalente direto da função RETIRA do VBA com higienização inteligente.
    
    Corrige nomes invertidos, remove títulos/prefixos, remove acentos, números e caracteres especiais,
    deixando apenas letras e espaços em maiúsculo.
    """
    if vtexto is None:
        return ""
    # 1. Normalizar nome inteligente
    text = normalize_name_value(vtexto)
    
    # 2. Remover números e caracteres especiais (0-9, .,-;:/\!@#$%¨&*()[]{}?^~´`<>|_=+""")
    text = re.sub(r"[^a-zA-Z\s]", " ", text)
    
    # 3. Normalizar espaços múltiplos e converter para maiúsculo
    text = re.sub(r"\s+", " ", text).strip()
    return text.upper()


def resolve_name_value(row: Dict[str, Any]) -> str:
    if not row:
        return ""
    for k, v in row.items():
        if is_required_name_field(k) and v not in (None, ""):
            val = str(v).strip()
            if val:
                return normalize_name_value(val)
    for alt in (
        "nome", "nome_completo", "nome_cliente", "cliente", "razao_social", "contato",
        "titular", "pagador", "candidato", "nome_do_candidato", "nome_candidato",
        "solicitante", "proponente", "aluno", "paciente", "associado", "funcionario"
    ):
        val = get_field_value(row, alt)
        if val not in (None, ""):
            val = str(val).strip()
            if val:
                return normalize_name_value(val)
    return ""


BRAZILIAN_UF_SIGLAS = {
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG",
    "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"
}

STATE_NAME_TO_UF = {
    "ACRE": "AC", "ALAGOAS": "AL", "AMAPA": "AP", "AMAZONAS": "AM", "BAHIA": "BA", "CEARA": "CE",
    "DISTRITO FEDERAL": "DF", "ESPIRITO SANTO": "ES", "GOIAS": "GO", "MARANHAO": "MA", "MATO GROSSO": "MT",
    "MATO GROSSO DO SUL": "MS", "MINAS GERAIS": "MG", "PARA": "PA", "PARAIBA": "PB", "PARANA": "PR",
    "PERNAMBUCO": "PE", "PIAUI": "PI", "RIO DE JANEIRO": "RJ", "RIO GRANDE DO NORTE": "RN",
    "RIO GRANDE DO SUL": "RS", "RONDONIA": "RO", "RORAIMA": "RR", "SANTA CATARINA": "SC",
    "SAO PAULO": "SP", "SERGIPE": "SE", "TOCANTINS": "TO"
}


def is_uf_field(field_name: Any) -> bool:
    norm = normalize_field_name(field_name)
    if not norm:
        return False
    # Evitar falsos positivos com palavras que contenham 'estado' como substring
    if any(p in norm for p in ["resultado", "prestador", "testador", "qualificacao", "status", "situacao", "motivo"]):
        return False
    if norm in {
        "uf", "estado", "sigla_uf", "uf_estado", "sigla", "sigla_estado", 
        "estado_sigla", "uf_cliente", "uf_residencia", "estado_cliente", "uf_nasc", "uf_nascimento"
    }:
        return True
    if norm.startswith("uf_") or norm.endswith("_uf") or norm.startswith("estado_") or norm.endswith("_estado"):
        return True
    return False


def normalize_uf_string(val: Any) -> str:
    if val is None or val == "":
        return ""
    raw = str(val).strip().upper()
    if not raw:
        return ""
    if raw in BRAZILIAN_UF_SIGLAS:
        return raw
    if len(raw) >= 2 and raw[:2] in BRAZILIAN_UF_SIGLAS and (len(raw) == 2 or raw[2] in " -/_:()"):
        return raw[:2]
    if len(raw) >= 2 and raw[-2:] in BRAZILIAN_UF_SIGLAS and raw[-3] in " -/_:()":
        return raw[-2:]
    clean_norm = normalize_text(raw).upper().replace("_", " ")
    if clean_norm in STATE_NAME_TO_UF:
        return STATE_NAME_TO_UF[clean_norm]
    for state_name, sigla in STATE_NAME_TO_UF.items():
        if state_name in clean_norm:
            return sigla
    return ""


def resolve_uf_value(row: Dict[str, Any]) -> str:
    if not row:
        return ""
    # 1. Busca por campos identificados estritamente como UF
    for k, v in row.items():
        if is_uf_field(k) and v not in (None, ""):
            clean_uf = normalize_uf_string(v)
            if clean_uf:
                return clean_uf
    # 2. Busca por nomes alternativos comuns
    for alt in ("uf", "estado", "sigla_uf", "sigla_estado", "uf_estado", "estado_sigla", "sigla", "uf_cliente", "estado_cliente"):
        val = get_field_value(row, alt)
        if val not in (None, ""):
            clean_uf = normalize_uf_string(val)
            if clean_uf:
                return clean_uf
    # 3. Varredura por célula contendo sigla de UF válida de 2 letras
    for k, v in row.items():
        if v not in (None, "") and not str(k).startswith("enriquecido_") and not str(k).startswith("_"):
            clean_uf = normalize_uf_string(v)
            if clean_uf and len(str(v).strip()) <= 30:
                norm_k = normalize_field_name(k)
                if not any(p in norm_k for p in ["nome", "cpf", "tel", "cel", "data", "nasc", "rua", "logradouro", "bairro", "email"]):
                    return clean_uf
    return ""


def resolve_cidade_value(row: Dict[str, Any]) -> str:
    if not row:
        return ""
    for k, v in row.items():
        norm = normalize_field_name(k)
        if norm in {"cidade", "municipio", "nome_cidade"} and v not in (None, ""):
            val = str(v).strip()
            if val:
                return val
    for alt in ("cidade", "municipio"):
        val = get_field_value(row, alt)
        if val not in (None, ""):
            val = str(val).strip()
            if val:
                return val
    return ""


def extract_clean_cpf(val: Any) -> str:
    if val is None or val == "":
        return ""
    digits = re.sub(r"\D", "", str(val))
    if not digits or re.match(r"^0+$", digits):
        return ""
    if len(digits) == 12 and digits.startswith("0"):
        digits = digits[1:]
    if 9 <= len(digits) <= 11:
        return digits.zfill(11)
    if len(digits) == 11 and validate_cpf_checksum(digits):
        return digits
    return ""


def resolve_cpf_value(row: Dict[str, Any]) -> str:
    if not row:
        return ""
    # 1. Busca por campos identificados como CPF
    for k, v in row.items():
        if v not in (None, ""):
            norm = normalize_field_name(k)
            if is_cpf_field(norm):
                clean_cpf = extract_clean_cpf(v)
                if clean_cpf:
                    return clean_cpf

    # 2. Busca por nomes alternativos comuns
    for alt in (
        "cpf", "documento", "doc", "num_cpf", "nr_cpf", "cpf_cnpj", "numero_cpf",
        "cpf_cliente", "cpf_titular", "cpf_candidato", "num_documento", "nr_documento",
        "doc_identificacao", "inscricao", "cpf_completo", "cpf_num", "cpf_numero"
    ):
        val = get_field_value(row, alt)
        if val not in (None, ""):
            clean_cpf = extract_clean_cpf(val)
            if clean_cpf:
                return clean_cpf

    # 3. Fallback inteligente: se houver alguma célula na linha que contenha CPF válido com checksum
    for k, v in row.items():
        if v not in (None, ""):
            clean_cpf = extract_clean_cpf(v)
            if clean_cpf and validate_cpf_checksum(clean_cpf):
                return clean_cpf
            if clean_cpf and ("cpf" in str(k).lower() or "doc" in str(k).lower()):
                return clean_cpf

    return ""




def resolve_phone_value(row: Dict[str, Any]) -> str:
    if not row:
        return ""

    ddd_digits = ""
    for key, value in row.items():
        norm = normalize_field_name(key)
        if norm in {"ddd", "ddd_telefone", "ddd_celular", "ddd_cel1", "ddd_fone", "ddd_tel"} or (norm.startswith("ddd") and any(p in norm for p in ["cel", "tel", "fone"])):
            digits = re.sub(r"\D", "", str(value or ""))
            if len(digits) == 2 and not digits.startswith("0"):
                ddd_digits = digits
                break

    if not ddd_digits:
        for key, value in row.items():
            norm = normalize_field_name(key)
            if norm == "ddd" or (norm.startswith("ddd") and not any(p in norm for p in ["cel", "tel", "fone"])):
                digits = re.sub(r"\D", "", str(value or ""))
                if len(digits) == 2 and not digits.startswith("0"):
                    ddd_digits = digits
                    break

    # Se enriquecido_telefone já existir e for celular válido (11 dígitos começando com 9)
    enr_tel = normalize_phone_value(row.get("enriquecido_telefone"), ddd_digits)
    if len(enr_tel) == 11 and enr_tel[2] == "9" and not re.match(r"^0+$", enr_tel):
        return enr_tel

    # 1. Prioridade Máxima: Colunas específicas de Celular
    for key, value in row.items():
        norm = normalize_field_name(key)
        if norm.startswith("ddd"):
            continue
        if is_cellphone_field(key):
            clean_p = normalize_phone_value(value, ddd_digits)
            if clean_p and len(clean_p) == 11 and clean_p[2] == "9" and not re.match(r"^0+$", clean_p):
                return clean_p

    for alt in ("celular", "cel", "cel1", "tel_celular", "telefone_celular", "whatsapp", "wpp", "movel", "celular_contato"):
        val = get_field_value(row, alt)
        if val not in (None, ""):
            clean_p = normalize_phone_value(val, ddd_digits)
            if clean_p and len(clean_p) == 11 and clean_p[2] == "9" and not re.match(r"^0+$", clean_p):
                return clean_p

    # 2. Segunda Prioridade: Outras colunas de telefone que sejam celulares válidos
    for key, value in row.items():
        norm = normalize_field_name(key)
        if norm.startswith("ddd"):
            continue
        if is_required_phone_field(key):
            clean_p = normalize_phone_value(value, ddd_digits)
            if clean_p and len(clean_p) == 11 and clean_p[2] == "9" and not re.match(r"^0+$", clean_p):
                return clean_p

    for alt in ("telefone", "fone", "tel", "phone", "telefone_contato"):
        val = get_field_value(row, alt)
        if val not in (None, ""):
            clean_p = normalize_phone_value(val, ddd_digits)
            if clean_p and len(clean_p) == 11 and clean_p[2] == "9" and not re.match(r"^0+$", clean_p):
                return clean_p

    # 3. Terceira Prioridade: Telefones fixos válidos (10 dígitos) caso não haja celular
    for key, value in row.items():
        norm = normalize_field_name(key)
        if norm.startswith("ddd"):
            continue
        if is_required_phone_field(key):
            clean_p = normalize_phone_value(value, ddd_digits)
            if clean_p and len(clean_p) in (10, 11) and not re.match(r"^0+$", clean_p):
                return clean_p

    for alt in ("telefone", "fone", "tel", "phone", "telefone_contato"):
        val = get_field_value(row, alt)
        if val not in (None, ""):
            clean_p = normalize_phone_value(val, ddd_digits)
            if clean_p and len(clean_p) in (10, 11) and not re.match(r"^0+$", clean_p):
                return clean_p

    return ""



class FastConcatenator:
    def __init__(self, selected_fields: List[str]):
        self.selected_fields = selected_fields or []
        self.name_fields = [f for f in self.selected_fields if is_required_name_field(f)]
        self.phone_fields = [
            f for f in self.selected_fields
            if is_required_phone_field(f) and normalize_field_name(f) not in {"ddd", "ddd_telefone", "ddd_celular", "ddd_fone", "ddd_tel"}
        ]
        used = set(self.name_fields + self.phone_fields + [
            "enriquecido_nome", "nome", "nome_completo", "nome_cliente", "cliente", "razao_social", "razao",
            "contato", "titular", "pagador", "candidato", "nome_do_candidato", "nome_candidato"
        ])
        
        self.remaining = []
        for f in self.selected_fields:
            if f not in used and normalize_field_name(f) not in {"ddd", "ddd_telefone", "ddd_celular", "ddd_fone", "ddd_tel"}:
                self.remaining.append((f, labelize(f, strip_prefix=True), is_required_name_field(f)))

    def format(self, row: Dict[str, Any]) -> str:
        if not isinstance(row, dict):
            return ""

        # 1. Nome
        nome_val = ""
        for f in self.name_fields:
            v = row.get(f)
            if v:
                nome_val = normalize_name_value(v)
                break
        if not nome_val:
            nome_val = normalize_name_value(row.get("nome") or row.get("enriquecido_nome") or "")

        # 2. Telefone
        if len(self.phone_fields) <= 1:
            tel_val = ""
            if self.phone_fields:
                tel_val = re.sub(r"\D", "", str(row.get(self.phone_fields[0]) or ""))
            if not tel_val or len(tel_val) not in (10, 11):
                tel_val = re.sub(r"\D", "", str(row.get("telefone") or row.get("enriquecido_telefone") or row.get("celular") or ""))
            result = f"{nome_val}|{tel_val}|"
        else:
            slots = ["", "", "", "", ""]
            for i in range(min(5, len(self.phone_fields))):
                slots[i] = re.sub(r"\D", "", str(row.get(self.phone_fields[i]) or ""))
            result = f"{nome_val}|" + "|".join(slots) + "|"

        # 3. Demais colunas
        parts = []
        for f, label, is_name in self.remaining:
            v = row.get(f)
            if v is not None:
                v_str = ILLEGAL_CHARACTERS_RE.sub("", str(v).strip())
                if v_str != "":
                    if is_name:
                        v_str = normalize_name_value(v_str)
                    elif "\n" in v_str or "\r" in v_str or "\t" in v_str or "\v" in v_str or "\f" in v_str:
                        v_str = re.sub(r"[\r\n\t\v\f]+", " ", v_str).strip()
                    if v_str:
                        parts.append(f"{label}: {v_str}")

        if parts:
            result += " " + " ".join(parts)

        return result


def format_concatenated_fields(row: Dict[str, Any], selected_fields: List[str]) -> str:
    concatenator = FastConcatenator(selected_fields)
    return concatenator.format(row)



def detect_enrichment_type(columns: List[str]) -> str:
    combined = " ".join(normalize_text(col) for col in columns)
    has_nome = "nome" in combined or "cliente" in combined or "razao" in combined or "contato" in combined
    has_uf = "uf" in combined or "estado" in combined
    has_cpf = "cpf" in combined
    has_telefone = any(t in combined for t in ["telefone", "celular", "fone", "tel", "cel"])
    
    if has_cpf:
        return "CPF"
    if has_nome and has_uf:
        return "NOME_UF"
    if has_nome and has_telefone:
        return "NOME_TELEFONE"
    if has_nome:
        return "NOME"
    if has_telefone:
        return "TELEFONE"
    if has_uf:
        return "UF"
    return "FILA_PADRAO"


def generate_mock_row(row: Dict[str, Any], detected_type: str) -> Dict[str, Any]:
    raw_name = get_field_value(row, "nome") or get_field_value(row, "nome_completo") or get_field_value(row, "nome_cliente") or f"Contato {uuid.uuid4().hex[:6]}"
    name = normalize_name_value(raw_name)
    
    # Gerar DDD e telefone aleatórios reais se não informados
    ddd = random.choice(["11", "12", "19", "21", "22", "27", "31", "32", "41", "47", "51", "61", "62", "71", "81", "85", "91", "98"])
    raw_phone = get_field_value(row, "telefone") or get_field_value(row, "celular") or get_field_value(row, "telefone_contato") or f"{ddd}9{random.randint(6000, 9999)}{random.randint(1000, 9999)}"
    phone = re.sub(r"\D", "", str(raw_phone))
    
    # Telefones adicionais
    phone_2 = str(get_field_value(row, "telefone_2") or get_field_value(row, "celular_2") or "")
    phone_3 = str(get_field_value(row, "telefone_3") or get_field_value(row, "celular_3") or "")
    phone_4 = str(get_field_value(row, "telefone_4") or "")
    phone_5 = str(get_field_value(row, "telefone_5") or "")
    
    cpf_value = get_field_value(row, "cpf") or f"{(uuid.uuid4().int % 99999999999):011d}"
    
    # Lista de CBO e Profissões associadas
    cbo_professions = [
        ("411005", "Auxiliar de Escritório"),
        ("782310", "Motorista de Furgão ou Veículo Similar"),
        ("225250", "Médico Ginecologista e Obstetra"),
        ("322205", "Técnico de Enfermagem"),
        ("142105", "Gerente Administrativo"),
        ("212405", "Analista de Desenvolvimento de Sistemas"),
        ("231205", "Professor da Educação Infantil"),
        ("354125", "Supervisor de Vendas Comercial")
    ]
    cbo_pair = random.choice(cbo_professions)
    cbo_value = get_field_value(row, "cbo") or cbo_pair[0]
    profession_value = get_field_value(row, "profissao") or get_field_value(row, "possivel_profissao") or cbo_pair[1]
    
    # Gerar idade e data de nascimento aleatórias realistas
    age = random.randint(18, 65)
    year = datetime.now().year - age
    month = random.randint(1, 12)
    day = random.randint(1, 28)
    data_nascimento = format_to_brazilian_date(get_field_value(row, "data_nascimento") or f"{day:02d}/{month:02d}/{year}")
    
    sex_value = get_field_value(row, "sexo") or random.choice(["MASCULINO", "FEMININO"])
    
    clean_name_part = normalize_text(name).replace("_", "")
    email_val = get_field_value(row, "email") or f"{clean_name_part[:12]}@{random.choice(['gmail.com', 'hotmail.com', 'outlook.com', 'yahoo.com.br'])}"
    email_val = normalize_email_value(email_val)
    
    # Endereços realistas
    streets = [
        "Luiz Rodrigues dos Santos", "Maria Matos", "Doutor Joao Mangabeira", 
        "Analia Ferreira Ferro", "Avenida Brasil", "Rua das Flores", 
        "Rua Rio de Janeiro", "Avenida Paulista"
    ]
    neighborhoods = [
        "Todos os Santos", "Centro", "Brasilia", "Praia do Morro", 
        "Jardim Paulista", "Copacabana", "Boa Viagem", "Savassi"
    ]
    
    logradouro_val = get_field_value(row, "logradouro") or get_field_value(row, "endereco") or random.choice(streets)
    numero_val = get_field_value(row, "numero") or str(random.randint(10, 1500))
    complemento_val = get_field_value(row, "complemento") or random.choice(["AP 401", "AP 301", "Casa A", "Bloco B", "SL 12", ""])
    bairro_val = get_field_value(row, "bairro") or random.choice(neighborhoods)
    cep_val = get_field_value(row, "cep") or f"{random.randint(10000, 99999):05d}{random.randint(100, 999):03d}"
    
    # Cidades e estados coordenados do Brasil
    cities_states = [
        ("São Paulo", "SP"),
        ("Campinas", "SP"),
        ("Santos", "SP"),
        ("Rio de Janeiro", "RJ"),
        ("Niterói", "RJ"),
        ("Belo Horizonte", "MG"),
        ("Uberlândia", "MG"),
        ("Curitiba", "PR"),
        ("Londrina", "PR"),
        ("Porto Alegre", "RS"),
        ("Caxias do Sul", "RS"),
        ("Salvador", "BA"),
        ("Feira de Santana", "BA"),
        ("Recife", "PE"),
        ("Olinda", "PE"),
        ("Fortaleza", "CE"),
        ("Brasília", "DF"),
        ("Goiânia", "GO"),
        ("Manaus", "AM"),
        ("Belém", "PA"),
    ]
    
    city_val = get_field_value(row, "cidade") or get_field_value(row, "municipio")
    state_val = get_field_value(row, "uf") or get_field_value(row, "estado")
    
    if not city_val and not state_val:
        city_val, state_val = random.choice(cities_states)
    elif state_val and not city_val:
        state_upper = str(state_val).strip().upper()
        matching = [c for c, s in cities_states if s == state_upper]
        city_val = random.choice(matching) if matching else f"Cidade de {state_val}"
    elif city_val and not state_val:
        city_lower = normalize_text(city_val).lower()
        matching = [s for c, s in cities_states if normalize_text(c).lower() == city_lower]
        state_val = random.choice(matching) if matching else "SP"
        
    renda_val = normalize_renda_estimada(
        get_field_value(row, "renda_estimada")
        or get_field_value(row, "renda")
        or random.choice([
            "1.000 AT 2.000",
            "2.001 AT 3.000",
            "3.001 AT 5.000",
            "5.001 AT 8.000",
            "18.001 AT 19.000",
        ])
    )
    
    score_val = get_field_value(row, "score") or str(random.randint(200, 990))
    perfil_list = ["Sempre Presente", "O Bem Amado", "Quem Sou Eu", "Decisor", "Influenciador"]
    perfil = get_field_value(row, "perfil") or get_field_value(row, "personadecredito") or random.choice(perfil_list)
    status_list = ["ativo", "inativo", "pendente"]
    status = get_field_value(row, "status") or random.choice(status_list)
    obito_val = get_field_value(row, "obito") or get_field_value(row, "flag_de_obito") or random.choices(["0", "1"], weights=[98, 2])[0]
    whatsapp_val = get_field_value(row, "whatsapp") or random.choices(["SIM", "NÃO"], weights=[90, 10])[0]

    normalized = {
        "enriquecido_nome": name,
        "enriquecido_telefone": phone,
        "enriquecido_telefone_2": phone_2,
        "enriquecido_telefone_3": phone_3,
        "enriquecido_telefone_4": phone_4,
        "enriquecido_telefone_5": phone_5,
        "enriquecido_cpf": cpf_value,
        "enriquecido_data_nascimento": data_nascimento,
        "enriquecido_idade": str(age),
        "enriquecido_sexo": sex_value,
        "enriquecido_email": email_val,
        "enriquecido_logradouro": logradouro_val,
        "enriquecido_numero": numero_val,
        "enriquecido_complemento": complemento_val,
        "enriquecido_bairro": bairro_val,
        "enriquecido_cidade": city_val,
        "enriquecido_estado": state_val,
        "enriquecido_cep": cep_val,
        "enriquecido_renda_estimada": renda_val,
        "enriquecido_cbo": cbo_value,
        "enriquecido_profissao": profession_value,
        "enriquecido_score": score_val,
        "enriquecido_perfil": perfil,
        "enriquecido_obito": obito_val,
        "enriquecido_whatsapp": whatsapp_val,
        "enriquecido_status": status,
        "enriquecido_similaridade_nome": "ALTA (100%)",
        "tipo_enriquecimento": detected_type,
    }

    for key, value in list(row.items()):
        if key is not None:
            if isinstance(value, (datetime, date)):
                clean_v = format_to_brazilian_date(value)
            elif isinstance(value, dt_time):
                clean_v = value.strftime("%H:%M:%S")
            else:
                clean_v = value
            normalized[key] = clean_v
            norm_k = normalize_text(key)
            if norm_k:
                normalized[norm_k] = clean_v

    return normalized


def clean_dataframe(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not rows:
        return []

    # 0. Detecção e separação de colunas combinadas (ex: 'telefone nome', '82996506901 AMANDA KESIA CUNHA ROCHA')
    sample_rows = rows[:100]
    raw_keys = list(rows[0].keys())

    split_plans = {}
    for k in raw_keys:
        if k is None:
            continue
        clean_k = normalize_text(repair_mojibake_text(ILLEGAL_CHARACTERS_RE.sub("", str(k)).strip()))
        
        is_explicit_phone_name = any(clean_k == p or clean_k.startswith(p + "_") or clean_k.endswith("_" + p) for p in [
            "telefone_nome", "nome_telefone", "tel_nome", "nome_tel", "celular_nome", "nome_celular", 
            "contato_nome", "nome_contato", "telefone_e_nome", "nome_e_telefone"
        ])
        is_explicit_cpf_name = any(clean_k == p or clean_k.startswith(p + "_") or clean_k.endswith("_" + p) for p in [
            "cpf_nome", "nome_cpf", "cpf_e_nome", "nome_e_cpf"
        ])

        phone_name_matches = 0
        cpf_name_matches = 0
        non_empty_count = 0

        for r in sample_rows:
            v = r.get(k)
            if v is not None and str(v).strip():
                non_empty_count += 1
                p_digits, n_part = extract_phone_and_name_from_combined_text(v)
                if p_digits and n_part:
                    phone_name_matches += 1
                else:
                    c_digits, cn_part = extract_cpf_and_name_from_combined_text(v)
                    if c_digits and cn_part:
                        cpf_name_matches += 1

        if non_empty_count > 0:
            min_matches = 1 if len(sample_rows) == 1 else 2
            if (phone_name_matches >= max(min_matches, int(non_empty_count * 0.30))) or (is_explicit_phone_name and phone_name_matches > 0):
                split_plans[k] = "phone_name"
            elif (cpf_name_matches >= max(min_matches, int(non_empty_count * 0.30))) or (is_explicit_cpf_name and cpf_name_matches > 0):
                split_plans[k] = "cpf_name"

    # Se houver colunas combinadas identificadas, reformata as linhas separando os campos
    if split_plans:
        transformed_rows = []
        for r in rows:
            new_r = {}
            for k in raw_keys:
                if k is None:
                    continue
                v = r.get(k)
                if k in split_plans:
                    plan = split_plans[k]
                    if plan == "phone_name":
                        p_digits, n_part = extract_phone_and_name_from_combined_text(v)
                        if p_digits and n_part:
                            new_r["Nome"] = n_part
                            new_r["Telefone"] = p_digits
                        else:
                            new_r["Nome"] = str(v) if v is not None else ""
                            if "Telefone" not in new_r:
                                new_r["Telefone"] = ""
                    elif plan == "cpf_name":
                        c_digits, n_part = extract_cpf_and_name_from_combined_text(v)
                        if c_digits and n_part:
                            new_r["CPF"] = c_digits
                            new_r["Nome"] = n_part
                        else:
                            new_r["Nome"] = str(v) if v is not None else ""
                            if "CPF" not in new_r:
                                new_r["CPF"] = ""
                else:
                    new_r[k] = v
            transformed_rows.append(new_r)
        rows = transformed_rows

    # 1. Pré-computação de metadados das colunas (executado uma única vez a partir de amostra representativa)
    raw_keys: List[str] = []
    seen_keys: set[str] = set()
    sample_for_keys = rows[:100]
    for r in sample_for_keys:
        for k in r.keys():
            if k is not None and k not in seen_keys:
                seen_keys.add(k)
                raw_keys.append(k)


    col_info = []
    ddd_key = None

    for k in raw_keys:
        if k is None:
            continue
        str_k = repair_mojibake_text(ILLEGAL_CHARACTERS_RE.sub("", str(k)).strip())
        clean_k = normalize_text(str_k)
        is_bdate = is_birthdate_field(str_k)
        is_name = is_required_name_field(str_k)
        is_cpf = is_cpf_field(str_k)
        is_phone = is_required_phone_field(str_k) and not clean_k.startswith("ddd")
        if not ddd_key and (clean_k in {"ddd", "ddd_telefone", "ddd_celular", "ddd_fone", "ddd_tel"} or (clean_k.startswith("ddd") and not any(p in clean_k for p in ["cel", "tel", "fone"]))):
            ddd_key = k
        col_info.append({
            "orig_key": k,
            "str_key": str_k,
            "clean_key": clean_k,
            "is_birthdate": is_bdate,
            "is_name": is_name,
            "is_cpf": is_cpf,
            "is_phone": is_phone
        })

    cleaned_rows: List[Dict[str, Any]] = []
    for row in rows:
        cleaned_row: Dict[str, Any] = {}
        ddd_val = ""
        if ddd_key and row.get(ddd_key):
            ddd_val = re.sub(r"\D", "", str(row.get(ddd_key)))[:2]

        for info in col_info:
            orig_k = info["orig_key"]
            value = row.get(orig_k)
            str_key = info["str_key"]
            clean_key = info["clean_key"]

            if isinstance(value, (datetime, date)):
                val_cleaned = format_to_brazilian_date(value)
            elif isinstance(value, dt_time):
                val_cleaned = value.strftime("%H:%M:%S")
            elif value is None:
                val_cleaned = ""
            else:
                if info["is_birthdate"]:
                    val_cleaned = format_to_brazilian_date(value)
                elif info["is_name"]:
                    val_cleaned = normalize_name_value(value)
                elif info["is_cpf"]:
                    val_cleaned = normalize_cpf_value(value)
                elif info["is_phone"]:
                    val_cleaned = normalize_phone_value(value, ddd_val)
                else:
                    if isinstance(value, str):
                        s_val = ILLEGAL_CHARACTERS_RE.sub("", value.strip())
                        if "\r" in s_val or "\n" in s_val or "\t" in s_val or "\v" in s_val or "\f" in s_val:
                            s_val = re.sub(r"[\r\n\t\v\f]+", " ", s_val).strip()
                        if "Ã" in s_val or "â" in s_val:
                            val_cleaned = repair_mojibake_text(s_val)
                        else:
                            val_cleaned = s_val
                    else:
                        val_cleaned = value


            cleaned_row[str_key] = val_cleaned

        cleaned_rows.append(cleaned_row)
    return cleaned_rows


def build_enriched_rows(rows: List[Dict[str, Any]], detected_type: str) -> List[Dict[str, Any]]:
    enriched_rows: List[Dict[str, Any]] = []
    for row in rows:
        enriched = generate_mock_row(row, detected_type)
        enriched_rows.append(enriched)
    return enriched_rows


def first_present_value(record: Dict[str, Any], possible_keys: List[str]) -> Any:
    for key in possible_keys:
        value = get_field_value(record, key)
        if value is not None and value != "":
            return value
    return None


def normalize_api_row(row: Dict[str, Any]) -> Dict[str, Any]:
    normalized: Dict[str, Any] = {}
    aliases = {
        "nome": ["nome", "nome_completo", "nome_cliente", "cliente", "name"],
        "telefone": ["telefone", "celular", "telefone_contato", "phone", "phone_number"],
        "telefone_2": ["telefone_2", "telefone2", "celular_2", "celular2", "segundo_telefone", "outro_telefone"],
        "telefone_3": ["telefone_3", "telefone3", "celular_3", "celular3"],
        "telefone_4": ["telefone_4", "telefone4", "celular_4", "celular4"],
        "telefone_5": ["telefone_5", "telefone5", "celular_5", "celular5"],
        "cpf": ["cpf", "documento", "numero_cpf", "cpf_cliente"],
        "data_nascimento": ["data_nascimento", "dt_nascimento", "nascimento", "data_de_nascimento", "nasc"],
        "idade": ["idade", "age"],
        "sexo": ["sexo", "genero"],
        "email": ["email", "email1", "email_contato"],
        "logradouro": ["logradouro", "endereco", "endereço", "rua", "avenida"],
        "numero": ["numero", "número", "num"],
        "complemento": ["complemento", "compl"],
        "bairro": ["bairro"],
        "cidade": ["cidade", "cidade_cliente", "municipio", "município"],
        "estado": ["estado", "uf", "sigla_estado", "state"],
        "cep": ["cep"],
        "renda_estimada": ["renda_estimada", "renda", "classe_economica", "classe_social"],
        "cbo": ["cbo", "codigo_cbo", "possivel_cbo"],
        "profissao": ["profissao", "profissão", "possivel_profissao", "possivel_profissão", "cargo", "ocupacao", "ocupação"],
        "score": ["score", "score_credito", "score_faixa"],
        "perfil": ["perfil", "perfil_comercial", "profile", "personadecredito"],
        "obito": ["obito", "óbito", "flag_de_obito", "falecido"],
        "whatsapp": ["whatsapp", "wpp", "whats"],
        "status": ["status", "situacao", "situação"],
    }

    for canonical_name, keys in aliases.items():
        candidate = first_present_value(row, keys)
        if candidate is not None:
            normalized[canonical_name] = candidate

    normalized["tipo_enriquecimento"] = str(normalized.get("tipo_enriquecimento", "")).strip()
    return normalized


def extract_api_rows(payload: Any) -> List[Dict[str, Any]]:
    if isinstance(payload, list):
        candidates = payload
    elif isinstance(payload, dict):
        for key in ("data", "results", "records", "items", "rows", "enriched"):
            value = payload.get(key)
            if isinstance(value, list):
                candidates = value
                break
        else:
            candidates = [payload]
    else:
        return []

    rows: List[Dict[str, Any]] = []
    for item in candidates:
        if isinstance(item, dict):
            rows.append(normalize_api_row(item))
    return rows


def ensure_api_row_shape(row: Dict[str, Any], detected_type: str) -> Dict[str, Any]:
    safe = {"tipo_enriquecimento": detected_type}
    for key in PREFIXED_ENRICH_FIELDS:
        if key in row and row[key] is not None:
            safe[key] = row[key]
    return safe


def log_api_debug(message: str, level: str = "") -> None:
    try:
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        msg_str = str(message)

        if not level:
            lower = msg_str.lower()
            if any(k in lower for k in ["erro", "falha", "critico", "critical", "failed", "exception"]):
                level = "ERRO"
            elif any(k in lower for k in ["aviso", "warning", "alerta", "inacessivel"]):
                level = "AVISO"
            elif any(k in lower for k in ["sucesso", "concluído", "concluido", "iniciado"]):
                level = "OK"
            else:
                level = "INFO"

        client_ip = ""
        try:
            from flask import has_request_context
            if has_request_context():
                ip = get_client_ip()
                if ip:
                    client_ip = f" [IP:{ip}]"
        except Exception:
            pass

        formatted = f"[{now_str}] [{level}]{client_ip} {msg_str}\n"
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(formatted)

        # Espelha no terminal (quando rodando via `py app.py`), sem o ruido de renovacao de token
        if not any(noise in msg_str for noise in ("Renovando token", "Efetuando geracao de Token", "Token extraído", "Novo token")):
            try:
                print(formatted.rstrip("\n"), flush=True)
            except Exception:
                pass
    except Exception:
        pass


def extract_xml_tag(xml_str: str, tag_name: str) -> str:
    match = re.search(fr"<{tag_name}(\s+[^>]*)?>(.*?)</{tag_name}>", xml_str, flags=re.DOTALL | re.IGNORECASE)
    return match.group(2).strip() if match else ""


def sanitize_token(token: Any) -> str:
    """Remove caracteres de controle (CR, LF, TAB, NUL...) do token para uso seguro em headers HTTP.

    O aiohttp recente recusa montar a requisicao com 'Forbidden control character detected
    in headers' se o token vier com quebra de linha no meio.
    """
    return re.sub(r"[\x00-\x1f\x7f]", "", str(token or "")).strip()


def extract_best_phone_and_whatsapp(xml_content: str) -> tuple[str, str, List[str]]:
    """
    Extrai exclusivamente números de CELULAR (DDD1 + CEL1, DDD2 + CEL2, etc.)
    e flag de WhatsApp do XML retornado pela API.
    Telefones fixos são totalmente descartados conforme preferência estrita por celular.
    """
    if not xml_content:
        return "", "NÃO", []
    
    xml = html.unescape(xml_content)
    
    mobiles: List[tuple[str, bool]] = []
    has_whatsapp = False
    seen_phones = set()

    # 1. Prioridade Máxima: Extração explícita de pares numerados de CELULAR (DDD1 + CEL1, DDD2 + CEL2, ...)
    for i in range(1, 11):
        cel_tags = [
            f"CEL{i}", f"CELULAR{i}", f"TEL_CELULAR{i}", f"TELEFONE_CELULAR{i}",
            f"NUMERO_CELULAR{i}", f"FONE_CELULAR{i}", f"CEL_{i}", f"CELULAR_{i}",
            f"CEL_CONTATO{i}", f"CELULAR_CONTATO{i}"
        ]
        ddd_tags = [f"DDD{i}", f"DDD_CEL{i}", f"DDD_CELULAR{i}", f"DDD_{i}", f"DDD_TEL{i}"]
        whats_tags = [f"WHATSAPP{i}", f"FLWHATS{i}", f"FL_WHATS{i}", f"WHATS{i}", f"TEM_WHATSAPP{i}"]

        cel_val = ""
        for ctag in cel_tags:
            m = re.search(fr"<{ctag}\b[^>]*>\s*([^<]+)\s*</", xml, flags=re.IGNORECASE)
            if m:
                digits = re.sub(r"\D", "", m.group(1))
                if digits and not re.match(r"^0+$", digits) and len(set(digits)) > 1:
                    cel_val = digits
                    break

        ddd_val = ""
        for dtag in ddd_tags:
            m = re.search(fr"<{dtag}\b[^>]*>\s*([1-9][0-9])\s*</", xml, flags=re.IGNORECASE)
            if m:
                ddd_val = m.group(1).strip()
                break

        if not ddd_val:
            m_ddd = re.search(r"<(?:DDD|DDD_CELULAR|DDD_TEL)\b[^>]*>\s*([1-9][0-9])\s*</", xml, flags=re.IGNORECASE)
            if m_ddd:
                ddd_val = m_ddd.group(1).strip()

        is_whats = False
        for wtag in whats_tags:
            m = re.search(fr"<{wtag}\b[^>]*>\s*([^<]+)\s*</", xml, flags=re.IGNORECASE)
            if m:
                w_str = m.group(1).strip().upper()
                if w_str in ("S", "SIM", "TRUE", "1", "WHATSAPP"):
                    is_whats = True
                    has_whatsapp = True
                    break

        if cel_val:
            full_phone = ""
            if len(cel_val) == 11 and cel_val[2] == "9" and not cel_val.startswith("0"):
                full_phone = cel_val
            elif len(cel_val) == 9 and cel_val.startswith("9") and ddd_val:
                full_phone = f"{ddd_val}{cel_val}"
            elif len(cel_val) == 8 and cel_val[0] in "6789" and ddd_val:
                full_phone = f"{ddd_val}9{cel_val}"
            elif len(cel_val) == 10 and cel_val[2] in "6789":
                full_phone = f"{cel_val[:2]}9{cel_val[2:]}"

            # Aceita estritamente celulares (11 dígitos com 9 na 3ª posição)
            if full_phone and len(full_phone) == 11 and full_phone[2] == "9" and full_phone not in seen_phones:
                seen_phones.add(full_phone)
                mobiles.append((full_phone, is_whats))

    # 2. Busca em blocos estruturados de celulares (<CELULARES>, <TELEFONES_MOVEIS>, <CELULAR>, etc.)
    block_patterns = [
        r"<CELULARES\b[^>]*>(.*?)</CELULARES>",
        r"<TELEFONES_MOVEIS\b[^>]*>(.*?)</TELEFONES_MOVEIS>",
        r"<CELULAR\b[^>]*>(.*?)</CELULAR>",
        r"<TELEFONES\b[^>]*>(.*?)</TELEFONES>",
        r"<CONTATOS\b[^>]*>(.*?)</CONTATOS>",
    ]
    for pattern in block_patterns:
        blocks = re.findall(pattern, xml, flags=re.DOTALL | re.IGNORECASE)
        for block in blocks:
            ddd_match = re.search(r"<(?:DDD|DDD_CELULAR|DDD_TEL|DDD_FONE)\b[^>]*>\s*([1-9][0-9])\s*</", block, flags=re.IGNORECASE)
            cel_match = re.search(r"<(?:CELULAR|TEL_CELULAR|TELEFONE_CELULAR|NUMERO_CELULAR|CEL|FONE_CELULAR)\b[^>]*>\s*([0-9]{8,11})\s*</", block, flags=re.IGNORECASE)
            tel_match = re.search(r"<(?:TELEFONE|TELEOFNE|NUMERO|FONE|TEL)\b[^>]*>\s*([0-9]{8,11})\s*</", block, flags=re.IGNORECASE) if not cel_match else None
            whats_match = re.search(r"<(?:FLWHATS|WHATSAPP|FL_WHATS|TEM_WHATSAPP|POSSUI_WHATSAPP)\b[^>]*>\s*([^<]+)\s*</", block, flags=re.IGNORECASE)

            is_block_whats = False
            if whats_match:
                val = whats_match.group(1).strip().upper()
                if val in ("S", "SIM", "TRUE", "1"):
                    has_whatsapp = True
                    is_block_whats = True

            target_match = cel_match or tel_match
            if target_match:
                tel = target_match.group(1).strip()
                if re.match(r"^0+$", tel) or len(set(tel)) == 1:
                    continue

                full_phone = ""
                if len(tel) == 11 and tel[2] == "9" and not tel.startswith("0"):
                    full_phone = tel
                elif len(tel) == 9 and tel.startswith("9") and ddd_match:
                    full_phone = f"{ddd_match.group(1).strip()}{tel}"
                elif len(tel) == 8 and tel[0] in "6789" and ddd_match:
                    full_phone = f"{ddd_match.group(1).strip()}9{tel}"
                elif len(tel) == 10 and not tel.startswith("0") and tel[2] in "6789":
                    full_phone = f"{tel[:2]}9{tel[2:]}"

                # Apenas celulares válidos
                if full_phone and len(full_phone) == 11 and full_phone[2] == "9" and full_phone not in seen_phones:
                    seen_phones.add(full_phone)
                    mobiles.append((full_phone, is_block_whats))

    # 3. Tags diretas com números completos de 11 dígitos iniciando com 9
    if not mobiles:
        direct_full = re.findall(
            r"<(?:CELULAR|TEL_CELULAR|TELEFONE_CELULAR|NUMERO_CELULAR|CEL|TELEFONE|NUMERO)\b[^>]*>\s*([1-9][0-9]9[0-9]{8})\s*</",
            xml,
            flags=re.IGNORECASE
        )
        for full in direct_full:
            if full not in seen_phones:
                seen_phones.add(full)
                mobiles.append((full, False))

    # 4. Fallback: Telefones Fixos válidos (10 dígitos) caso nenhum celular tenha sido localizado
    if not mobiles:
        direct_fixos = re.findall(
            r"<(?:TELEFONE|FONE|TEL|TEL_FIXO|TELEFONE_FIXO|NUMERO_TELEFONE|TEL_COMERCIAL|TEL_RESIDENCIAL|NUMERO)\b[^>]*>\s*([1-9][0-9][2-5][0-9]{7})\s*</",
            xml,
            flags=re.IGNORECASE
        )
        for fixo in direct_fixos:
            if fixo not in seen_phones:
                seen_phones.add(fixo)
                mobiles.append((fixo, False))


    selected_phone = ""
    whatsapp_str = "SIM" if has_whatsapp else "NÃO"
    
    if mobiles:
        selected_phone = mobiles[0][0]
        whats_mobiles = [m[0] for m in mobiles if m[1]]
        if whats_mobiles and whats_mobiles[0] == selected_phone:
            whatsapp_str = "SIM"
        elif any(m[1] for m in mobiles):
            whatsapp_str = "SIM"

    distinct_phones = []
    if selected_phone:
        distinct_phones.append(selected_phone)
    for ph, _ in mobiles:
        if ph and ph not in distinct_phones:
            distinct_phones.append(ph)

    return selected_phone, whatsapp_str, distinct_phones


def normalize_text_for_comparison(text: str) -> str:
    if not text:
        return ""
    nfkd = unicodedata.normalize("NFKD", str(text))
    ascii_text = "".join(c for c in nfkd if not unicodedata.combining(c))
    # Remove qualquer caractere que não seja letra (incluindo números como '31', '01', pontuações, etc.)
    cleaned = re.sub(r"[^A-Za-z\s]", " ", ascii_text).upper()
    return re.sub(r"\s+", " ", cleaned).strip()


NAME_CONNECTIVES = {"DE", "DA", "DO", "DAS", "DOS", "E", "DEL", "LA", "LE", "DI", "DU"}

COMMON_AGGLUTINATED_PREFIXES = [
    ("ANAPAULA", "ANA PAULA"),
    ("ANACLARA", "ANA CLARA"),
    ("ANALUIZA", "ANA LUIZA"),
    ("ANABEATRIZ", "ANA BEATRIZ"),
    ("ANACECILIA", "ANA CECILIA"),
    ("ANACRISTINA", "ANA CRISTINA"),
    ("ANAPATRICIA", "ANA PATRICIA"),
    ("ANALUCIA", "ANA LUCIA"),
    ("ANAMARIA", "ANA MARIA"),
    ("JOAOPAULO", "JOAO PAULO"),
    ("JOAOPEDRO", "JOAO PEDRO"),
    ("JOAOVITOR", "JOAO VITOR"),
    ("JOAOGABRIEL", "JOAO GABRIEL"),
    ("JOAOMARCOS", "JOAO MARCOS"),
    ("MARIACLARA", "MARIA CLARA"),
    ("MARIAEDUARDA", "MARIA EDUARDA"),
    ("MARIALUIZA", "MARIA LUIZA"),
    ("MARIAALICE", "MARIA ALICE"),
    ("MARIAHELENA", "MARIA HELENA"),
    ("MARIACECILIA", "MARIA CECILIA"),
    ("MARIAAPARECIDA", "MARIA APARECIDA"),
    ("MARIAFERNANDA", "MARIA FERNANDA"),
    ("MARIAJULIA", "MARIA JULIA"),
    ("MARIACRISTINA", "MARIA CRISTINA"),
    ("LUIZCARLOS", "LUIZ CARLOS"),
    ("LUISCARLOS", "LUIS CARLOS"),
    ("LUIZOTAVIO", "LUIZ OTAVIO"),
    ("LUIZEDUARDO", "LUIZ EDUARDO"),
    ("LUIZFERNANDO", "LUIZ FERNANDO"),
    ("LUIZANTONIO", "LUIZ ANTONIO"),
    ("CARLOSALBERTO", "CARLOS ALBERTO"),
    ("CARLOSHENRIQUE", "CARLOS HENRIQUE"),
    ("CARLOSAUGUSTO", "CARLOS AUGUSTO"),
    ("CARLOSEDUARDO", "CARLOS EDUARDO"),
    ("CARLOSALEXANDRE", "CARLOS ALEXANDRE"),
    ("PEDROHENRIQUE", "PEDRO HENRIQUE"),
    ("PEDROPAULO", "PEDRO PAULO"),
    ("PEDROAUGUSTO", "PEDRO AUGUSTO"),
    ("PEDROVITOR", "PEDRO VITOR"),
    ("VITORHUGO", "VITOR HUGO"),
    ("PAULOFERNANDO", "PAULO FERNANDO"),
    ("PAULOCESAR", "PAULO CESAR"),
    ("PAULOROBERTO", "PAULO ROBERTO"),
    ("PAULOEDUARDO", "PAULO EDUARDO"),
    ("PAULOHENRIQUE", "PAULO HENRIQUE"),
    ("MARCOSANTONIO", "MARCOS ANTONIO"),
    ("MARCOSAURELIO", "MARCOS AURELIO"),
    ("MARCOSVINICIUS", "MARCOS VINICIUS"),
    ("JOSECARLOS", "JOSE CARLOS"),
    ("JOSEANTONIO", "JOSE ANTONIO"),
    ("JOSEEDUARDO", "JOSE EDUARDO"),
    ("JOSEFERNANDO", "JOSE FERNANDO"),
    ("JOSEROBERTO", "JOSE ROBERTO"),
]


def generate_name_query_variants(name_val: str) -> List[str]:
    """Gera variações de busca para nomes com aglutinações comuns (ex: 'ANAPAULA DA SILVA' -> 'ANA PAULA DA SILVA')."""
    if not name_val:
        return []
    variants = [name_val]
    upper_name = name_val.upper()
    tokens = upper_name.split()
    if not tokens:
        return variants

    first_word = tokens[0]
    # 1. Verifica tabela de aglutinações diretas
    for agg, sep in COMMON_AGGLUTINATED_PREFIXES:
        if first_word == agg:
            new_first = sep
            variant = " ".join([new_first] + tokens[1:])
            if variant not in variants:
                variants.append(variant)
            break

    # 2. Heurística para prefixos conhecidos (ANA, MARIA, JOAO, LUIZ, LUIS, JOSE, CARLOS, PEDRO, PAULO, VITOR, MARCOS)
    known_prefixes = ["ANA", "MARIA", "JOAO", "LUIZ", "LUIS", "JOSE", "CARLOS", "PEDRO", "PAULO", "VITOR", "MARCOS"]
    for pfx in known_prefixes:
        if first_word.startswith(pfx) and len(first_word) >= len(pfx) + 3:
            rem = first_word[len(pfx):]
            variant = " ".join([pfx, rem] + tokens[1:])
            if variant not in variants:
                variants.append(variant)
            break

    return variants


def levenshtein_similarity(s1: str, s2: str) -> float:
    if s1 == s2:
        return 1.0
    if not s1 or not s2:
        return 0.0
    len1, len2 = len(s1), len(s2)
    matrix = [[0] * (len2 + 1) for _ in range(len1 + 1)]
    for i in range(len1 + 1):
        matrix[i][0] = i
    for j in range(len2 + 1):
        matrix[0][j] = j
    for i in range(1, len1 + 1):
        for j in range(1, len2 + 1):
            cost = 0 if s1[i - 1] == s2[j - 1] else 1
            matrix[i][j] = min(
                matrix[i - 1][j] + 1,
                matrix[i][j - 1] + 1,
                matrix[i - 1][j - 1] + cost
            )
    dist = matrix[len1][len2]
    max_len = max(len1, len2)
    base_sim = max(0.0, 1.0 - (dist / max_len))
    # Se a diferença for de apenas 1 caractere em nomes com pelo menos 4 letras (ex: LUIZ vs LUIS)
    if dist == 1 and min(len1, len2) >= 4:
        return max(base_sim, 0.88)
    return base_sim


def calculate_name_similarity(name_input: str, name_api: str) -> Tuple[float, str, str]:
    """
    Calcula a similaridade semântica entre o nome fornecido na planilha e o nome retornado pela API.
    Trata com precisão:
      - Aglutinação de nomes compostos (ex: 'Anapaula' vs 'Ana Paula', 'Joaopaulo' vs 'Joao Paulo').
      - Números e códigos agregados (ex: 'Eliana Silva31' -> 'Eliana Silva').
      - Subconjunto e nomes intermediários (ex: 'Eliana Silva' vs 'Eliana Bentes da Silva Dias').
      - Abreviações e variações fonéticas.
    """
    norm_in = normalize_text_for_comparison(name_input)
    norm_api = normalize_text_for_comparison(name_api)

    if not norm_in:
        return 100.0, "NAO_APLICAVEL", "Origem sem nome (consulta direta)"

    if not norm_api:
        return 0.0, "DIVERGENTE", "API sem nome retornado"

    if norm_in == norm_api:
        return 100.0, "ALTA", "Nome idêntico"

    tokens_in = [t for t in norm_in.split() if t not in NAME_CONNECTIVES]
    tokens_api = [t for t in norm_api.split() if t not in NAME_CONNECTIVES]

    if not tokens_in or not tokens_api:
        return 0.0, "DIVERGENTE", "Nome inválido ou apenas conectivos"

    # 1. Comparação do Primeiro Nome (âncora principal)
    first_in = tokens_in[0]
    first_api = tokens_api[0]
    first_sim = levenshtein_similarity(first_in, first_api)
    first_match = (first_in == first_api) or (first_sim >= 0.82)
    first_abbrev = (len(first_in) == 1 and first_api.startswith(first_in)) or (len(first_api) == 1 and first_in.startswith(first_api))

    # Tratamento de aglutinação de primeiro nome (ex: ANAPAULA vs ANA PAULA ou vice-versa)
    if not first_match and not first_abbrev:
        if len(tokens_api) >= 2 and first_in == tokens_api[0] + tokens_api[1]:
            first_match = True
        elif len(tokens_in) >= 2 and tokens_in[0] + tokens_in[1] == first_api:
            first_match = True
        elif len(tokens_api) >= 2 and first_in.startswith(tokens_api[0]) and len(tokens_api[0]) >= 3:
            first_match = True

    if not first_match and not first_abbrev:
        if first_in in tokens_api:
            first_match = True
        else:
            common_tokens = set(tokens_in).intersection(set(tokens_api))
            if not common_tokens:
                return 0.0, "DIVERGENTE", "Titular divergente (nenhum token em comum)"
            else:
                return 25.0, "DIVERGENTE", f"Primeiro nome divergente ({first_in} vs {first_api})"

    # 2. Token Matching Inteligente (com suporte a concatenação de compostos, ordem, subconjuntos e abreviações)
    matched_in_indices = set()
    matched_api_indices = set()
    total_token_scores = 0.0

    for i, t_in in enumerate(tokens_in):
        best_sim = 0.0
        best_j = -1
        best_j2 = -1

        for j, t_api in enumerate(tokens_api):
            if j in matched_api_indices:
                continue
            if t_in == t_api:
                sim = 1.0
            # Caso Aglutinado: 'ANAPAULA' -> 'ANA' + 'PAULA'
            elif j + 1 < len(tokens_api) and (j + 1) not in matched_api_indices and t_in == t_api + tokens_api[j + 1]:
                sim = 1.0
                best_j2 = j + 1
            # Caso Inverso: 'ANA' + 'PAULA' -> 'ANAPAULA'
            elif i + 1 < len(tokens_in) and tokens_in[i] + tokens_in[i + 1] == t_api:
                sim = 1.0
            elif (len(t_in) == 1 and t_api.startswith(t_in)) or (len(t_api) == 1 and t_in.startswith(t_api)):
                sim = 0.95  # Match de abreviação
            else:
                sim = levenshtein_similarity(t_in, t_api)
                if sim < 0.80:
                    sim = 0.0
            
            if sim > best_sim:
                best_sim = sim
                best_j = j
        
        if best_sim > 0.0 and best_j != -1:
            matched_in_indices.add(i)
            matched_api_indices.add(best_j)
            if best_j2 != -1:
                matched_api_indices.add(best_j2)
            total_token_scores += best_sim

    matched_ratio_in = total_token_scores / len(tokens_in)
    
    # 3. Caso Subconjunto Exato (ex: "Paulo Silva" -> "Paulo Isac Saverio Silva" ou "Eliana Silva" -> "Eliana Bentes da Silva Dias")
    if len(matched_in_indices) == len(tokens_in) and matched_ratio_in >= 0.95:
        last_in = tokens_in[-1]
        last_found = any(tokens_api[j] == last_in or levenshtein_similarity(tokens_api[j], last_in) >= 0.85 for j in matched_api_indices)
        if last_found:
            return 100.0, "ALTA", "Subconjunto exato (nomes intermediários adicionados)"
        else:
            return 95.0, "ALTA", "Subconjunto de tokens compatível"

    # 4. Caso de Nomes com Abreviações ou Variações Ortográficas Leves
    if matched_ratio_in >= 0.80:
        return round(matched_ratio_in * 100, 1), "ALTA", "Alta similaridade com variação leve"

    if matched_ratio_in >= 0.50:
        return round(matched_ratio_in * 100, 1), "MEDIA", "Similaridade parcial (sobrenomes divergentes/faltantes)"

    return round(matched_ratio_in * 100, 1), "BAIXA", "Baixa correspondência"


def extract_best_candidate_from_xml(
    xml_content: str, target_name: str = "", target_uf: str = ""
) -> Tuple[str, str, str, float, str]:
    """
    Analisa o XML retornado pela API (podendo conter 1 ou múltiplos registros candidatos)
    e seleciona o candidato com maior pontuação de similaridade semântica com o target_name e target_uf.
    Retorna: (cpf_selecionado, nome_selecionado, xml_bloco, score_similaridade, nivel_confianca)
    """
    if not xml_content:
        return "", "", "", 0.0, "DIVERGENTE"

    xml = html.unescape(xml_content)
    if "Nada Consta" in xml and "<CADASTRAIS>" not in xml and "<DOCUMENTO>" not in xml and "<AVANCADAS>" not in xml and "<NOME>" not in xml:
        return "", "", "", 0.0, "DIVERGENTE"

    # 1. Procura blocos de registros múltiplos (<REGISTRO>, <PESSOA>, <ITEM>, <CADASTRAIS>, <CLIENTE>, etc.)
    record_patterns = [
        r"<REGISTRO\b[^>]*>(.*?)</REGISTRO>",
        r"<PESSOA\b[^>]*>(.*?)</PESSOA>",
        r"<ITEM\b[^>]*>(.*?)</ITEM>",
        r"<CLIENTE\b[^>]*>(.*?)</CLIENTE>",
        r"<DADOSCADASTRAIS\b[^>]*>(.*?)</DADOSCADASTRAIS>",
        r"<CADASTRAIS\b[^>]*>(.*?)</CADASTRAIS>",
        r"<WS_PESSOA\b[^>]*>(.*?)</WS_PESSOA>",
        r"<SOCIO\b[^>]*>(.*?)</SOCIO>",
    ]

    candidates: List[Tuple[str, str, str, str]] = []  # (cpf, nome, uf, xml_block)
    for pat in record_patterns:
        matches = re.findall(pat, xml, flags=re.DOTALL | re.IGNORECASE)
        if matches:
            for block in matches:
                m_cpf = re.search(r"<(?:CPF|DOCUMENTO|NUM_CPF|DOC)\b[^>]*>\s*([^<]+)\s*</", block, flags=re.IGNORECASE)
                m_nome = re.search(r"<(?:NOME|NOME_PESSOA|RAZAO_SOCIAL)\b[^>]*>\s*([^<]+)\s*</", block, flags=re.IGNORECASE)
                m_uf = re.search(r"<(?:UF|ESTADO|ESTADO_SIGLA)\b[^>]*>\s*([^<]+)\s*</", block, flags=re.IGNORECASE)
                cpf = re.sub(r"\D", "", m_cpf.group(1)) if m_cpf else ""
                nome = m_nome.group(1).strip() if m_nome else ""
                uf = m_uf.group(1).strip().upper() if m_uf else ""
                if cpf or nome:
                    candidates.append((cpf, nome, uf, block))
            if candidates:
                break

    # Suporte a respostas AVANCADAS com blocos repetidos de TIPO
    if not candidates and ("<AVANCADAS>" in xml or "<AVANÇADAS>" in xml):
        tipo_blocks = re.split(r"<TIPO\b[^>]*>", xml, flags=re.IGNORECASE)
        for block in tipo_blocks[1:]:
            m_cpf = re.search(r"<(?:CPF|DOCUMENTO|NUM_CPF|DOC)\b[^>]*>\s*([^<]+)\s*</", block, flags=re.IGNORECASE)
            m_nome = re.search(r"<(?:NOME|NOME_PESSOA|RAZAO_SOCIAL)\b[^>]*>\s*([^<]+)\s*</", block, flags=re.IGNORECASE)
            m_uf = re.search(r"<(?:UF|ESTADO|ESTADO_SIGLA)\b[^>]*>\s*([^<]+)\s*</", block, flags=re.IGNORECASE)
            cpf = re.sub(r"\D", "", m_cpf.group(1)) if m_cpf else ""
            nome = m_nome.group(1).strip() if m_nome else ""
            uf = m_uf.group(1).strip().upper() if m_uf else ""
            if cpf or nome:
                candidates.append((cpf, nome, uf, block))

    # Se não encontrou blocos filhos específicos, avalia o XML raiz como candidato único
    if not candidates:
        m_cpf = re.search(r"<(?:CPF|DOCUMENTO|NUM_CPF|DOC)\b[^>]*>\s*([^<]+)\s*</", xml, flags=re.IGNORECASE)
        m_nome = re.search(r"<(?:NOME|NOME_PESSOA|RAZAO_SOCIAL)\b[^>]*>\s*([^<]+)\s*</", xml, flags=re.IGNORECASE)
        m_uf = re.search(r"<(?:UF|ESTADO|ESTADO_SIGLA)\b[^>]*>\s*([^<]+)\s*</", xml, flags=re.IGNORECASE)
        cpf = re.sub(r"\D", "", m_cpf.group(1)) if m_cpf else ""
        nome = m_nome.group(1).strip() if m_nome else ""
        uf = m_uf.group(1).strip().upper() if m_uf else ""
    # Se houver múltiplos candidatos: filtra pela UF solicitada
    if target_uf:
        req_uf = normalize_uf_string(target_uf)
        uf_candidates = [c for c in candidates if normalize_uf_string(c[2]) == req_uf]
        if len(uf_candidates) == 1:
            candidates = uf_candidates
        elif len(uf_candidates) > 1:
            # 2 ou mais pessoas com o mesmo nome na mesma UF -> FALHA (homônimo)
            return "", "", "", 0.0, "DIVERGENTE"
        elif not uf_candidates and candidates:
            # Candidatos encontrados são de outras UFs -> FALHA
            return "", "", "", 0.0, "DIVERGENTE"
    else:
        distinct_persons = list({(c[0], normalize_name_value(c[1]).upper()) for c in candidates if c[0] or c[1]})
        if len(distinct_persons) > 1:
            return "", "", "", 0.0, "DIVERGENTE"

    if not candidates:
        return "", "", "", 0.0, "DIVERGENTE"

    if not target_name:
        first_cpf, first_nome, first_uf, first_block = candidates[0]
        return first_cpf, first_nome, first_block, 100.0, "NAO_APLICAVEL"

    # Avalia a similaridade do candidato selecionado com o nome alvo
    best_cand = candidates[0]
    cand_cpf, cand_nome, cand_uf, cand_block = best_cand
    score, level, reason = calculate_name_similarity(target_name, cand_nome)
    return cand_cpf, cand_nome, cand_block, score, level


def parse_nova_vida_full_xml(xml_str: str) -> Dict[str, str]:
    if not xml_str:
        return {}
    
    clean_str = xml_str.strip()
    # 1. Se for JSON retornado pelo endpoint /PESQUISAR (consultas 2421 / 2420 / 2145)
    if clean_str.startswith("{") or clean_str.startswith("[") or '"CADASTRAIS"' in clean_str or '"CPF"' in clean_str or '"NOME"' in clean_str:
        try:
            decoder = json.JSONDecoder()
            obj, _ = decoder.raw_decode(clean_str)
            
            # Se for uma lista de resultados, pega o primeiro item
            if isinstance(obj, list) and obj:
                obj = obj[0]

            if isinstance(obj, dict):
                cad_raw = obj.get("CADASTRAIS") or obj.get("DADOSCADASTRAIS") or obj.get("PESSOA") or obj.get("CLIENTE") or obj
                if isinstance(cad_raw, list) and cad_raw:
                    cad = cad_raw[0] if isinstance(cad_raw[0], dict) else {}
                elif isinstance(cad_raw, dict):
                    cad = cad_raw
                else:
                    cad = {}

                cpf = re.sub(r"\D", "", str(cad.get("CPF") or cad.get("DOCUMENTO") or obj.get("CPF") or obj.get("DOCUMENTO") or ""))
                nome = cad.get("NOME") or obj.get("NOME") or ""
                sexo = cad.get("SEXO") or obj.get("SEXO") or ""
                nasc = format_to_brazilian_date(cad.get("NASCIMENTO") or cad.get("NASC") or obj.get("NASCIMENTO") or obj.get("NASC") or "")
                idade = str(cad.get("IDADE") or obj.get("IDADE") or "")
                if not idade and nasc:
                    ym = re.search(r"\b\d{4}\b", nasc)
                    if ym:
                        try:
                            idade = str(datetime.now().year - int(ym.group(0)))
                        except Exception:
                            idade = ""

                renda = normalize_renda_estimada(str(cad.get("RENDA") or cad.get("RENDA_FAMILIAR") or obj.get("RENDA") or ""))
                rg = cad.get("RG") or obj.get("RG") or ""

                # Endereços: suporta dict único "ENDERECO" (2421) ou lista "ENDERECOS"
                end_raw = obj.get("ENDERECO") or obj.get("ENDERECOS") or cad.get("ENDERECO") or cad.get("ENDERECOS") or {}
                if isinstance(end_raw, list) and end_raw:
                    end0 = end_raw[0] if isinstance(end_raw[0], dict) else {}
                elif isinstance(end_raw, dict):
                    end0 = end_raw
                else:
                    end0 = {}

                def _clean_end_field(val: Any) -> str:
                    s = str(val or "").strip()
                    return "" if s.upper() in ("NULL", "NONE") else s

                logradouro = _clean_end_field(end0.get("LOGRADOURO") or obj.get("LOGRADOURO") or cad.get("LOGRADOURO"))
                numero = _clean_end_field(end0.get("NUMERO") or obj.get("NUMERO") or cad.get("NUMERO"))
                complemento = _clean_end_field(end0.get("COMPLEMENTO") or obj.get("COMPLEMENTO") or cad.get("COMPLEMENTO"))
                bairro = _clean_end_field(end0.get("BAIRRO") or obj.get("BAIRRO") or cad.get("BAIRRO"))
                cidade = _clean_end_field(end0.get("CIDADE") or obj.get("CIDADE") or cad.get("CIDADE"))
                uf = _clean_end_field(end0.get("UF") or obj.get("UF") or cad.get("UF"))
                cep = _clean_end_field(end0.get("CEP") or obj.get("CEP") or cad.get("CEP"))

                # Contatos e Telefones: suporta lista "CONTATOS" (2421) ou "TELEFONES" / "CELULARES"
                telefones = obj.get("CONTATOS") or obj.get("TELEFONES") or cad.get("CONTATOS") or cad.get("TELEFONES") or obj.get("CELULARES") or []
                distinct_phones = []
                phone_val = ""
                has_whatsapp = False
                if isinstance(telefones, list):
                    for t_item in telefones:
                        if isinstance(t_item, dict):
                            ddd_t = str(t_item.get("DDD") or "").strip()
                            num_t = str(t_item.get("TELEFONE") or t_item.get("CELULAR") or t_item.get("NUMERO") or "").strip()
                            w_flag = str(t_item.get("WHATSAPP") or t_item.get("FLWHATS") or "").upper()
                            if w_flag in ("S", "SIM", "TRUE", "1"):
                                has_whatsapp = True
                            full_t = normalize_phone_value(num_t, ddd_t)
                            if full_t and full_t not in distinct_phones:
                                distinct_phones.append(full_t)
                        elif isinstance(t_item, str):
                            full_t = normalize_phone_value(t_item)
                            if full_t and full_t not in distinct_phones:
                                distinct_phones.append(full_t)
                elif isinstance(telefones, dict):
                    ddd_t = str(telefones.get("DDD") or "").strip()
                    num_t = str(telefones.get("TELEFONE") or telefones.get("CELULAR") or telefones.get("NUMERO") or "").strip()
                    full_t = normalize_phone_value(num_t, ddd_t)
                    if full_t:
                        distinct_phones.append(full_t)

                if distinct_phones:
                    phone_val = distinct_phones[0]

                emails = obj.get("EMAILS") or cad.get("EMAILS") or []
                email_val = ""
                if isinstance(emails, list) and emails and isinstance(emails[0], dict):
                    email_val = normalize_email_value(emails[0].get("EMAIL", ""))
                elif isinstance(emails, str):
                    email_val = normalize_email_value(emails)

                score = str(cad.get("SCORE") or obj.get("SCORE") or "")
                perfil = str(cad.get("PERFIL") or obj.get("PERFIL") or "")
                cbo = str(cad.get("CBO") or obj.get("CBO") or "")
                profissao = str(cad.get("POSSIVEL_PROFISSAO") or cad.get("PROFISSAO") or obj.get("POSSIVEL_PROFISSAO") or obj.get("PROFISSAO") or "")
                status = str(cad.get("STATUS") or cad.get("DESCRICAO") or obj.get("STATUS") or "REGULAR")

                return {
                    "cpf": cpf,
                    "nome": nome,
                    "sexo": sexo,
                    "nascimento": nasc,
                    "idade": idade,
                    "telefone": phone_val,
                    "telefone_2": distinct_phones[1] if len(distinct_phones) > 1 else "",
                    "telefone_3": distinct_phones[2] if len(distinct_phones) > 2 else "",
                    "telefone_4": distinct_phones[3] if len(distinct_phones) > 3 else "",
                    "telefone_5": distinct_phones[4] if len(distinct_phones) > 4 else "",
                    "whatsapp": "SIM" if (has_whatsapp or phone_val) else "NÃO",
                    "email": email_val,
                    "logradouro": logradouro,
                    "numero": numero,
                    "complemento": complemento,
                    "bairro": bairro,
                    "cidade": cidade,
                    "estado": uf,
                    "cep": cep,
                    "renda_estimada": renda,
                    "score": score,
                    "perfil": perfil,
                    "cbo": cbo,
                    "profissao": profissao,
                    "obito": "0",
                    "status": status,
                }
        except Exception:
            pass

    # 2. Resposta em formato XML (NVCHECK / NVCHECK_TELEFONE / PesquisaAtributos)
    xml = html.unescape(xml_str)
    if "Nada Consta" in xml and "<CADASTRAIS>" not in xml and "<DOCUMENTO>" not in xml and "<AVANCADAS>" not in xml:
        return {}

    def get_tag(tag: str, text: str = xml) -> str:
        m = re.search(fr"<{tag}(?:\s+[^>]*)?>(.*?)</{tag}>", text, flags=re.DOTALL | re.IGNORECASE)
        return m.group(1).strip() if m else ""

    # Cadastrais
    cpf = get_tag("CPF") or get_tag("DOCUMENTO")
    cpf = re.sub(r"\D", "", cpf)
    nome = get_tag("NOME")
    sexo = get_tag("SEXO")
    nasc = get_tag("NASC") or get_tag("NASCIMENTO") or get_tag("DATA_NASCIMENTO")
    nasc = format_to_brazilian_date(nasc)
    idade = get_tag("IDADE")
    if not idade and nasc:
        ym = re.search(r"\b\d{4}\b", nasc)
        if ym:
            try:
                idade = str(datetime.now().year - int(ym.group(0)))
            except Exception:
                idade = ""

    # Situacao / Status
    status = get_tag("DESCRICAO") or "REGULAR"

    # Renda / Score / Persona / CBO / Profissao
    renda = normalize_renda_estimada(get_tag("RENDA") or get_tag("RENDA_ESTIMADA") or get_tag("CLASSEECONOMICA"))
    score = get_tag("SCORE") or get_tag("SCORE_VALOR") or get_tag("SCORE_PONTUACAO")
    perfil = get_tag("PERSONACREDITO") or get_tag("PERFIL") or get_tag("PERSONADECREDITO")
    cbo = get_tag("CODIGO_CBO") or get_tag("CBO") or get_tag("POSSIVEL_CBO")
    profissao = get_tag("POSSIVELPROFISSAO") or get_tag("PROFISSAO") or get_tag("POSSIVEL_PROFISSAO")

    # Obito
    flobito = get_tag("FLOBITO") or get_tag("OBITO") or get_tag("FLAG_DE_OBITO")
    obito = "1" if (flobito.upper() in ("S", "SIM", "1", "TRUE") or "FALECIDO" in status.upper()) else "0"

    # Enderecos
    logradouro = get_tag("LOGRADOURO")
    numero = get_tag("NUMERO")
    complemento = get_tag("COMPLEMENTO")
    bairro = get_tag("BAIRRO")
    cidade = get_tag("CIDADE")
    uf = get_tag("UF")
    cep = get_tag("CEP")

    # Telefones & WhatsApp estritamente validados
    phone_val, whatsapp, distinct_phones = extract_best_phone_and_whatsapp(xml)

    # Email
    email = normalize_email_value(get_tag("EMAIL"))

    return {
        "cpf": cpf,
        "nome": nome,
        "sexo": sexo,
        "nascimento": nasc,
        "idade": idade,
        "telefone": phone_val,
        "telefone_2": distinct_phones[1] if len(distinct_phones) > 1 else "",
        "telefone_3": distinct_phones[2] if len(distinct_phones) > 2 else "",
        "telefone_4": distinct_phones[3] if len(distinct_phones) > 3 else "",
        "telefone_5": distinct_phones[4] if len(distinct_phones) > 4 else "",
        "whatsapp": whatsapp,
        "email": email,
        "logradouro": logradouro,
        "numero": numero,
        "complemento": complemento,
        "bairro": bairro,
        "cidade": cidade,
        "estado": uf,
        "cep": cep,
        "renda_estimada": renda,
        "score": score,
        "perfil": perfil,
        "cbo": cbo,
        "profissao": profissao,
        "obito": obito,
        "status": status,
    }


def parse_and_validate_xml_for_target(
    xml_resp: str,
    target_name: str = "",
    target_uf: str = "",
    is_cpf_query: bool = False
) -> Tuple[bool, str, str, float, str, Dict[str, Any]]:
    """
    Analisa e valida se o XML ou JSON retornado pela API é compatível com os critérios da consulta:
    1. Se for consulta por CPF (is_cpf_query=True), aceita 100% do retorno sem bloquear por nome.
    2. Se for consulta por Nome/UF ou Nome/Telefone, valida similaridade semântica e UF.
    """
    if not xml_resp:
        return False, "", "", 0.0, "DIVERGENTE", {}

    xml_clean = html.unescape(xml_resp).strip()
    if "Nada Consta" in xml_clean and "<CADASTRAIS>" not in xml_clean and "<DOCUMENTO>" not in xml_clean and "<AVANCADAS>" not in xml_clean and "<NOME>" not in xml_clean and '"CADASTRAIS"' not in xml_clean:
        return False, "", "", 0.0, "DIVERGENTE", {}

    # Se a resposta for JSON (PESQUISAR 2145)
    if xml_clean.startswith("{") or xml_clean.startswith("[") or '"CADASTRAIS"' in xml_clean or '"CPF"' in xml_clean or '"NOME"' in xml_clean:
        try:
            decoder = json.JSONDecoder()
            obj, _ = decoder.raw_decode(xml_clean)
            
            # Se for uma lista de candidatos retornada
            if isinstance(obj, list):
                if target_uf:
                    req_uf = normalize_uf_string(target_uf)
                    filtered = []
                    for item in obj:
                        item_str = json.dumps(item) if isinstance(item, dict) else str(item)
                        p = parse_nova_vida_full_xml(item_str)
                        if normalize_uf_string(p.get("estado", "")) == req_uf:
                            filtered.append(item)
                    if len(filtered) == 1:
                        obj = filtered[0]
                        xml_clean = json.dumps(obj)
                    elif len(filtered) > 1:
                        # 2 ou mais pessoas com o mesmo nome na mesma UF -> FALHA
                        return False, "", "", 0.0, "DIVERGENTE", {}
                    else:
                        # Candidatos encontrados são de outras UFs -> FALHA
                        return False, "", "", 0.0, "DIVERGENTE", {}
                else:
                    if len(obj) > 1:
                        return False, "", "", 0.0, "DIVERGENTE", {}
                    elif len(obj) == 1:
                        obj = obj[0]
                        xml_clean = json.dumps(obj)
            elif isinstance(obj, dict):
                cad_raw = obj.get("CADASTRAIS") or obj.get("DADOSCADASTRAIS") or obj.get("PESSOA") or obj.get("CLIENTE")
                if isinstance(cad_raw, list):
                    if target_uf:
                        req_uf = normalize_uf_string(target_uf)
                        filtered = []
                        for item in cad_raw:
                            item_str = json.dumps(item) if isinstance(item, dict) else str(item)
                            p = parse_nova_vida_full_xml(item_str)
                            if normalize_uf_string(p.get("estado", "")) == req_uf:
                                filtered.append(item)
                        if len(filtered) == 1:
                            obj["CADASTRAIS"] = filtered[0]
                            xml_clean = json.dumps(obj)
                        elif len(filtered) > 1:
                            return False, "", "", 0.0, "DIVERGENTE", {}
                        else:
                            return False, "", "", 0.0, "DIVERGENTE", {}
                    else:
                        if len(cad_raw) > 1:
                            return False, "", "", 0.0, "DIVERGENTE", {}
        except Exception:
            pass

        parsed = parse_nova_vida_full_xml(xml_clean)
        cand_cpf = parsed.get("cpf", "")
        cand_nome = parsed.get("nome", "")

        if is_cpf_query:
            if cand_cpf or cand_nome or parsed.get("telefone"):
                return True, cand_cpf, cand_nome, 100.0, "NAO_APLICAVEL", parsed
            return False, "", "", 0.0, "DIVERGENTE", {}

        if target_name:
            if not cand_nome:
                if cand_cpf:
                    return True, cand_cpf, target_name, 100.0, "ALTA", parsed
                return False, cand_cpf, cand_nome, 0.0, "DIVERGENTE", {}

            score, level, reason = calculate_name_similarity(target_name, cand_nome)
            if level == "DIVERGENTE" or score < 45.0:
                if cand_cpf and (parsed.get("nascimento") or parsed.get("logradouro") or parsed.get("telefone")):
                    return True, cand_cpf, cand_nome, max(score, 60.0), "MEDIA", parsed
                return False, cand_cpf, cand_nome, score, "DIVERGENTE", {}
            return True, cand_cpf, cand_nome, score, level, parsed
        return True, cand_cpf, cand_nome, 100.0, "NAO_APLICAVEL", parsed

    cpf_extracted, best_cand_nome, cand_block, sim_score, conf_level = extract_best_candidate_from_xml(
        xml_clean, target_name=target_name if not is_cpf_query else "", target_uf=target_uf
    )

    target_block = cand_block if cand_block else xml_clean

    if not cpf_extracted:
        cpf_extracted = extract_xml_tag(target_block, "CPF") or extract_xml_tag(target_block, "DOCUMENTO")
        cpf_extracted = re.sub(r"\D", "", cpf_extracted)

    # 1. Quando enriquecido com CPF: aceita 100% do retorno caso contenha CPF, Nome ou dados cadastrais
    if is_cpf_query:
        parsed = parse_nova_vida_full_xml(target_block) or {}
        if cpf_extracted and not parsed.get("cpf"):
            parsed["cpf"] = cpf_extracted
        if best_cand_nome and not parsed.get("nome"):
            parsed["nome"] = best_cand_nome

        has_any_data = bool(
            parsed.get("cpf") or parsed.get("nome") or parsed.get("telefone") or 
            parsed.get("nascimento") or parsed.get("logradouro") or parsed.get("cidade")
        )
        if has_any_data:
            return True, parsed.get("cpf") or cpf_extracted, parsed.get("nome") or best_cand_nome, 100.0, "NAO_APLICAVEL", parsed
        return False, "", "", 0.0, "DIVERGENTE", {}

    # 2. Se o usuário solicitou um nome e o retorno é DIVERGENTE (< 45%), rejeita imediatamente este XML
    if target_name and (conf_level == "DIVERGENTE" or sim_score < 45.0):
        return False, cpf_extracted, best_cand_nome, sim_score, "DIVERGENTE", {}

    parsed = parse_nova_vida_full_xml(target_block) or {}
    if not parsed.get("cpf") and cpf_extracted:
        parsed["cpf"] = cpf_extracted
    if not parsed.get("nome") and best_cand_nome:
        parsed["nome"] = best_cand_nome

    has_any_data = bool(
        parsed.get("cpf") or parsed.get("nome") or parsed.get("telefone") or 
        parsed.get("nascimento") or parsed.get("logradouro") or parsed.get("cidade")
    )
    if has_any_data:
        return True, parsed.get("cpf") or cpf_extracted, parsed.get("nome") or best_cand_nome, sim_score, conf_level, parsed

    return False, "", "", 0.0, "DIVERGENTE", {}



_CACHED_NOVA_VIDA_TOKEN = None
_CACHED_NOVA_VIDA_TOKEN_TIME = 0.0
_NOVA_VIDA_TOKEN_LOCK = threading.Lock()

def is_token_expired_response(text: str) -> bool:
    if not text:
        return False
    raw_upper = text.upper()
    # Termos explícitos de erro de autenticação / token / sessão / bloqueio
    if any(err in raw_upper for err in [
        "TOKEN EXPIRADO", "TOKEN INVALIDO", "TOKEN INVÁLIDO", "SESSAO EXPIRADA", 
        "SESSÃO EXPIRADA", "ACESSO NEGADO", "NAO AUTORIZADO", "NÃO AUTORIZADO",
        "TOKEN NAO INFORMADO", "TOKEN NÃO INFORMADO", "TOKEN NAO ENCONTRADO",
        "TOKEN NÃO ENCONTRADO", "FALHA NA AUTENTICACAO", "FALHA NA AUTENTICAÇÃO",
        "AUTENTICACAO INVALIDA", "AUTENTICAÇÃO INVÁLIDA", "BLOQUEIO TEMPORARIO",
        "BLOQUEIO TEMPORÁRIO", "LIMITE DE CONEXOES", "LIMITE DE CONEXÕES",
        "ERRO AO AUTENTICAR", "SESSAO INVALIDA", "SESSÃO INVÁLIDA"
    ]):
        return True

    unescaped = html.unescape(text)
    unescaped_upper = unescaped.upper()
    if any(tag in unescaped_upper for tag in [
        "<CONSULTA", "<CADASTRAIS", "<AVANCADAS", "<AVANÇADAS", 
        "<DOCUMENTO", "<NOME", "<REGISTRO", "<PESSOA", "<ITEM", 
        "<CLIENTE", "NADA CONSTA"
    ]):
        return False

    if "<STRING" in raw_upper:
        return True

    return False

def get_nova_vida_token(force_refresh: bool = False) -> str | None:
    global _CACHED_NOVA_VIDA_TOKEN, _CACHED_NOVA_VIDA_TOKEN_TIME
    with _NOVA_VIDA_TOKEN_LOCK:
        if NOVA_VIDA_API_KEY:
            log_api_debug("Usando NOVA_VIDA_API_KEY")
            return sanitize_token(NOVA_VIDA_API_KEY)

        now = time.time()
        # Se um token foi gerado há menos de 5 segundos, reutiliza para evitar avalanche de requisições simultâneas
        if _CACHED_NOVA_VIDA_TOKEN and (now - _CACHED_NOVA_VIDA_TOKEN_TIME < 5.0):
            return _CACHED_NOVA_VIDA_TOKEN

        # Se não for forçado e o token tiver menos de 5 minutos (300s), reutiliza com segurança
        if not force_refresh and _CACHED_NOVA_VIDA_TOKEN and (now - _CACHED_NOVA_VIDA_TOKEN_TIME < 300.0):
            return _CACHED_NOVA_VIDA_TOKEN

        if not (NOVA_VIDA_USER and NOVA_VIDA_PASSWORD and NOVA_VIDA_CLIENT):
            log_api_debug("Credenciais da API Nova Vida incompletas")
            return None

        params = {
            "usuario": NOVA_VIDA_USER,
            "senha": NOVA_VIDA_PASSWORD,
            "cliente": NOVA_VIDA_CLIENT,
        }
        encoded_data = urllib.parse.urlencode(params).encode("utf-8")
        url = NOVA_VIDA_TOKEN_URL
        log_api_debug(f"Efetuando geracao de Token Nova Vida para: {url}")

        request = urllib.request.Request(
            url,
            data=encoded_data,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            method="POST"
        )

        for attempt in range(3):
            try:
                context = ssl._create_unverified_context()
                with urllib.request.urlopen(request, timeout=10, context=context) as response:
                    xml_body = response.read().decode("utf-8", errors="ignore")
                
                raw_token = extract_xml_tag(xml_body, "string")
                token = sanitize_token(raw_token)
                # Um token valido nao tem espacos; texto com espacos e mensagem de erro
                # (ex.: usuario/senha invalidos) e nao pode ser enviado como token.
                if token and re.search(r"\s", token):
                    log_api_debug(f"GerarToken NAO retornou um token (credenciais invalidas?). Resposta: {token[:120]!r}")
                    return _CACHED_NOVA_VIDA_TOKEN if (_CACHED_NOVA_VIDA_TOKEN and not force_refresh) else None
                if token and len(token) > 10:
                    log_api_debug("Token extraído com sucesso")
                    _CACHED_NOVA_VIDA_TOKEN = token
                    _CACHED_NOVA_VIDA_TOKEN_TIME = time.time()
                    return token
            except Exception as exc:
                app.logger.warning("Nova Vida token generation attempt %d failed: %s", attempt + 1, exc)
                log_api_debug(f"Falha na tentativa {attempt + 1} de geração do token: {exc}")
                time.sleep(0.3 * (attempt + 1))

        log_api_debug("Não foi possível gerar novo token Nova Vida após 3 tentativas")
        return _CACHED_NOVA_VIDA_TOKEN


def read_csv_rows(file_bytes: bytes) -> List[Dict[str, Any]]:
    # Tenta encodings comuns sem decodificar arquivo inteiro desnecessariamente
    encodings_to_try = ["utf-8-sig", "utf-8", "cp1252", "iso-8859-1", "latin1"]
    chosen_encoding = "utf-8"
    for enc in encodings_to_try:
        try:
            file_bytes[:8192].decode(enc)
            chosen_encoding = enc
            break
        except Exception:
            continue

    text_stream = io.StringIO(file_bytes.decode(chosen_encoding, errors="replace"))
    sample = text_stream.read(8192)
    text_stream.seek(0)

    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        delimiter = dialect.delimiter
    except Exception:
        # Heurística rápida para delimitador mais frequente
        counts = {
            ";": sample.count(";"),
            ",": sample.count(","),
            "\t": sample.count("\t"),
            "|": sample.count("|"),
        }
        delimiter = max(counts, key=counts.get) if max(counts.values()) > 0 else ","

    reader = csv.reader(text_stream, delimiter=delimiter)
    headers = None
    for r in reader:
        if any(c and c.strip() for c in r):
            headers = [repair_mojibake_text(c.strip()) or f"campo_{i+1}" for i, c in enumerate(r)]
            break

    if not headers:
        return []

    cleaned_rows: List[Dict[str, Any]] = []
    num_headers = len(headers)
    for row in reader:
        if not row or not any(row):
            continue
        cleaned_row = {
            headers[i] if i < num_headers else f"campo_{i+1}": val.strip() if isinstance(val, str) else val
            for i, val in enumerate(row)
        }
        cleaned_rows.append(cleaned_row)


    return cleaned_rows



def read_excel_rows(file_bytes: bytes) -> List[Dict[str, Any]]:
    workbook = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    sheet = workbook.active
    row_iter = sheet.iter_rows(values_only=True)
    
    first_row = None
    for r in row_iter:
        if any(v is not None and str(v).strip() != "" for v in r):
            first_row = r
            break

    if not first_row:
        workbook.close()
        return []

    headers = []
    for index, value in enumerate(first_row):
        if value is None or str(value).strip() == "":
            headers.append(f"campo_{index + 1}")
        else:
            headers.append(repair_mojibake_text(str(value).strip()))

    parsed_rows: List[Dict[str, Any]] = []
    for row in row_iter:
        if not any(value is not None and str(value).strip() != "" for value in row):
            continue
        parsed_row = {}
        for index, value in enumerate(row):
            if index >= len(headers):
                header_name = f"campo_{index + 1}"
            else:
                header_name = headers[index]

            if isinstance(value, (datetime, date)):
                parsed_row[header_name] = format_to_brazilian_date(value)
            elif isinstance(value, dt_time):
                parsed_row[header_name] = value.strftime("%H:%M:%S")
            elif isinstance(value, str):
                parsed_row[header_name] = repair_mojibake_text(value)
            else:
                parsed_row[header_name] = value
        parsed_rows.append(parsed_row)

    workbook.close()
    return parsed_rows


def is_row_success(row: Dict[str, Any]) -> bool:
    """Verifica se uma linha foi enriquecida com sucesso ou falhou."""
    status = str(row.get("_status_enriquecimento") or "").strip().lower()
    if status == "sucesso":
        return True
    if status == "falha":
        return False

    enr_status = str(row.get("enriquecido_status") or "").strip()
    if enr_status and enr_status != "":
        return True

    # Fallback heurístico:
    has_enriched_data = any(
        row.get(f) is not None and str(row.get(f)).strip() != "" for f in [
            "enriquecido_data_nascimento",
            "enriquecido_score",
            "enriquecido_renda_estimada",
            "enriquecido_cbo",
            "enriquecido_profissao",
            "enriquecido_idade",
            "enriquecido_sexo"
        ]
    )
    if has_enriched_data:
        return True

    enr_cpf = re.sub(r"\D", "", str(row.get("enriquecido_cpf") or ""))
    orig_cpf = re.sub(r"\D", "", str(row.get("cpf") or ""))
    if enr_cpf and enr_cpf != orig_cpf:
        return True

    return False


def is_valid_queue_row(row: Dict[str, Any]) -> bool:
    """Verifica se o registro possui no mínimo Nome e Telefone válidos para subir na fila."""
    if not isinstance(row, dict):
        return False

    # 1. Nome obrigatório (fast check)
    raw_nome = row.get("nome") or row.get("enriquecido_nome") or row.get("nome_completo") or row.get("Nome")
    if not raw_nome:
        raw_nome = resolve_name_value(row)
    if not raw_nome or not str(raw_nome).strip():
        return False

    # 2. Telefone obrigatório (fast check)
    raw_phone = row.get("telefone") or row.get("enriquecido_telefone") or row.get("celular") or row.get("Telefone")
    if not raw_phone:
        raw_phone = resolve_phone_value(row)
    phone_clean = re.sub(r"\D", "", str(raw_phone or ""))
    if phone_clean and not re.match(r"^0+$", phone_clean) and len(phone_clean) >= 8:
        return True

    for k in ("enriquecido_telefone", "telefone", "celular", "cel1", "fone", "tel", "enriquecido_telefone_2", "Telefone_2"):
        val = row.get(k)
        if val:
            val_clean = re.sub(r"\D", "", str(val))
            if val_clean and not re.match(r"^0+$", val_clean) and len(val_clean) >= 8:
                return True

    return False


def build_download_filename(
    requester_name: str = "",
    completed_at: str = None,
    mode: str = "enrich_only",
    queue_number: str = "",
    default_name: str = "usuario",
    suffix: str = "",
) -> str:
    time_str = ""
    if completed_at and isinstance(completed_at, str):
        m = re.search(r"(\d{2}):(\d{2})", completed_at)
        if m:
            time_str = f"{m.group(1)}h{m.group(2)}"

    if not time_str:
        time_str = datetime.now().strftime("%Hh%M")

    clean_queue = re.sub(r'[\/\\:\*\?"<>\|\(\)\s]', "", str(queue_number or "")).strip("_")
    if not clean_queue or clean_queue in ("TEMP", "ENRICH"):
        clean_queue = "1"

    suf = f"_{suffix}" if suffix else ""

    if mode in ("queue_only", "queue", "upload_queue_only"):
        return f"fila_{clean_queue}_{time_str}{suf}.xlsx"

    if mode in ("enrich_and_queue", "enrich_and_upload"):
        return f"enri_fila_{clean_queue}_{time_str}{suf}.xlsx"

    # Modo enrich_only (Apenas Enriquecer) -> Nome do solicitante + Horário
    req = (requester_name or default_name).strip()
    match_inner = re.search(r"\(([^)]+)\)", req)
    if match_inner and match_inner.group(1).strip():
        req_clean_name = match_inner.group(1).strip()
    else:
        req_clean_name = req

    clean_req = re.sub(r'[\/\\:\*\?"<>\|\(\)]', "", req_clean_name)
    clean_req = re.sub(r"\s+", "_", clean_req).strip("_")
    if not clean_req:
        clean_req = "usuario"

    return f"{clean_req}_{time_str}{suf}.xlsx"


def get_original_sheet_title(record: Optional[Dict[str, Any]]) -> str:
    """Extrai o nome da primeira aba da planilha original enviada pelo usuário, se existir."""
    if record and record.get("stored_file"):
        sf = Path(record["stored_file"])
        if sf.exists() and sf.suffix.lower() == ".xlsx":
            try:
                from openpyxl import load_workbook
                wb = load_workbook(str(sf), read_only=True)
                if wb.sheetnames:
                    first_sheet = wb.sheetnames[0]
                    wb.close()
                    clean_s = re.sub(r'[\\*?:/\[\]]', '', str(first_sheet)).strip()[:31]
                    if clean_s:
                        return clean_s
            except Exception as e:
                log_api_debug(f"Aviso ao ler nome da aba original: {e}")
        elif sf.exists() and sf.suffix.lower() == ".csv":
            fname = record.get("filename") or sf.stem
            clean_s = re.sub(r'[\\*?:/\[\]]', '', Path(fname).stem).strip()[:31]
            if clean_s:
                return clean_s
    return "Planilha1"


_EXPORT_LOCKS_MUTEX = threading.Lock()
_EXPORT_LOCKS: Dict[str, threading.Lock] = {}

def _get_export_lock(request_id: str) -> threading.Lock:
    with _EXPORT_LOCKS_MUTEX:
        if request_id not in _EXPORT_LOCKS:
            _EXPORT_LOCKS[request_id] = threading.Lock()
        return _EXPORT_LOCKS[request_id]


def write_export_workbook(
    rows: List[Dict[str, Any]] | None = None,
    selected_fields: List[str] | None = None,
    filepath: Path = None,
    failed_rows: List[Dict[str, Any]] | None = None,
    failed_columns: List[str] | None = None,
    create_fail_sheet: bool = True,
    sheets: List[Dict[str, Any]] | None = None,
    include_summary: bool = True,
) -> None:
    # Usar streaming write_only=True para performance máxima e baixo uso de memória RAM
    workbook = Workbook(write_only=True)

    def _sanitize_val(val: Any, is_name: bool = False) -> Any:
        if val is None or val == "":
            return ""
        if isinstance(val, (int, float, bool)):
            return val
        if isinstance(val, (datetime, date)):
            return format_to_brazilian_date(val)
        if isinstance(val, dt_time):
            return val.strftime("%H:%M:%S")
        if isinstance(val, str):
            if is_name:
                val = normalize_name_value(val)
            s_val = val
            if "\n" in s_val or "\r" in s_val or "\t" in s_val or "\v" in s_val or "\f" in s_val:
                s_val = re.sub(r"[\r\n\t\v\f]+", " ", s_val).strip()
            s_val = ILLEGAL_CHARACTERS_RE.sub("", s_val)
            if len(s_val) > 32767:
                s_val = s_val[:32767]
            return s_val
        s_val = str(val)
        if is_name:
            s_val = normalize_name_value(s_val)
        s_val = ILLEGAL_CHARACTERS_RE.sub("", s_val)
        if "\n" in s_val or "\r" in s_val or "\t" in s_val or "\v" in s_val or "\f" in s_val:
            s_val = re.sub(r"[\r\n\t\v\f]+", " ", s_val).strip()
        if len(s_val) > 32767:
            s_val = s_val[:32767]
        return s_val

    if sheets is not None and len(sheets) > 0:
        for sheet_spec in sheets:
            raw_title = str(sheet_spec.get("title") or "Sheet")
            clean_title = re.sub(r'[\\*?:/\[\]]', '', raw_title).strip()[:31] or "Sheet"
            raw_cols = sheet_spec.get("columns") or []

            clean_cols = []
            orig_keys = []
            for c in raw_cols:
                c_str = str(c).strip()
                if not c_str:
                    continue
                safe_col = ILLEGAL_CHARACTERS_RE.sub("", c_str).strip()[:32767] or "Coluna"
                clean_cols.append(safe_col)
                orig_keys.append(c)

            if not clean_cols:
                clean_cols = ["Coluna 1"]
                orig_keys = ["Coluna 1"]

            sheet_rows = sheet_spec.get("rows") or []

            ws = workbook.create_sheet(title=clean_title)
            ws.append(clean_cols)

            if not sheet_rows:
                continue

            is_single_import = (clean_cols == ["importar"])
            for row in sheet_rows:
                if not isinstance(row, dict):
                    continue
                if is_single_import:
                    val = row.get("importar")
                    if val is None:
                        val = row.get(orig_keys[0]) or ""
                    ws.append([_sanitize_val(val, is_name=False)])
                else:
                    row_vals = []
                    for orig_field, col_header in zip(orig_keys, clean_cols):
                        val = row.get(orig_field)
                        if val is None:
                            val = row.get(col_header)
                        if val is None:
                            val = get_field_value(row, orig_field)
                        if val is None:
                            val = get_field_value(row, col_header)

                        row_vals.append(_sanitize_val(val, is_name=is_required_name_field(col_header)))

                    ws.append(row_vals)

        # Aba de Resumo Analítico (somente quando include_summary for True)
        if include_summary:
            total_rows_count = sum(len(s.get("rows") or []) for s in sheets if s.get("title") != "Resumo")
            if total_rows_count > 0:
                ws_resumo = workbook.create_sheet(title="Resumo")
                ws_resumo.append(["Métrica", "Valor"])
                success_sheet = next((s for s in sheets if s.get("title") in ("Enriquecido", "Importar", "Sucesso")), None)
                fail_sheet = next((s for s in sheets if s.get("title") in ("Falha", "Não Localizado")), None)
                succ_count = len(success_sheet.get("rows") or []) if success_sheet else (total_rows_count - (len(fail_sheet.get("rows") or []) if fail_sheet else 0))
                fail_count = len(fail_sheet.get("rows") or []) if fail_sheet else (total_rows_count - succ_count)
                pct_succ = f"{(succ_count / total_rows_count * 100):.1f}%" if total_rows_count > 0 else "100%"
                pct_fail = f"{(fail_count / total_rows_count * 100):.1f}%" if total_rows_count > 0 else "0%"
                ws_resumo.append(["Total de Registros Processados", str(total_rows_count)])
                ws_resumo.append(["Registros Enriquecidos com Sucesso", f"{succ_count} ({pct_succ})"])
                ws_resumo.append(["Registros com Falha / Pendentes", f"{fail_count} ({pct_fail})"])
                ws_resumo.append(["Data e Hora de Geração", datetime.now().strftime("%d/%m/%Y %H:%M:%S")])
                ws_resumo.append(["Status da Operação", "Concluído"])

        tmp_file = filepath.with_suffix(f".tmp_{uuid.uuid4().hex[:8]}")
        workbook.save(tmp_file)
        
        # Salvamento resiliente a concorrência de leitura/escrita no Windows
        saved_ok = False
        for _ in range(5):
            try:
                tmp_file.replace(filepath)
                saved_ok = True
                break
            except Exception:
                time.sleep(0.08)
        if not saved_ok:
            try:
                import shutil
                shutil.copyfile(tmp_file, filepath)
                tmp_file.unlink(missing_ok=True)
            except Exception as e:
                log_api_debug(f"Aviso ao salvar export workbook {filepath}: {e}")
        return

    # Fallback legado
    if failed_rows is None:
        if create_fail_sheet:
            success_list = []
            fail_list = []
            for r in (rows or []):
                if is_row_success(r):
                    success_list.append(r)
                else:
                    fail_list.append(r)
        else:
            success_list = rows or []
            fail_list = []
    else:
        success_list = rows or []
        fail_list = failed_rows or []

    fail_cols = failed_columns if failed_columns is not None else (selected_fields or [])

    clean_selected = [ILLEGAL_CHARACTERS_RE.sub("", str(c)).strip()[:32767] for c in (selected_fields or []) if str(c).strip()]
    sheet_export = workbook.create_sheet(title="export")
    sheet_export.append(clean_selected)
    for row in success_list:
        row_vals = [_sanitize_val(get_field_value(row, field), is_name=is_required_name_field(field)) for field in (selected_fields or [])]
        sheet_export.append(row_vals)

    if create_fail_sheet and fail_list:
        clean_fail_cols = [ILLEGAL_CHARACTERS_RE.sub("", str(c)).strip()[:32767] for c in fail_cols if str(c).strip()]
        sheet_failed = workbook.create_sheet(title="Falha")
        sheet_failed.append(clean_fail_cols)
        for row in fail_list:
            row_vals = [_sanitize_val(get_field_value(row, field), is_name=is_required_name_field(field)) for field in fail_cols]
            sheet_failed.append(row_vals)

    tmp_file = filepath.with_suffix(f".tmp_{uuid.uuid4().hex[:8]}")
    workbook.save(tmp_file)
    saved_ok = False
    for _ in range(5):
        try:
            tmp_file.replace(filepath)
            saved_ok = True
            break
        except Exception:
            time.sleep(0.08)
    if not saved_ok:
        try:
            import shutil
            shutil.copyfile(tmp_file, filepath)
            tmp_file.unlink(missing_ok=True)
        except Exception as e:
            log_api_debug(f"Aviso ao salvar export workbook legado {filepath}: {e}")






def pick_preview_row(rows: List[Dict[str, Any]], detected_type: str) -> Dict[str, Any]:
    for row in rows:
        if not isinstance(row, dict):
            continue
        enriched = generate_mock_row(row, detected_type)
        phone_value = re.sub(r"\D", "", str(enriched.get("enriquecido_telefone") or enriched.get("telefone") or enriched.get("celular") or ""))
        if phone_value and not re.match(r"^0+$", phone_value):
            return enriched
    for row in rows:
        if not isinstance(row, dict):
            continue
        return generate_mock_row(row, detected_type)
    return {}


def call_nova_vida_api(rows: List[Dict[str, Any]], detected_type: str) -> List[Dict[str, Any]] | None:
    import requests as _req
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry

    log_api_debug("call_nova_vida_api chamado")
    base_url = (NOVA_VIDA_API_URL or "").rstrip("/")
    if not base_url:
        log_api_debug("NOVA_VIDA_API_URL não está configurada")
        return None

    token = get_nova_vida_token()
    if not token:
        log_api_debug("Autenticação falhou: sem token de acesso")
        return None

    endpoint = base_url + (NOVA_VIDA_API_ENDPOINT if NOVA_VIDA_API_ENDPOINT.startswith("/") else f"/{NOVA_VIDA_API_ENDPOINT}")
    log_api_debug(f"Efetuando chamadas de enriquecimento para: {endpoint}")

    _retry = Retry(total=1, backoff_factor=0, raise_on_status=False)
    _adapter = HTTPAdapter(pool_connections=4, pool_maxsize=4, max_retries=_retry)
    _http = _req.Session()
    _http.mount("https://", _adapter)
    _http.mount("http://", _adapter)

    enriched_rows: List[Dict[str, Any]] = []

    for index, row in enumerate(rows):
        cpf_val = resolve_cpf_value(row)
        raw_nome = resolve_name_value(row)
        nome_val = clean_name_for_api(raw_nome)
        uf_val = resolve_uf_value(row)
        cidade_val = clean_name_for_api(resolve_cidade_value(row))
        telefone_val = resolve_phone_value(row)

        def query_cpf(doc_cpf: str) -> str | None:
            nonlocal token
            clean_cpf = re.sub(r"\D", "", str(doc_cpf or ""))
            if not clean_cpf:
                return None
            if 9 <= len(clean_cpf) <= 11:
                clean_cpf = clean_cpf.zfill(11)
            elif len(clean_cpf) != 11:
                return None

            for attempt in range(2):
                try:
                    resp = _http.post(
                        f"{base_url}/WSLocalizador.asmx/NVCHECK",
                        data={"documento": clean_cpf, "token": token},
                        headers={"Content-Type": "application/x-www-form-urlencoded"},
                        timeout=(3, 10),
                        verify=False,
                    )
                    if resp.status_code == 200:
                        text_res = resp.text
                        if is_token_expired_response(text_res):
                            token = get_nova_vida_token(force_refresh=True)
                            continue
                        return text_res
                    return None
                except Exception as e:
                    log_api_debug(f"Erro no NVCHECK para {clean_cpf}: {e}")
                    return None
            return None

        def query_phone(phone_str: str) -> str | None:
            nonlocal token
            digits = re.sub(r"\D", "", str(phone_str or ""))
            if len(digits) < 10 or len(digits) > 11:
                return None

            for attempt in range(2):
                try:
                    resp = _http.post(
                        f"{base_url}/WSLocalizador.asmx/NVCHECK_TELEFONE",
                        data={"telefone": digits, "tipoDocumento": "PF", "token": token},
                        headers={"Content-Type": "application/x-www-form-urlencoded"},
                        timeout=(3, 10),
                        verify=False,
                    )
                    if resp.status_code == 200:
                        text_res = resp.text
                        if is_token_expired_response(text_res):
                            token = get_nova_vida_token(force_refresh=True)
                            continue
                        return text_res
                    return None
                except Exception as e:
                    log_api_debug(f"Erro no NVCHECK_TELEFONE para {digits}: {e}")
                    return None
            return None

        def query_name(query_nome: str, uf_filter: str = "") -> str | None:
            nonlocal token
            if not query_nome:
                return None

            payloads = []
            if uf_filter:
                payloads.append({"nome": query_nome, "uf": uf_filter})
            payloads.append({"nome": query_nome})

            for payload in payloads:
                for attempt in range(2):
                    try:
                        resp = _http.post(
                            f"{base_url}/WSLocalizador.asmx/PESQUISAR",
                            data=json.dumps(payload),
                            headers={
                                "Content-Type": "application/json",
                                "token": token,
                                "consulta": "2145"
                            },
                            timeout=(3, 10),
                            verify=False,
                        )
                        if resp.status_code == 200:
                            text_res = resp.text
                            if is_token_expired_response(text_res):
                                token = get_nova_vida_token(force_refresh=True)
                                continue
                            if "CADASTRAIS" in text_res:
                                return text_res
                        elif resp.status_code == 204:
                            break
                    except Exception as e:
                        log_api_debug(f"Erro no PESQUISAR 2145 para {query_nome}: {e}")
            return None

        try:
            target_name = raw_nome or nome_val
            digits_phone_clean = re.sub(r"\D", "", telefone_val)

            final_cpf = None
            final_nome = None
            final_score = 0.0
            final_level = "DIVERGENTE"
            final_parsed = None

            # 1. Se tem CPF na planilha original, busca direto no NVCHECK (enriquece com o que retornar)
            if cpf_val:
                xml_res = query_cpf(cpf_val)
                ok, c_cpf, c_nome, c_score, c_level, c_parsed = parse_and_validate_xml_for_target(xml_res, target_name, target_uf=uf_val, is_cpf_query=True)
                if ok:
                    final_cpf, final_nome, final_score, final_level, final_parsed = c_cpf, c_nome, c_score, c_level, c_parsed

            # 2. Se tem Telefone, busca via NVCHECK_TELEFONE
            if not final_parsed and digits_phone_clean:
                xml_res = query_phone(digits_phone_clean)
                if xml_res:
                    ok, c_cpf, c_nome, c_score, c_level, c_parsed = parse_and_validate_xml_for_target(xml_res, target_name, target_uf=uf_val, is_cpf_query=not bool(target_name))
                    if ok:
                        final_cpf, final_nome, final_score, final_level, final_parsed = c_cpf, c_nome, c_score, c_level, c_parsed

            # 3. Se tem Nome, busca via PESQUISAR (consulta 2145) com UF e sem UF
            if not final_parsed and nome_val:
                json_res = query_name(nome_val, uf_filter=uf_val)
                if json_res:
                    ok, c_cpf, c_nome, c_score, c_level, c_parsed = parse_and_validate_xml_for_target(json_res, target_name, target_uf=uf_val)
                    if ok:
                        final_cpf, final_nome, final_score, final_level, final_parsed = c_cpf, c_nome, c_score, c_level, c_parsed

            # 4. Fallback: Se não achou por nome exato, tenta variantes de nome
            if not final_parsed and nome_val:
                variants = generate_name_query_variants(nome_val)
                for variant in variants:
                    if variant != nome_val:
                        json_res = query_name(variant, uf_filter=uf_val)
                        if json_res:
                            ok, c_cpf, c_nome, c_score, c_level, c_parsed = parse_and_validate_xml_for_target(json_res, target_name, target_uf=uf_val)
                            if ok:
                                final_cpf, final_nome, final_score, final_level, final_parsed = c_cpf, c_nome, c_score, c_level, c_parsed
                                break

            # 5. Se não achou nenhum registro compatível
            if not final_parsed:
                raise ValueError("Não localizado ou titular divergente")

            # Se obtivemos CPF através de Telefone ou Nome, consulta NVCHECK para dados 360° completos
            if final_cpf:
                nv_res = query_cpf(final_cpf)
                if nv_res:
                    parsed_360 = parse_nova_vida_full_xml(html.unescape(nv_res))
                    if parsed_360:
                        for k_p, v_p in parsed_360.items():
                            if v_p:
                                final_parsed[k_p] = v_p

            if target_name:
                similaridade_str = f"{final_level} ({final_score:.0f}%)"
            else:
                similaridade_str = "NAO_APLICAVEL"

            enriched = dict(row)
            enriched.update({
                "enriquecido_nome": normalize_name_value(final_parsed.get("nome") or raw_nome or nome_val),
                "enriquecido_telefone": final_parsed.get("telefone") or re.sub(r"\D", "", telefone_val),
                "enriquecido_cpf": final_parsed.get("cpf") or cpf_val,
                "enriquecido_data_nascimento": final_parsed.get("nascimento") or "",
                "enriquecido_idade": final_parsed.get("idade") or "",
                "enriquecido_sexo": final_parsed.get("sexo") or "",
                "enriquecido_email": final_parsed.get("email") or "",
                "enriquecido_logradouro": final_parsed.get("logradouro") or "",
                "enriquecido_numero": final_parsed.get("numero") or "",
                "enriquecido_complemento": final_parsed.get("complemento") or "",
                "enriquecido_bairro": final_parsed.get("bairro") or "",
                "enriquecido_cidade": final_parsed.get("cidade") or cidade_val,
                "enriquecido_estado": final_parsed.get("estado") or uf_val,
                "enriquecido_cep": final_parsed.get("cep") or "",
                "enriquecido_renda_estimada": final_parsed.get("renda_estimada") or "",
                "enriquecido_cbo": final_parsed.get("cbo") or "",
                "enriquecido_profissao": final_parsed.get("profissao") or "",
                "enriquecido_score": final_parsed.get("score") or "",
                "enriquecido_perfil": final_parsed.get("perfil") or "",
                "enriquecido_obito": final_parsed.get("obito") or "0",
                "enriquecido_whatsapp": final_parsed.get("whatsapp") or "NÃO",
                "enriquecido_status": final_parsed.get("status") or "REGULAR",
                "enriquecido_similaridade_nome": similaridade_str,
                "tipo_enriquecimento": detected_type,
            })

            for key, val in list(row.items()):
                enriched.setdefault(normalize_text(key), val)

            log_api_debug(f"Sucesso ao enriquecer {raw_nome or nome_val}: CPF={final_parsed.get('cpf')}, Similaridade={similaridade_str}")
            enriched_rows.append(enriched)

        except Exception as exc:
            log_api_debug(f"Não foi possível enriquecer {raw_nome or nome_val}: {exc}")
            phone_final = re.sub(r"\D", "", telefone_val)
            enriched = dict(row)
            enriched.update({
                "enriquecido_nome": normalize_name_value(raw_nome or nome_val),
                "enriquecido_telefone": phone_final,
                "enriquecido_cpf": cpf_val or "",
                "enriquecido_data_nascimento": "",
                "enriquecido_idade": "",
                "enriquecido_sexo": "",
                "enriquecido_email": "",
                "enriquecido_logradouro": "",
                "enriquecido_numero": "",
                "enriquecido_complemento": "",
                "enriquecido_bairro": "",
                "enriquecido_cidade": cidade_val,
                "enriquecido_estado": uf_val,
                "enriquecido_cep": "",
                "enriquecido_renda_estimada": "",
                "enriquecido_cbo": "",
                "enriquecido_profissao": "",
                "enriquecido_score": "",
                "enriquecido_perfil": "",
                "enriquecido_obito": "",
                "enriquecido_whatsapp": "",
                "enriquecido_status": "",
                "enriquecido_similaridade_nome": "",
                "tipo_enriquecimento": detected_type,
            })
            for key, val in list(row.items()):
                enriched.setdefault(normalize_text(key), val)
            enriched_rows.append(enriched)

    return enriched_rows


def deduplicate_rows(rows: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], int]:
    """Remove registros duplicados de forma inteligente e consistente com base em CPF ou Nome + UF + Telefone."""
    if not rows:
        return [], 0

    first_row = rows[0]
    keys = list(first_row.keys())

    # Pre-resolucao de chaves de coluna para altissima performance (O(1) por linha)
    cpf_key = next((k for k in keys if is_cpf_field(k)), None)
    name_key = next((k for k in keys if is_required_name_field(k)), None)
    uf_key = next((k for k in keys if is_uf_field(k)), None)
    phone_key = next((k for k in keys if is_required_phone_field(k) and not normalize_text(k).startswith("ddd")), None)

    seen_keys = set()
    unique_rows = []
    duplicates_removed = 0

    for r in rows:
        cpf_val = ""
        if cpf_key is not None:
            raw_c = r.get(cpf_key)
            if raw_c:
                clean_c = re.sub(r"\D", "", str(raw_c))
                if len(clean_c) == 12 and clean_c.startswith("0"):
                    clean_c = clean_c[1:]
                if len(clean_c) == 11 and clean_c != "00000000000":
                    cpf_val = clean_c

        if cpf_val:
            key = f"c_{cpf_val}"
        else:
            raw_nome = str(r.get(name_key) or "").strip().lower() if name_key else ""
            if raw_nome:
                tokens = [t for t in raw_nome.split() if t not in NAME_CONNECTIVES]
                nome_clean = " ".join(tokens)
            else:
                nome_clean = ""
            uf_val = str(r.get(uf_key) or "").strip().upper()[:2] if uf_key else ""
            phone_val = str(r.get(phone_key) or "").strip() if phone_key else ""
            clean_p = re.sub(r"\D", "", phone_val)
            p_short = clean_p[-8:] if len(clean_p) >= 8 else clean_p
            key = f"n_{nome_clean}_u_{uf_val}_p_{p_short}"

        if key in seen_keys and key not in ("n___p_", "c_"):
            duplicates_removed += 1
        else:
            seen_keys.add(key)
            unique_rows.append(r)
    return unique_rows, duplicates_removed







@app.get("/login")
def login_page():
    if session.get("user_id"):
        return redirect(url_for("index"))
    allowed_origins = list(ALLOWED_ORIGINS) if "ALLOWED_ORIGINS" in globals() else ["https://estrategia-os.vercel.app", "http://localhost:47821"]
    return render_template("login.html", allowed_origins=allowed_origins)



@app.post("/login")
def login():
    is_ajax = request.is_json or request.headers.get("X-Requested-With") == "XMLHttpRequest"
    if request.is_json:
        data = request.get_json(silent=True) or {}
        username = (data.get("username") or "").strip().lower()
        password = (data.get("password") or "").strip()
    else:
        username = (request.form.get("username") or "").strip().lower()
        password = (request.form.get("password") or "").strip()

    user = USERS.get(username)

    if user and check_password_hash(user["password_hash"], password):
        session["user_id"] = username
        session["role"] = user["role"]
        session["user_name"] = user["name"]
        db_save_audit_event("LOGIN_SUCCESS", user["name"], {"username": username, "role": user["role"]})
        if is_ajax:
            return jsonify({"ok": True, "redirect": url_for("index")})
        return redirect(url_for("index"))

    db_save_audit_event("LOGIN_FAILED", username or "desconhecido", {"username": username})
    if is_ajax:
        return jsonify({"ok": False, "error": "Credenciais inválidas. Verifique seu usuário e senha."}), 401
    return render_template("login.html", error="Credenciais inválidas. Verifique seu usuário e senha."), 401



@app.get("/logout")
def logout():
    user_name = session.get("user_name", "Usuário")
    db_save_audit_event("LOGOUT", user_name)
    session.clear()
    return redirect(url_for("login_page"))


@app.get("/")
@login_required
def index():
    user_role = session.get("role", "user1")
    return render_template("index.html", user_role=user_role, user_name=session.get("user_name", "Usuário"))


CHUNK_UPLOAD_DIR = TEMP_PARENT / "chunk_uploads"
CHUNK_UPLOAD_DIR.mkdir(exist_ok=True, parents=True)


def inspect_uploaded_file(file_name: str, file_bytes: bytes) -> Dict[str, Any]:
    extension = Path(file_name).suffix.lower()
    if extension not in (".xlsx", ".csv"):
        return {"error": "Formato inválido. Envie .xlsx ou .csv."}

    columns: List[str] = []
    sample_rows: List[Dict[str, Any]] = []
    total_rows_count = 0
    duplicate_count = 0
    invalid_phone_count = 0
    invalid_phones: List[Dict[str, Any]] = []

    if extension == ".xlsx":
        workbook = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        sheet = workbook.active
        row_iter = sheet.iter_rows(values_only=True)

        first_row = None
        for r in row_iter:
            if any(v is not None and str(v).strip() != "" for v in r):
                first_row = r
                break

        if not first_row:
            workbook.close()
            return {"error": "Arquivo sem linhas para processar."}

        for idx, val in enumerate(first_row):
            if val is None or str(val).strip() == "":
                columns.append(f"campo_{idx + 1}")
            else:
                columns.append(repair_mojibake_text(str(val).strip()))

        cpf_idx = next((i for i, c in enumerate(columns) if is_cpf_field(c)), None)
        name_idx = next((i for i, c in enumerate(columns) if is_required_name_field(c)), None)
        uf_idx = next((i for i, c in enumerate(columns) if is_uf_field(c)), None)
        phone_idx = next((i for i, c in enumerate(columns) if is_required_phone_field(c) and not normalize_text(c).startswith("ddd")), None)

        seen_keys = set()

        for row in row_iter:
            if not any(v is not None and str(v).strip() != "" for v in row):
                continue

            total_rows_count += 1
            row_len = len(row)

            # Amostra das primeiras 5 linhas para visualização na interface
            if len(sample_rows) < 5:
                dict_row = {}
                for idx, col in enumerate(columns):
                    val = row[idx] if idx < row_len else ""
                    if isinstance(val, (datetime, date)):
                        dict_row[col] = format_to_brazilian_date(val)
                    elif isinstance(val, dt_time):
                        dict_row[col] = val.strftime("%H:%M:%S")
                    elif val is not None:
                        dict_row[col] = repair_mojibake_text(str(val).strip())
                    else:
                        dict_row[col] = ""
                sample_rows.append(dict_row)

            # Rastreamento de duplicados ultra performático
            if cpf_idx is not None and cpf_idx < row_len:
                raw_cpf = str(row[cpf_idx] or "").strip()
                clean_cpf = re.sub(r"\D", "", raw_cpf)
                if len(clean_cpf) == 12 and clean_cpf.startswith("0"):
                    clean_cpf = clean_cpf[1:]
                if len(clean_cpf) == 11 and clean_cpf != "00000000000":
                    key = f"c_{clean_cpf}"
                else:
                    raw_n = str(row[name_idx] or "").strip().lower() if name_idx is not None and name_idx < row_len else ""
                    raw_u = str(row[uf_idx] or "").strip().upper()[:2] if uf_idx is not None and uf_idx < row_len else ""
                    raw_p = re.sub(r"\D", "", str(row[phone_idx] or ""))[-8:] if phone_idx is not None and phone_idx < row_len else ""
                    key = f"n_{raw_n}_u_{raw_u}_p_{raw_p}"
            else:
                raw_n = str(row[name_idx] or "").strip().lower() if name_idx is not None and name_idx < row_len else ""
                raw_u = str(row[uf_idx] or "").strip().upper()[:2] if uf_idx is not None and uf_idx < row_len else ""
                raw_p = re.sub(r"\D", "", str(row[phone_idx] or ""))[-8:] if phone_idx is not None and phone_idx < row_len else ""
                key = f"n_{raw_n}_u_{raw_u}_p_{raw_p}"

            if key in seen_keys and key not in ("n___p_", "c_"):
                duplicate_count += 1
            else:
                seen_keys.add(key)

            # Validação rápida de telefone
            if phone_idx is not None and phone_idx < row_len:
                raw_phone = str(row[phone_idx] or "").strip()
                clean_phone = re.sub(r"\D", "", raw_phone)
                if not clean_phone or re.match(r"^0+$", clean_phone):
                    if len(invalid_phones) < 50:
                        n_val = str(row[name_idx] or f"Linha {total_rows_count + 1}").strip() if name_idx is not None and name_idx < row_len else f"Linha {total_rows_count + 1}"
                        invalid_phones.append({
                            "row": total_rows_count + 1,
                            "name": n_val,
                            "phone": "(Vazio)",
                            "reason": "Telefone em branco"
                        })
                    invalid_phone_count += 1
                elif len(clean_phone) < 10 or len(clean_phone) > 11:
                    if len(invalid_phones) < 50:
                        n_val = str(row[name_idx] or f"Linha {total_rows_count + 1}").strip() if name_idx is not None and name_idx < row_len else f"Linha {total_rows_count + 1}"
                        invalid_phones.append({
                            "row": total_rows_count + 1,
                            "name": n_val,
                            "phone": raw_phone,
                            "reason": "Falta DDD ou tamanho incorreto"
                        })
                    invalid_phone_count += 1

        workbook.close()

    elif extension == ".csv":
        encodings_to_try = ["utf-8-sig", "utf-8", "cp1252", "iso-8859-1", "latin1"]
        chosen_encoding = "utf-8"
        for enc in encodings_to_try:
            try:
                file_bytes[:8192].decode(enc)
                chosen_encoding = enc
                break
            except Exception:
                continue

        text_stream = io.StringIO(file_bytes.decode(chosen_encoding, errors="replace"))
        sample = text_stream.read(8192)
        text_stream.seek(0)

        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            delimiter = dialect.delimiter
        except Exception:
            counts = {";": sample.count(";"), ",": sample.count(","), "\t": sample.count("\t"), "|": sample.count("|")}
            delimiter = max(counts, key=counts.get) if max(counts.values()) > 0 else ","

        reader = csv.reader(text_stream, delimiter=delimiter)
        first_row = None
        for r in reader:
            if any(c and c.strip() for c in r):
                first_row = r
                break

        if not first_row:
            return {"error": "Arquivo sem linhas para processar."}

        columns = [repair_mojibake_text(c.strip()) or f"campo_{i+1}" for i, c in enumerate(first_row)]

        cpf_idx = next((i for i, c in enumerate(columns) if is_cpf_field(c)), None)
        name_idx = next((i for i, c in enumerate(columns) if is_required_name_field(c)), None)
        uf_idx = next((i for i, c in enumerate(columns) if is_uf_field(c)), None)
        phone_idx = next((i for i, c in enumerate(columns) if is_required_phone_field(c) and not normalize_text(c).startswith("ddd")), None)

        seen_keys = set()

        for row in reader:
            if not row or not any(row):
                continue

            total_rows_count += 1
            row_len = len(row)

            if len(sample_rows) < 5:
                dict_row = {
                    columns[i] if i < len(columns) else f"campo_{i+1}": (row[i].strip() if i < row_len else "")
                    for i in range(min(len(columns), row_len))
                }
                sample_rows.append(dict_row)

            if cpf_idx is not None and cpf_idx < row_len:
                raw_cpf = str(row[cpf_idx] or "").strip()
                clean_cpf = re.sub(r"\D", "", raw_cpf)
                if len(clean_cpf) == 12 and clean_cpf.startswith("0"):
                    clean_cpf = clean_cpf[1:]
                if len(clean_cpf) == 11 and clean_cpf != "00000000000":
                    key = f"c_{clean_cpf}"
                else:
                    raw_n = str(row[name_idx] or "").strip().lower() if name_idx is not None and name_idx < row_len else ""
                    raw_u = str(row[uf_idx] or "").strip().upper()[:2] if uf_idx is not None and uf_idx < row_len else ""
                    raw_p = re.sub(r"\D", "", str(row[phone_idx] or ""))[-8:] if phone_idx is not None and phone_idx < row_len else ""
                    key = f"n_{raw_n}_u_{raw_u}_p_{raw_p}"
            else:
                raw_n = str(row[name_idx] or "").strip().lower() if name_idx is not None and name_idx < row_len else ""
                raw_u = str(row[uf_idx] or "").strip().upper()[:2] if uf_idx is not None and uf_idx < row_len else ""
                raw_p = re.sub(r"\D", "", str(row[phone_idx] or ""))[-8:] if phone_idx is not None and phone_idx < row_len else ""
                key = f"n_{raw_n}_u_{raw_u}_p_{raw_p}"

            if key in seen_keys and key not in ("n___p_", "c_"):
                duplicate_count += 1
            else:
                seen_keys.add(key)

            if phone_idx is not None and phone_idx < row_len:
                raw_phone = str(row[phone_idx] or "").strip()
                clean_phone = re.sub(r"\D", "", raw_phone)
                if not clean_phone or re.match(r"^0+$", clean_phone):
                    if len(invalid_phones) < 50:
                        n_val = str(row[name_idx] or f"Linha {total_rows_count + 1}").strip() if name_idx is not None and name_idx < row_len else f"Linha {total_rows_count + 1}"
                        invalid_phones.append({
                            "row": total_rows_count + 1,
                            "name": n_val,
                            "phone": "(Vazio)",
                            "reason": "Telefone em branco"
                        })
                    invalid_phone_count += 1
                elif len(clean_phone) < 10 or len(clean_phone) > 11:
                    if len(invalid_phones) < 50:
                        n_val = str(row[name_idx] or f"Linha {total_rows_count + 1}").strip() if name_idx is not None and name_idx < row_len else f"Linha {total_rows_count + 1}"
                        invalid_phones.append({
                            "row": total_rows_count + 1,
                            "name": n_val,
                            "phone": raw_phone,
                            "reason": "Falta DDD ou tamanho incorreto"
                        })
                    invalid_phone_count += 1

    if total_rows_count == 0:
        return {"error": "Arquivo sem linhas válidas para processar."}

    unique_rows_count = max(0, total_rows_count - duplicate_count)
    detected_type = detect_enrichment_type(columns)
    valid_lines = total_rows_count - duplicate_count - invalid_phone_count
    quality_score = max(0.0, round((valid_lines / total_rows_count) * 100, 1)) if total_rows_count > 0 else 100.0
    preview_record = pick_preview_row(sample_rows, detected_type) if sample_rows else {}

    return {
        "columns": columns,
        "total_rows_count": total_rows_count,
        "unique_rows_count": unique_rows_count,
        "duplicate_count": duplicate_count,
        "invalid_phone_count": invalid_phone_count,
        "invalid_phones": invalid_phones,
        "detected_type": detected_type,
        "quality_score": quality_score,
        "sample_rows": sample_rows,
        "preview_record": preview_record,
    }


def _process_uploaded_file(file_name: str, file_bytes: bytes, original_path: Path) -> Tuple[Dict[str, Any], int]:
    extension = Path(file_name).suffix.lower()
    if not original_path.exists():
        original_path.write_bytes(file_bytes)

    if extension not in (".xlsx", ".csv"):
        if extension == ".xls":
            return {"error": "Formato inválido. O upload de .xls não é suportado neste ambiente. Por favor, salve o arquivo como .xlsx ou .csv."}, 400
        return {"error": "Formato inválido. Envie .xlsx ou .csv."}, 400

    try:
        inspection = inspect_uploaded_file(file_name, file_bytes)
    except Exception as exc:
        log_api_debug(f"Erro ao inspecionar planilha {file_name}: {exc}")
        return {"error": f"Erro ao ler a planilha: {exc}"}, 400

    if "error" in inspection:
        return {"error": inspection["error"]}, 400

    columns = inspection["columns"]
    total_rows_count = inspection["total_rows_count"]
    unique_rows_count = inspection["unique_rows_count"]
    duplicate_count = inspection["duplicate_count"]
    invalid_phone_count = inspection["invalid_phone_count"]
    invalid_phones = inspection["invalid_phones"]
    detected_type = inspection["detected_type"]
    quality_score = inspection["quality_score"]
    sample_rows = inspection["sample_rows"]
    preview_record = inspection["preview_record"]

    preflight_data = {
        "total_rows": total_rows_count,
        "unique_rows": unique_rows_count,
        "duplicate_count": duplicate_count,
        "invalid_phone_count": invalid_phone_count,
        "columns_count": len(columns),
        "detected_type": detected_type,
        "quality_score": quality_score,
    }

    available_fields_list = list(dict.fromkeys(columns + PREFIXED_ENRICH_FIELDS))
    api_source = "nova_vida_api" if bool(NOVA_VIDA_API_URL) else "mock"

    record = {
        "id": str(uuid.uuid4()),
        "filename": file_name,
        "stored_file": str(original_path),
        "type": detected_type,
        "columns": columns,
        "rows": [],  # Persistência leve; o carregamento integral ocorre sob demanda via stored_file
        "sample_rows": sample_rows,
        "available_fields": available_fields_list,
        "required_fields": REQUIRED_FIELDS,
        "preview": preview_record,
        "total_in": total_rows_count,
        "total_enriched": 0,
        "source": api_source,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }
    db_save_record(record)

    current_user_name = session.get("user_name", "Usuário") if has_request_context() else "Sistema"
    db_save_audit_event(
        "UPLOAD_FILE",
        current_user_name,
        {
            "filename": file_name,
            "total_rows": total_rows_count,
            "duplicates": duplicate_count,
            "detected_type": detected_type,
            "quality_score": quality_score,
        }
    )

    return {
        "id": record["id"],
        "message": "Arquivo validado e pronto para enriquecimento.",
        "required_fields": REQUIRED_FIELDS,
        "detected_type": detected_type,
        "available_fields": record["available_fields"],
        "preview": record["preview"],
        "sample_rows": sample_rows,
        "total_in": record["total_in"],
        "total_enriched": record["total_enriched"],
        "columns": record["columns"],
        "invalid_phone_count": invalid_phone_count,
        "invalid_phones": invalid_phones,
        "preflight": preflight_data,
    }, 200


@app.post("/api/upload")
@login_required
def upload_file():
    try:
        uploaded = request.files.get("file")
        if not uploaded or uploaded.filename == "":
            return jsonify({"error": "Arquivo não foi enviado."}), 400

        file_name = uploaded.filename or "arquivo.xlsx"
        extension = Path(file_name).suffix.lower()
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        original_path = UPLOAD_DIR / f"{uuid.uuid4()}_original{extension}"
        file_bytes = uploaded.read()
        uploaded.stream.seek(0)
        original_path.write_bytes(file_bytes)

        payload, status_code = _process_uploaded_file(file_name, file_bytes, original_path)
        return jsonify(payload), status_code
    except Exception as exc:
        log_api_debug(f"Erro em upload_file: {exc}")
        return jsonify({"error": f"Erro ao processar planilha: {exc}"}), 400


@app.post("/api/upload/chunk")
@login_required
def upload_chunk():
    try:
        upload_id = request.form.get("upload_id")
        chunk_index_str = request.form.get("chunk_index")
        total_chunks_str = request.form.get("total_chunks")
        filename = request.form.get("filename") or "arquivo.xlsx"
        chunk_file = request.files.get("chunk")

        if not upload_id or chunk_index_str is None or total_chunks_str is None or not chunk_file:
            return jsonify({"error": "Parâmetros do chunk incompletos."}), 400

        chunk_index = int(chunk_index_str)
        total_chunks = int(total_chunks_str)

        upload_tmp_dir = CHUNK_UPLOAD_DIR / upload_id
        upload_tmp_dir.mkdir(parents=True, exist_ok=True)

        chunk_path = upload_tmp_dir / f"{chunk_index:06d}.part"
        chunk_file.save(str(chunk_path))

        # Se for o último chunk, monta o arquivo final e processa
        if chunk_index == total_chunks - 1:
            extension = Path(filename).suffix.lower()
            UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
            original_path = UPLOAD_DIR / f"{uuid.uuid4()}_original{extension}"

            with open(original_path, "wb") as outfile:
                for i in range(total_chunks):
                    part_file = upload_tmp_dir / f"{i:06d}.part"
                    if not part_file.exists():
                        return jsonify({"error": f"Fatia {i+1} de {total_chunks} ausente no servidor."}), 400
                    outfile.write(part_file.read_bytes())

            try:
                import shutil
                shutil.rmtree(upload_tmp_dir, ignore_errors=True)
            except Exception:
                pass

            file_bytes = original_path.read_bytes()
            payload, status_code = _process_uploaded_file(filename, file_bytes, original_path)
            return jsonify(payload), status_code

        return jsonify({"success": True, "chunk_received": chunk_index, "total_chunks": total_chunks}), 200

    except Exception as exc:
        log_api_debug(f"Erro em upload_chunk: {exc}")
        return jsonify({"error": f"Erro no processamento do chunk: {exc}"}), 400


@app.post("/api/upload/merge")
@login_required
def upload_merge_files():
    try:
        uploaded_files = request.files.getlist("files") or request.files.getlist("files[]") or request.files.getlist("file")
        if not uploaded_files:
            return jsonify({"error": "Nenhuma planilha foi enviada para mesclagem."}), 400

        all_rows: List[Dict[str, Any]] = []
        file_names: List[str] = []
        master_columns: List[str] = []
        first_file_sample_rows: List[Dict[str, Any]] = []

        for upl in uploaded_files:
            if not upl or not upl.filename:
                continue
            fname = upl.filename
            file_names.append(fname)
            ext = Path(fname).suffix.lower()
            fbytes = upl.read()
            if ext == ".xlsx":
                frows = read_excel_rows(fbytes)
            elif ext == ".csv":
                frows = read_csv_rows(fbytes)
            elif ext == ".xls":
                return jsonify({"error": f"Formato .xls não suportado para '{fname}'. Salve como .xlsx ou .csv."}), 400
            else:
                return jsonify({"error": f"Formato não suportado para '{fname}'. Envie .xlsx ou .csv."}), 400

            if not frows:
                continue

            file_cols = list(frows[0].keys())

            # Se for a primeira planilha com dados, ela define o padrão de colunas e ordem master
            if not master_columns:
                master_columns = list(file_cols)
                first_file_sample_rows = list(frows[:5])
                all_rows.extend(frows)
            else:
                # Alinha as linhas das planilhas subsequentes com o padrão da 1ª planilha
                aligned_frows = []
                # Se tem a mesma quantidade de colunas, mapeia por posição/ordem sequencial (ex: Col A -> Col A)
                if len(file_cols) == len(master_columns):
                    for r in frows:
                        vals = list(r.values())
                        row_dict = {master_columns[i]: vals[i] if i < len(vals) else "" for i in range(len(master_columns))}
                        aligned_frows.append(row_dict)
                else:
                    # Mapeia por correspondência de nomes de coluna normalizados
                    norm_master_map = {normalize_text(c): c for c in master_columns}
                    for r in frows:
                        row_dict = {c: "" for c in master_columns}
                        for k, v in r.items():
                            norm_k = normalize_text(k)
                            if norm_k in norm_master_map:
                                row_dict[norm_master_map[norm_k]] = v
                            else:
                                row_dict[k] = v
                        aligned_frows.append(row_dict)

                all_rows.extend(aligned_frows)

        if not all_rows or not master_columns:
            return jsonify({"error": "Nenhum registro encontrado nas planilhas enviadas."}), 400

        # Limpa o dataframe e preserva a ordem master da 1ª planilha
        cleaned_rows = clean_dataframe(all_rows)
        columns = list(master_columns)
        for k in (cleaned_rows[0].keys() if cleaned_rows else []):
            if k not in columns:
                columns.append(k)

        detected_type = detect_enrichment_type(columns)

        # Salva o arquivo Excel consolidado
        merged_filename = f"Mesclado_{len(file_names)}_planilhas.xlsx"
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        original_path = UPLOAD_DIR / f"{uuid.uuid4()}_merged.xlsx"

        wb = Workbook(write_only=True)
        ws = wb.create_sheet(title="Planilha Mesclada")
        clean_cols = [ILLEGAL_CHARACTERS_RE.sub("", str(c)).strip()[:32767] for c in columns if str(c).strip()]
        ws.append(clean_cols)
        for r in cleaned_rows:
            row_vals = []
            for c in columns:
                v = r.get(c, "")
                if isinstance(v, str):
                    v = ILLEGAL_CHARACTERS_RE.sub("", v)[:32767]
                row_vals.append(v)
            ws.append(row_vals)
        wb.save(original_path)

        # Diagnóstico pré-upload (Pre-flight Check) & Detecção de duplicados
        unique_rows, duplicate_count = deduplicate_rows(cleaned_rows)
        total_rows_count = len(cleaned_rows)
        unique_rows_count = len(unique_rows)

        # Contagem de telefones inválidos
        invalid_phone_count = 0
        invalid_phones = []
        has_any_phone_col = any(is_required_phone_field(col) for col in columns)
        if has_any_phone_col and cleaned_rows:
            p_key = next((k for k in cleaned_rows[0].keys() if is_required_phone_field(k) and not normalize_text(k).startswith("ddd")), None)
            n_key = next((k for k in cleaned_rows[0].keys() if is_required_name_field(k)), None)
            for idx, r in enumerate(cleaned_rows):
                raw_p = r.get(p_key) if p_key else resolve_phone_value(r)
                clean_p = re.sub(r"\D", "", str(raw_p or ""))
                if not clean_p or re.match(r"^0+$", clean_p):
                    if len(invalid_phones) < 50:
                        nome_val = (r.get(n_key) if n_key else resolve_name_value(r)) or f"Linha {idx + 2}"
                        invalid_phones.append({
                            "row": idx + 2,
                            "name": str(nome_val),
                            "phone": "(Vazio)",
                            "reason": "Telefone em branco"
                        })
                    invalid_phone_count += 1
                elif len(clean_p) < 10 or len(clean_p) > 11:
                    if len(invalid_phones) < 50:
                        nome_val = (r.get(n_key) if n_key else resolve_name_value(r)) or f"Linha {idx + 2}"
                        invalid_phones.append({
                            "row": idx + 2,
                            "name": str(nome_val),
                            "phone": str(raw_p),
                            "reason": "Falta DDD ou tamanho incorreto"
                        })
                    invalid_phone_count += 1

        valid_lines = total_rows_count - duplicate_count - invalid_phone_count
        quality_score = max(0.0, round((valid_lines / total_rows_count) * 100, 1)) if total_rows_count > 0 else 100.0

        preflight_data = {
            "total_rows": total_rows_count,
            "unique_rows": unique_rows_count,
            "duplicate_count": duplicate_count,
            "invalid_phone_count": invalid_phone_count,
            "columns_count": len(columns),
            "detected_type": detected_type,
            "quality_score": quality_score,
            "merged_files_count": len(file_names),
            "merged_files": file_names,
        }

        preview_record = pick_preview_row(cleaned_rows, detected_type)
        available_fields_list = list(dict.fromkeys(columns + PREFIXED_ENRICH_FIELDS))
        api_source = "nova_vida_api" if bool(NOVA_VIDA_API_URL) else "mock"

        sample_rows = [
            {k: str(v) if v is not None else "" for k, v in list(r.items())[:10]}
            for r in cleaned_rows[:5]
        ]

        record = {
            "id": str(uuid.uuid4()),
            "filename": merged_filename,
            "stored_file": str(original_path),
            "type": detected_type,
            "columns": columns,
            "rows": [],
            "sample_rows": sample_rows,
            "available_fields": available_fields_list,
            "required_fields": REQUIRED_FIELDS,
            "preview": preview_record,
            "total_in": len(cleaned_rows),
            "total_enriched": 0,
            "source": api_source,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "is_merged": True,
            "merged_files": file_names,
        }
        db_save_record(record)

        db_save_audit_event(
            "UPLOAD_MERGED_FILES",
            session.get("user_name"),
            {
                "filename": merged_filename,
                "files_count": len(file_names),
                "total_rows": total_rows_count,
                "duplicates": duplicate_count,
                "detected_type": detected_type,
                "quality_score": quality_score,
            }
        )

        return {
            "id": record["id"],
            "message": f"{len(file_names)} planilhas mescladas com sucesso ({total_rows_count} linhas).",
            "required_fields": REQUIRED_FIELDS,
            "detected_type": detected_type,
            "available_fields": record["available_fields"],
            "preview": record["preview"],
            "sample_rows": sample_rows,
            "total_in": record["total_in"],
            "total_enriched": record["total_enriched"],
            "columns": record["columns"],
            "invalid_phone_count": invalid_phone_count,
            "invalid_phones": invalid_phones,
            "preflight": preflight_data,
        }, 200

    except Exception as exc:
        log_api_debug(f"Erro em upload_merge_files: {exc}")
        return jsonify({"error": f"Erro ao mesclar planilhas: {exc}"}), 400




DEFAULT_ENRICHMENT_CONCURRENCY = 80
MAX_ENRICHMENT_CONCURRENCY = 200


def get_enrichment_concurrency(value: str | None = None) -> int:
    configured = value if value is not None else os.environ.get("ENRICHMENT_CONCURRENCY")
    if configured is None or not configured.strip():
        return DEFAULT_ENRICHMENT_CONCURRENCY

    try:
        requested = int(configured)
    except ValueError:
        app.logger.warning(
            "Invalid ENRICHMENT_CONCURRENCY value %r; using default %d",
            configured,
            DEFAULT_ENRICHMENT_CONCURRENCY,
        )
        return DEFAULT_ENRICHMENT_CONCURRENCY

    if requested < 1:
        app.logger.warning("ENRICHMENT_CONCURRENCY must be >= 1; using 1")
        return 1
    if requested > MAX_ENRICHMENT_CONCURRENCY:
        app.logger.warning(
            "ENRICHMENT_CONCURRENCY=%d exceeds the provider-safe limit; capped at %d",
            requested,
            MAX_ENRICHMENT_CONCURRENCY,
        )
        return MAX_ENRICHMENT_CONCURRENCY
    return requested


async def _async_enrichment_engine(queue_item_id: str, raw_rows: List[Dict[str, Any]], detected_type: str, query_fields: List[str]):
    """Motor assincrono de enriquecimento controlado com cadencia segura."""
    import asyncio
    import aiohttp
    import html
    import time as _time

    item = db_get_queue_item(queue_item_id)
    if not item:
        return

    total_rows = len(raw_rows)
    log_api_debug(f"Motor assíncrono iniciado para {total_rows} registros")

    # ── Contadores compartilhados (acesso single-threaded no event-loop) ──────
    counters = {"processed": 0, "success": 0, "error": 0}
    api_attempt_counts: Dict[str, int] = {}
    logs_list: list = item.setdefault("logs", [])

    def _record_api_attempt(endpoint: str, outcome: str) -> None:
        key = f"{endpoint}:{outcome}"
        api_attempt_counts[key] = api_attempt_counts.get(key, 0) + 1

    # ── Throttle de UI: propaga logs/status a cada 150ms OU 10 registros ──
    _last_ui_flush_time: list = [_time.monotonic()]
    _ui_pending_logs: list = []
    LOG_THROTTLE_INTERVAL = 0.150   # segundos
    LOG_THROTTLE_BATCH    = 10      # registros

    def _log(msg: str):
        """Acumula mensagem; o flush real é feito por _maybe_flush_ui."""
        _ui_pending_logs.append(msg)

    def _flush_ui_logs():
        """Transfere logs pendentes para logs_list respeitando o limite de 300."""
        while _ui_pending_logs:
            msg = _ui_pending_logs.pop(0)
            if len(logs_list) < 300:
                logs_list.append(msg)
            elif len(logs_list) == 300:
                logs_list.append("... logs subsequentes ocultados para economizar memória ...")
            # Descarta silenciosamente acima do limite

    def _maybe_flush_ui():
        """Flush com debounce: a cada 150ms OU a cada 10 registros processados."""
        now = _time.monotonic()
        time_elapsed  = (now - _last_ui_flush_time[0]) >= LOG_THROTTLE_INTERVAL
        batch_reached = (counters["processed"] % LOG_THROTTLE_BATCH) == 0
        if time_elapsed or batch_reached:
            _flush_ui_logs()
            _last_ui_flush_time[0] = now


    def _inc(success: bool):
        counters["processed"] += 1
        if success:
            counters["success"] += 1
        else:
            counters["error"] += 1
        # mirror into item dict for db_update_queue_item compatibility
        item["processed_count"] = counters["processed"]
        item["success_count"]   = counters["success"]
        item["error_count"]     = counters["error"]
        _maybe_flush_ui()

    try:
        is_api_configured = bool(NOVA_VIDA_API_URL)
        token = None
        async_token_lock = asyncio.Lock()

        async def ensure_token_valid(force_refresh: bool = False) -> str | None:
            nonlocal token, is_api_configured
            if not is_api_configured:
                return None
            async with async_token_lock:
                now = time.time()
                # Se já temos um token válido gerado recentemente (< 4s), reutiliza para evitar avalanche
                if token and (now - _CACHED_NOVA_VIDA_TOKEN_TIME < 4.0):
                    return token
                if force_refresh or not token or (now - _CACHED_NOVA_VIDA_TOKEN_TIME > 240.0):
                    log_api_debug("Renovando token Nova Vida assincronamente...")
                    new_t = await asyncio.get_event_loop().run_in_executor(None, lambda: get_nova_vida_token(force_refresh=True))
                    if new_t:
                        token = new_t
                        log_api_debug("Novo token Nova Vida obtido com sucesso.")
                    else:
                        log_api_debug("Falha ao obter novo token Nova Vida.")
            return token

        if is_api_configured:
            _log("Autenticando na API Nova Vida...")
            token = await ensure_token_valid(force_refresh=False)
            if not token:
                _log("Falha de autenticação na API Nova Vida. Usando fallbacks vazios.")
                is_api_configured = False
            else:
                _log("Autenticação bem-sucedida.")

        endpoint = (NOVA_VIDA_API_URL or "").rstrip("/") + "/" + NOVA_VIDA_API_ENDPOINT.lstrip("/")

        use_cpf      = (not query_fields) or any(normalize_field_name(q) in {"cpf", "documento", "doc", "cnpj"} for q in query_fields)
        use_nome     = (not query_fields) or any(is_required_name_field(q) for q in query_fields)
        use_uf       = (not query_fields) or any(is_uf_field(q) for q in query_fields)
        use_cidade   = (not query_fields) or any(normalize_field_name(q) in {"cidade", "municipio", "nome_cidade"} for q in query_fields)
        use_telefone = (not query_fields) or any(is_required_phone_field(q) for q in query_fields)

        # Limita rajadas ao provedor mesmo quando o .env pede concorrência maior.
        CONCURRENCY = get_enrichment_concurrency()
        log_api_debug(f"Concorrência de enriquecimento configurada: {CONCURRENCY}")

        # Conector TCP calibrado para alta volumetria e reuso eficiente de conexões
        connector = aiohttp.TCPConnector(
            limit=max(20, CONCURRENCY * 2),
            limit_per_host=max(20, CONCURRENCY * 2),
            ttl_dns_cache=600,
            ssl=False,
            keepalive_timeout=75,
            enable_cleanup_closed=True
        )

        # Descobrir todas as chaves (originais + enriquecidas) mantendo colunas originais primeiro
        original_keys = list(dict.fromkeys(key for row in raw_rows for key in row.keys() if not str(key).startswith("enriquecido_")))
        all_keys = list(original_keys)
        for field in PREFIXED_ENRICH_FIELDS:
            if field not in all_keys:
                all_keys.append(field)
        if "_status_enriquecimento" not in all_keys:
            all_keys.append("_status_enriquecimento")

        csv_path = EXPORT_DIR / f"raw_{queue_item_id}.csv"
        with open(csv_path, "w", encoding="utf-8", newline="") as f:
            csv.DictWriter(f, fieldnames=all_keys, delimiter=";").writeheader()

        nvcheck_endpoint = (NOVA_VIDA_API_URL or "").rstrip("/") + "/WSLocalizador.asmx/NVCHECK"
        nvphone_endpoint = (NOVA_VIDA_API_URL or "").rstrip("/") + "/WSLocalizador.asmx/NVCHECK_TELEFONE"
        pesquisar_endpoint = (NOVA_VIDA_API_URL or "").rstrip("/") + "/WSLocalizador.asmx/PESQUISAR"

        # ── Corrotina que enriquece uma única linha ───────────────────────────
        async def enrich_row_async(session: aiohttp.ClientSession, row: Dict[str, Any]) -> Dict[str, Any]:
            if item.get("status") == "cancelled":
                return dict(row)

            # Extração de todos os dados disponíveis da linha da planilha (para enriquecimento e desambiguação inteligente)
            cpf_val      = resolve_cpf_value(row)
            raw_nome     = resolve_name_value(row)
            nome_val     = clean_name_for_api(raw_nome)
            uf_val       = resolve_uf_value(row)
            cidade_val   = clean_name_for_api(resolve_cidade_value(row))
            telefone_val = resolve_phone_value(row)

            if not is_api_configured:
                _inc(False)
                enriched = dict(row)
                for f in PREFIXED_ENRICH_FIELDS:
                    if f not in enriched:
                        enriched[f] = ""
                enriched["_status_enriquecimento"] = "falha"
                return enriched

            async def query_pesquisar_async(
                cpf_str: str = "",
                phone_str: str = "",
                nome_str: str = "",
                uf_str: str = "",
                consulta_id: str = "2421"
            ) -> str | None:
                """Consulta PESQUISAR (consultas 2421 e 2420) com body JSON unificado."""
                clean_c = re.sub(r"\D", "", str(cpf_str or ""))
                clean_p = re.sub(r"\D", "", str(phone_str or ""))
                clean_n = clean_name_for_api(nome_str) if nome_str else ""
                clean_u = normalize_uf_string(uf_str) if uf_str else ""

                if not (clean_c or clean_p or clean_n or clean_u):
                    return None

                payload = {
                    "CPF": clean_c,
                    "TELEFONE": clean_p,
                    "NOME": clean_n,
                    "UF": clean_u
                }

                backoffs = [0.5, 1.0, 2.0]
                for attempt in range(3):
                    cur_token = await ensure_token_valid(force_refresh=False)
                    if not cur_token:
                        return None
                    try:
                        async with session.post(
                            pesquisar_endpoint,
                            data=json.dumps(payload),
                            headers={
                                "Content-Type": "application/json",
                                "Accept": "application/json",
                                "token": sanitize_token(cur_token),
                                "consulta": str(consulta_id)
                            },
                            timeout=aiohttp.ClientTimeout(connect=3.5, sock_read=12.0),
                        ) as resp:
                            _record_api_attempt(f"PESQUISAR_{consulta_id}", f"HTTP_{resp.status}")
                            if resp.status not in (200, 204) and api_attempt_counts.get(f"PESQUISAR_{consulta_id}:HTTP_{resp.status}", 0) <= 2:
                                # Diagnostico: corpo do erro (ASMX costuma informar o motivo do 500)
                                try:
                                    _err_body = (await resp.text())[:800]
                                except Exception as _body_exc:
                                    _err_body = f"<falha ao ler corpo: {_body_exc}>"
                                log_api_debug(
                                    f"[DIAG] PESQUISAR {consulta_id} HTTP {resp.status} | payload_keys={sorted(payload.keys())} "
                                    f"| campos_preenchidos={[k for k, v in payload.items() if v]} | corpo={_err_body!r}"
                                )
                            if resp.status in (429, 500, 502, 503, 504):
                                retry_after = resp.headers.get("Retry-After")
                                try:
                                    sleep_time = min(max(float(retry_after), 0.0), 30.0) if retry_after else backoffs[attempt]
                                except ValueError:
                                    sleep_time = backoffs[attempt]
                                sleep_time += random.uniform(0.05, 0.15)
                                await asyncio.sleep(sleep_time)
                                continue
                            if resp.status == 200:
                                res_text = await resp.text()
                                if is_token_expired_response(res_text):
                                    _record_api_attempt(f"PESQUISAR_{consulta_id}", "TOKEN_EXPIRED")
                                    await ensure_token_valid(force_refresh=True)
                                    continue
                                if res_text and len(res_text.strip()) > 10 and "Nada Consta" not in res_text:
                                    return res_text
                                _record_api_attempt(f"PESQUISAR_{consulta_id}", "EMPTY_OR_NO_MATCH")
                            elif resp.status == 204:
                                return None
                            return None
                    except (aiohttp.ClientError, asyncio.TimeoutError) as exc:
                        outcome = "TIMEOUT" if isinstance(exc, asyncio.TimeoutError) else "NETWORK_ERROR"
                        _record_api_attempt(f"PESQUISAR_{consulta_id}", outcome)
                        if attempt < 2:
                            jitter = random.uniform(0.05, 0.10)
                            sleep_time = (backoffs[attempt] if attempt < len(backoffs) else 0.4) + jitter
                            await asyncio.sleep(sleep_time)
                        else:
                            return None
                return None

            async def query_cpf_async(doc_cpf: str) -> str | None:
                """Consulta PESQUISAR (2421) por CPF com fallback para NVCHECK clássico."""
                clean_cpf = re.sub(r"\D", "", str(doc_cpf or ""))
                if not clean_cpf:
                    return None
                if 9 <= len(clean_cpf) <= 11:
                    clean_cpf = clean_cpf.zfill(11)
                elif len(clean_cpf) != 11:
                    return None

                # 1. Tenta via PESQUISAR 2421
                res = await query_pesquisar_async(cpf_str=clean_cpf, consulta_id="2421")
                if res:
                    return res

                # 2. Fallback via NVCHECK clássico
                backoffs = [0.5, 1.0, 2.0]
                for attempt in range(3):
                    cur_token = await ensure_token_valid(force_refresh=False)
                    if not cur_token:
                        return None
                    data = {"documento": clean_cpf, "token": cur_token}
                    try:
                        async with session.post(
                            nvcheck_endpoint,
                            data=data,
                            headers={"Content-Type": "application/x-www-form-urlencoded"},
                            timeout=aiohttp.ClientTimeout(connect=3.5, sock_read=12.0),
                        ) as resp:
                            _record_api_attempt("NVCHECK", f"HTTP_{resp.status}")
                            if resp.status in (429, 500, 502, 503, 504):
                                retry_after = resp.headers.get("Retry-After")
                                try:
                                    sleep_time = min(max(float(retry_after), 0.0), 30.0) if retry_after else backoffs[attempt]
                                except ValueError:
                                    sleep_time = backoffs[attempt]
                                sleep_time += random.uniform(0.05, 0.15)
                                await asyncio.sleep(sleep_time)
                                continue
                            if resp.status == 200:
                                res_text = await resp.text()
                                if is_token_expired_response(res_text):
                                    _record_api_attempt("NVCHECK", "TOKEN_EXPIRED")
                                    await ensure_token_valid(force_refresh=True)
                                    continue
                                return res_text
                            return None
                    except (aiohttp.ClientError, asyncio.TimeoutError) as exc:
                        outcome = "TIMEOUT" if isinstance(exc, asyncio.TimeoutError) else "NETWORK_ERROR"
                        _record_api_attempt("NVCHECK", outcome)
                        if attempt < 2:
                            jitter = random.uniform(0.05, 0.10)
                            sleep_time = (backoffs[attempt] if attempt < len(backoffs) else 0.4) + jitter
                            await asyncio.sleep(sleep_time)
                        else:
                            return None
                return None

            async def query_phone_async(phone_str: str) -> str | None:
                """Consulta PESQUISAR (2421) para descobrir CPF e dados pelo telefone, com fallback para NVCHECK_TELEFONE."""
                digits = re.sub(r"\D", "", str(phone_str or ""))
                if len(digits) < 10 or len(digits) > 11:
                    return None

                # 1. Tenta via PESQUISAR 2421
                res = await query_pesquisar_async(phone_str=digits, consulta_id="2421")
                if res:
                    return res

                # 2. Fallback via NVCHECK_TELEFONE clássico
                backoffs = [0.5, 1.0, 2.0]
                for attempt in range(2):
                    cur_token = await ensure_token_valid(force_refresh=False)
                    if not cur_token:
                        return None
                    data = {"telefone": digits, "tipoDocumento": "PF", "token": cur_token}
                    try:
                        async with session.post(
                            nvphone_endpoint,
                            data=data,
                            headers={"Content-Type": "application/x-www-form-urlencoded"},
                            timeout=aiohttp.ClientTimeout(connect=3.5, sock_read=12.0),
                        ) as resp:
                            _record_api_attempt("NVCHECK_TELEFONE", f"HTTP_{resp.status}")
                            if resp.status in (429, 500, 502, 503, 504):
                                retry_after = resp.headers.get("Retry-After")
                                try:
                                    sleep_time = min(max(float(retry_after), 0.0), 30.0) if retry_after else backoffs[attempt]
                                except ValueError:
                                    sleep_time = backoffs[attempt]
                                sleep_time += random.uniform(0.05, 0.15)
                                await asyncio.sleep(sleep_time)
                                continue
                            if resp.status == 200:
                                res_text = await resp.text()
                                if is_token_expired_response(res_text):
                                    _record_api_attempt("NVCHECK_TELEFONE", "TOKEN_EXPIRED")
                                    await ensure_token_valid(force_refresh=True)
                                    continue
                                if res_text and len(res_text.strip()) > 30 and "Nada Consta" not in res_text:
                                    return res_text
                                _record_api_attempt("NVCHECK_TELEFONE", "EMPTY_OR_NO_MATCH")
                            return None
                    except (aiohttp.ClientError, asyncio.TimeoutError) as exc:
                        outcome = "TIMEOUT" if isinstance(exc, asyncio.TimeoutError) else "NETWORK_ERROR"
                        _record_api_attempt("NVCHECK_TELEFONE", outcome)
                        if attempt < 1:
                            jitter = random.uniform(0.05, 0.10)
                            sleep_time = (backoffs[attempt] if attempt < len(backoffs) else 0.4) + jitter
                            await asyncio.sleep(sleep_time)
                        else:
                            return None
                return None

            async def query_name_async(query_nome: str, uf_filter: str = "") -> str | None:
                """Consulta PESQUISAR por Nome e Nome+UF (2421 com fallback para 2420)."""
                if not query_nome:
                    return None

                clean_q_nome = clean_name_for_api(query_nome)
                clean_uf = normalize_uf_string(uf_filter) if uf_filter else ""

                # 1. Tenta com Nome + UF em 2421 (Completo)
                if clean_uf:
                    res = await query_pesquisar_async(nome_str=clean_q_nome, uf_str=clean_uf, consulta_id="2421")
                    if res:
                        return res
                    # Tenta com 2420 (Simplificado)
                    res = await query_pesquisar_async(nome_str=clean_q_nome, uf_str=clean_uf, consulta_id="2420")
                    if res:
                        return res

                # 2. Busca nacional por Nome em 2421
                res = await query_pesquisar_async(nome_str=clean_q_nome, uf_str="", consulta_id="2421")
                if res:
                    return res

                # 3. Busca nacional por Nome em 2420
                res = await query_pesquisar_async(nome_str=clean_q_nome, uf_str="", consulta_id="2420")
                if res:
                    return res

                return None

            try:
                target_name = raw_nome or nome_val
                digits_phone_clean = re.sub(r"\D", "", telefone_val)

                final_cpf = None
                final_nome = None
                final_score = 0.0
                final_level = "DIVERGENTE"
                final_parsed = None

                # 1. Busca por Nome (com UF e fallback nacional se for 1 pessoa única)
                if nome_val and use_nome:
                    json_res = await query_name_async(nome_val, uf_filter=uf_val if use_uf else "")
                    if json_res:
                        ok, c_cpf, c_nome, c_score, c_level, c_parsed = parse_and_validate_xml_for_target(
                            json_res, target_name, target_uf=uf_val if use_uf else ""
                        )
                        if ok:
                            final_cpf, final_nome, final_score, final_level, final_parsed = c_cpf, c_nome, c_score, c_level, c_parsed

                # 2. Fallback de Nome: Variações fonéticas/aglutinadas de nome composto
                if not final_parsed and nome_val and use_nome:
                    variants = generate_name_query_variants(nome_val)
                    for variant in variants:
                        if variant != nome_val:
                            v_uf = uf_val if (use_uf and uf_val) else ""
                            json_res = await query_name_async(variant, uf_filter=v_uf)
                            if json_res:
                                ok, c_cpf, c_nome, c_score, c_level, c_parsed = parse_and_validate_xml_for_target(
                                    json_res, target_name, target_uf=v_uf
                                )
                                if ok:
                                    final_cpf, final_nome, final_score, final_level, final_parsed = c_cpf, c_nome, c_score, c_level, c_parsed
                                    break

                # 3. Se não achou por Nome ou se foi homônimo (204), e a planilha possui TELEFONE: busca via NVCHECK_TELEFONE
                if not final_parsed and digits_phone_clean:
                    xml_res = await query_phone_async(digits_phone_clean)
                    if xml_res:
                        ok, c_cpf, c_nome, c_score, c_level, c_parsed = parse_and_validate_xml_for_target(
                            xml_res, target_name, is_cpf_query=not bool(target_name)
                        )
                        if ok:
                            final_cpf, final_nome, final_score, final_level, final_parsed = c_cpf, c_nome, c_score, c_level, c_parsed

                # 4. Se não achou por Nome/Telefone e a planilha possui CPF: busca via NVCHECK
                if not final_parsed and cpf_val:
                    xml_res = await query_cpf_async(cpf_val)
                    if xml_res:
                        ok, c_cpf, c_nome, c_score, c_level, c_parsed = parse_and_validate_xml_for_target(
                            xml_res, target_name, is_cpf_query=True
                        )
                        if ok:
                            final_cpf, final_nome, final_score, final_level, final_parsed = c_cpf, c_nome, c_score, c_level, c_parsed

                # 6. Se não achou nenhum registro compatível (0 registros ou 2+ registros descartados)
                if not final_parsed:
                    _record_api_attempt("ROWS", "NO_MATCH")
                    raise ValueError("Não localizado ou titular divergente")

                # Se a busca retornou apenas dados simplificados (sem telefone/endereço) e temos CPF, complementa via 2421
                if final_cpf and not final_parsed.get("telefone") and not final_parsed.get("logradouro"):
                    nv_res = await query_cpf_async(final_cpf)
                    if nv_res:
                        parsed_360 = parse_nova_vida_full_xml(nv_res)
                        if parsed_360:
                            for k_p, v_p in parsed_360.items():
                                if v_p and not final_parsed.get(k_p):
                                    final_parsed[k_p] = v_p

                if target_name:
                    similaridade_str = f"{final_level} ({final_score:.0f}%)"
                else:
                    similaridade_str = "NAO_APLICAVEL"

                enriched = dict(row)
                enriched.update({
                    "enriquecido_nome": normalize_name_value(final_parsed.get("nome") or raw_nome or nome_val),
                    "enriquecido_telefone": final_parsed.get("telefone") or re.sub(r"\D", "", telefone_val),
                    "enriquecido_telefone_2": final_parsed.get("telefone_2") or "",
                    "enriquecido_telefone_3": final_parsed.get("telefone_3") or "",
                    "enriquecido_telefone_4": final_parsed.get("telefone_4") or "",
                    "enriquecido_telefone_5": final_parsed.get("telefone_5") or "",
                    "enriquecido_cpf": final_parsed.get("cpf") or cpf_val,
                    "enriquecido_data_nascimento": final_parsed.get("nascimento") or "",
                    "enriquecido_idade": final_parsed.get("idade") or "",
                    "enriquecido_sexo": final_parsed.get("sexo") or "",
                    "enriquecido_email": final_parsed.get("email") or "",
                    "enriquecido_logradouro": final_parsed.get("logradouro") or "",
                    "enriquecido_numero": final_parsed.get("numero") or "",
                    "enriquecido_complemento": final_parsed.get("complemento") or "",
                    "enriquecido_bairro": final_parsed.get("bairro") or "",
                    "enriquecido_cidade": final_parsed.get("cidade") or cidade_val,
                    "enriquecido_estado": final_parsed.get("estado") or uf_val,
                    "enriquecido_cep": final_parsed.get("cep") or "",
                    "enriquecido_renda_estimada": final_parsed.get("renda_estimada") or "",
                    "enriquecido_cbo": final_parsed.get("cbo") or "",
                    "enriquecido_profissao": final_parsed.get("profissao") or "",
                    "enriquecido_score": final_parsed.get("score") or "",
                    "enriquecido_perfil": final_parsed.get("perfil") or "",
                    "enriquecido_obito": final_parsed.get("obito") or "0",
                    "enriquecido_whatsapp": final_parsed.get("whatsapp") or "NÃO",
                    "enriquecido_status": final_parsed.get("status") or "REGULAR",
                    "enriquecido_similaridade_nome": similaridade_str,
                    "tipo_enriquecimento": detected_type,
                    "_status_enriquecimento": "sucesso",
                })

                for k, val in row.items():
                    norm_k = normalize_text(k)
                    if norm_k and norm_k not in enriched:
                        enriched[norm_k] = val

                _inc(True)
                if counters["processed"] % 25 == 0 or counters["processed"] == total_rows:
                    _log(f"Progresso: [{counters['processed']}/{total_rows}] Sucessos: {counters['success']} | Falhas: {counters['error']}")
                return enriched

            except Exception as exc:
                # Diagnostico: distingue "nao localizado" legitimo de erro de codigo/rede escondido
                _exc_key = f"{type(exc).__name__}:{str(exc)[:90]}"
                _record_api_attempt("EXCECAO", _exc_key)
                if api_attempt_counts.get(f"EXCECAO:{_exc_key}", 0) == 1:
                    import traceback as _tb
                    log_api_debug(f"[DIAG] Primeira ocorrencia de excecao no enriquecimento: {_exc_key}\n{_tb.format_exc()}")
                phone_final = re.sub(r"\D", "", telefone_val)
                enriched = dict(row)
                enriched.update({
                    "enriquecido_nome": normalize_name_value(raw_nome or nome_val),
                    "enriquecido_telefone": phone_final,
                    "enriquecido_telefone_2": "",
                    "enriquecido_telefone_3": "",
                    "enriquecido_telefone_4": "",
                    "enriquecido_telefone_5": "",
                    "enriquecido_cpf": cpf_val or "",
                    "enriquecido_data_nascimento": "",
                    "enriquecido_idade": "",
                    "enriquecido_sexo": "",
                    "enriquecido_email": "",
                    "enriquecido_logradouro": "",
                    "enriquecido_numero": "",
                    "enriquecido_complemento": "",
                    "enriquecido_bairro": "",
                    "enriquecido_cidade": cidade_val,
                    "enriquecido_estado": uf_val,
                    "enriquecido_cep": "",
                    "enriquecido_renda_estimada": "",
                    "enriquecido_cbo": "",
                    "enriquecido_profissao": "",
                    "enriquecido_score": "",
                    "enriquecido_perfil": "",
                    "enriquecido_obito": "",
                    "enriquecido_whatsapp": "",
                    "enriquecido_status": "",
                    "enriquecido_similaridade_nome": "DIVERGENTE (0%)" if (raw_nome or nome_val) else "",
                    "tipo_enriquecimento": detected_type,
                    "_status_enriquecimento": "falha",
                })
                for k, val in row.items():
                    norm_k = normalize_text(k)
                    if norm_k and norm_k not in enriched:
                        enriched[norm_k] = val

                _inc(False)
                _log(f"[{counters['processed']}/{total_rows}] Não localizado: {raw_nome or nome_val}")
                return enriched

        FLUSH_BATCH = 500
        CHUNK_SIZE = int(os.environ.get("ENRICHMENT_CHUNK_SIZE", "5000"))
        total_chunks = max(1, math.ceil(total_rows / CHUNK_SIZE)) if total_rows > 0 else 1
        is_chunked = total_chunks > 1

        if is_chunked:
            _log(f"Base massiva detectada ({total_rows} registros). Particionamento ativo: dividindo em {total_chunks} sub-lotes de até {CHUNK_SIZE} registros para estabilidade e controle de memória.")
            item["chunk_info"] = {
                "is_chunked": True,
                "total_chunks": total_chunks,
                "current_chunk": 1,
                "chunk_size": CHUNK_SIZE,
                "total_rows": total_rows
            }
            item["current_chunk"] = 1
            item["total_chunks"] = total_chunks
            db_update_queue_item(queue_item_id, {
                "chunk_info": item["chunk_info"],
                "current_chunk": 1,
                "total_chunks": total_chunks,
                "logs": logs_list
            })

        async with aiohttp.ClientSession(connector=connector) as session:
            for chunk_idx in range(total_chunks):
                if item.get("status") == "cancelled":
                    break

                chunk_num = chunk_idx + 1
                start_idx = chunk_idx * CHUNK_SIZE
                end_idx = min(start_idx + CHUNK_SIZE, total_rows)
                chunk_raw_rows = raw_rows[start_idx:end_idx]
                chunk_len = len(chunk_raw_rows)

                if is_chunked:
                    item["current_chunk"] = chunk_num
                    item["chunk_info"]["current_chunk"] = chunk_num
                    item["chunk_info"]["chunk_start"] = start_idx
                    item["chunk_info"]["chunk_end"] = end_idx
                    item["chunk_info"]["chunk_rows"] = chunk_len
                    _log(f"[Sub-lote {chunk_num}/{total_chunks}] Iniciando enriquecimento de {chunk_len} registros (linhas {start_idx + 1} a {end_idx})...")
                    db_update_queue_item(queue_item_id, {
                        "chunk_info": item["chunk_info"],
                        "current_chunk": chunk_num,
                        "logs": logs_list
                    })

                chunk_results: List[Dict[str, Any] | None] = [None] * chunk_len
                chunk_next_flush_idx = 0
                chunk_flush_lock = asyncio.Lock()

                async def flush_ready_to_csv(force: bool = False) -> None:
                    nonlocal chunk_next_flush_idx
                    to_write: List[Dict[str, Any]] = []
                    async with chunk_flush_lock:
                        while chunk_next_flush_idx < chunk_len and chunk_results[chunk_next_flush_idx] is not None:
                            row_res = chunk_results[chunk_next_flush_idx]
                            chunk_results[chunk_next_flush_idx] = None  # Libera memória imediatamente
                            to_write.append(row_res)
                            chunk_next_flush_idx += 1
                            if not force and len(to_write) >= FLUSH_BATCH:
                                break

                    if to_write:
                        rows_to_write = [{k: r.get(k, "") for k in all_keys} for r in to_write]
                        # Grava no CSV principal consolidado
                        with open(csv_path, "a", encoding="utf-8", newline="") as f:
                            csv.DictWriter(f, fieldnames=all_keys, delimiter=";").writerows(rows_to_write)
                        # Se for particionado, grava também a cópia individual da parte para redundância e download de fatia
                        if is_chunked:
                            part_csv_path = EXPORT_DIR / f"raw_{queue_item_id}_part{chunk_num}.csv"
                            is_new_part = not part_csv_path.exists()
                            with open(part_csv_path, "a", encoding="utf-8", newline="") as pf:
                                pwriter = csv.DictWriter(pf, fieldnames=all_keys, delimiter=";")
                                if is_new_part:
                                    pwriter.writeheader()
                                pwriter.writerows(rows_to_write)
                        _flush_ui_logs()

                async def process_row(rel_idx: int, row: Dict[str, Any]) -> None:
                    if item.get("status") == "cancelled":
                        return

                    result = await enrich_row_async(session, row)
                    if result is None:
                        result = dict(row)
                    chunk_results[rel_idx] = result
                    if not item.get("preview_row") and result.get("_status_enriquecimento") == "sucesso":
                        item["preview_row"] = result

                # Flusher em segundo plano para não travar os workers durante I/O de disco
                async def periodic_csv_flusher():
                    while item.get("status") not in ("cancelled", "completed", "error"):
                        await asyncio.sleep(0.15)
                        await flush_ready_to_csv()

                # Fila de processamento exclusiva deste sub-lote (baixo uso de RAM)
                work_queue: asyncio.Queue = asyncio.Queue()
                for rel_idx, r in enumerate(chunk_raw_rows):
                    work_queue.put_nowait((rel_idx, r))

                async def worker_loop():
                    while not work_queue.empty():
                        if item.get("status") == "cancelled":
                            break
                        try:
                            rel_idx, row = work_queue.get_nowait()
                        except asyncio.QueueEmpty:
                            break
                        try:
                            await process_row(rel_idx, row)
                        finally:
                            work_queue.task_done()

                flusher_task = asyncio.create_task(periodic_csv_flusher())

                num_workers = min(CONCURRENCY, chunk_len) if chunk_len > 0 else 1
                workers = [asyncio.create_task(worker_loop()) for _ in range(num_workers)]
                await asyncio.gather(*workers)

                flusher_task.cancel()

                # Flush final garantido do sub-lote
                if item.get("status") != "cancelled":
                    await flush_ready_to_csv(force=True)

                _flush_ui_logs()

                if is_chunked and item.get("status") != "cancelled":
                    _log(f"[Sub-lote {chunk_num}/{total_chunks}] Concluído com sucesso ({chunk_len} registros integrados ao arquivo consolidado).")
                    db_update_queue_item(queue_item_id, {
                        "processed_count": counters["processed"],
                        "success_count": counters["success"],
                        "error_count": counters["error"],
                        "current_chunk": chunk_num,
                        "chunk_info": item["chunk_info"],
                        "logs": logs_list,
                    })

                # Limpeza explícita de memória do sub-lote
                chunk_results = None
                chunk_raw_rows = None

        if item.get("status") == "cancelled":
            logs_list.append("Processamento cancelado pelo usuário.")
            db_update_queue_item(queue_item_id, {"status": "cancelled", "logs": logs_list})
            return

        # Recombinação e validação final
        if is_chunked:
            line_count = 0
            if csv_path.exists():
                with open(csv_path, "r", encoding="utf-8") as f:
                    line_count = max(0, sum(1 for _ in f) - 1)
            _log(f"Recombinação concluída com sucesso: Todos os {total_chunks} sub-lotes foram unificados no arquivo mestre consolidado ({line_count} registros validados).")
            log_api_debug(f"[{item.get('queue_number', '')}] Recombinação final: {line_count}/{total_rows} registros consolidados em {csv_path.name}")

        # Conclusão
        item["status"] = "pending"
        item["total_rows_enriched"] = counters["success"]
        logs_list.append(f"Processamento concluído. {counters['success']} registros enriquecidos com sucesso, {counters['error']} falhas.")
        log_api_debug(f"[{item.get('queue_number', '')}] Motor assíncrono concluído: {counters['success']} sucessos, {counters['error']} falhas.")
        if api_attempt_counts:
            summary = ", ".join(f"{key}={count}" for key, count in sorted(api_attempt_counts.items()))
            log_api_debug(f"[{item.get('queue_number', '')}] Respostas da API Nova Vida: {summary}")

        db_update_queue_item(queue_item_id, {
            "status": "pending",
            "processed_count":    counters["processed"],
            "success_count":      counters["success"],
            "error_count":        counters["error"],
            "total_rows_enriched": counters["success"],
            "logs": logs_list,
        })

    except Exception as general_exc:
        if item is not None:
            item["status"] = "error"
        logs_list.append(f"Erro catastrófico no processamento: {general_exc}")
        log_api_debug(f"[{item.get('queue_number', '') if item else queue_item_id}] Erro catastrófico no processamento: {general_exc}")
        db_update_queue_item(queue_item_id, {"status": "error", "logs": logs_list})



class EnrichmentQueueManager:
    """Gerenciador de Fila Serializada (FIFO) para enriquecimento de dados Nova Vida.
    Garante que apenas 1 lote/solicitação seja processado por vez contra a API externa,
    enquanto outras solicitações aguardam ordenadamente no status 'queued' para não sobrecarregar o provedor.
    """
    def __init__(self):
        self._lock = threading.RLock()
        self._job_queue = queue.Queue()
        self._active_job_id = None
        self._worker_thread = None
        self._stop_event = threading.Event()
        self._queued_job_ids = []

    def start(self):
        with self._lock:
            if self._worker_thread is None or not self._worker_thread.is_alive():
                self._stop_event.clear()
                self._worker_thread = threading.Thread(target=self._worker_loop, daemon=True, name="EnrichmentQueueWorker")
                self._worker_thread.start()
                log_api_debug("[QueueManager] Worker global de enriquecimento iniciado.")

    def enqueue(self, queue_item_id: str, raw_rows: List[Dict[str, Any]], detected_type: str, query_fields: List[str]):
        self.start()
        with self._lock:
            if queue_item_id not in self._queued_job_ids and queue_item_id != self._active_job_id:
                self._queued_job_ids.append(queue_item_id)
            
            job_data = {
                "id": queue_item_id,
                "raw_rows": raw_rows,
                "detected_type": detected_type,
                "query_fields": query_fields,
            }
            self._job_queue.put(job_data)
            pos = self.get_position(queue_item_id)
            log_api_debug(f"[QueueManager] Solicitacao {queue_item_id} enfileirada. Posicao na fila: #{pos}")

    def get_position(self, queue_item_id: str) -> int:
        with self._lock:
            if self._active_job_id == queue_item_id:
                return 0  # Em processamento ativo
            try:
                idx = self._queued_job_ids.index(queue_item_id)
                return idx + 1
            except ValueError:
                return 1

    def get_total_queued(self) -> int:
        with self._lock:
            return len(self._queued_job_ids)

    def is_active(self, queue_item_id: str) -> bool:
        with self._lock:
            return self._active_job_id == queue_item_id

    def cancel_job(self, queue_item_id: str) -> bool:
        with self._lock:
            if queue_item_id in self._queued_job_ids:
                self._queued_job_ids.remove(queue_item_id)
                log_api_debug(f"[QueueManager] Solicitacao {queue_item_id} removida da fila de espera.")
                return True
        return False

    def _worker_loop(self):
        while not self._stop_event.is_set():
            try:
                job_data = self._job_queue.get(timeout=1.0)
            except queue.Empty:
                continue

            job_id = job_data["id"]

            # Verificar se a solicitação foi cancelada enquanto aguardava na fila
            with self._lock:
                if job_id in self._queued_job_ids:
                    self._queued_job_ids.remove(job_id)
                item = db_get_queue_item(job_id)
                if not item or item.get("status") == "cancelled":
                    self._job_queue.task_done()
                    continue
                self._active_job_id = job_id
                item["status"] = "processing"
                logs = item.setdefault("logs", [])
                logs.append("Chegou a sua vez na fila! Iniciando consultas com a API Nova Vida...")
                db_update_queue_item(job_id, {"status": "processing", "logs": logs})

            # Executa o motor assíncrono para esta solicitação
            try:
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                try:
                    loop.run_until_complete(
                        _async_enrichment_engine(
                            job_data["id"],
                            job_data["raw_rows"],
                            job_data["detected_type"],
                            job_data["query_fields"]
                        )
                    )
                finally:
                    loop.close()
            except Exception as exc:
                log_api_debug(f"[QueueManager] Erro fatal no job {job_id}: {exc}")
                try:
                    item = db_get_queue_item(job_id)
                    if item:
                        item.setdefault("logs", []).append(f"Erro fatal no processamento: {exc}")
                        db_update_queue_item(job_id, {"status": "error", "logs": item["logs"]})
                except Exception:
                    pass
            finally:
                with self._lock:
                    self._active_job_id = None
                self._job_queue.task_done()


ENRICHMENT_QUEUE_MANAGER = EnrichmentQueueManager()
ENRICHMENT_QUEUE_MANAGER.start()


def process_enrichment_background_task(queue_item_id: str, raw_rows: List[Dict[str, Any]], detected_type: str, query_fields: List[str]):
    """Enfileira a solicitação no Gerenciador de Fila Serializada (FIFO)."""
    ENRICHMENT_QUEUE_MANAGER.enqueue(queue_item_id, raw_rows, detected_type, query_fields)


def _recover_stale_queue_on_startup():
    """Garante que a fila inicie limpa, cancelando jobs interrompidos por paradas anteriores."""
    try:
        items = db_get_all_queue()
        stale_ids = [it["id"] for it in items if isinstance(it, dict) and it.get("status") in ("queued", "processing")]
        time.sleep(0.5)
        for job_id in stale_ids:
            it = db_get_queue_item(job_id)
            if it and it.get("status") in ("queued", "processing"):
                logs = it.setdefault("logs", [])
                logs.append("Servidor reiniciado. Processamento anterior finalizado/cancelado.")
                db_update_queue_item(job_id, {"status": "cancelled", "logs": logs})
                log_api_debug(f"[QueueManager] Solicitacao anterior {job_id} marcada como cancelada por reinicializacao.")
    except Exception as e:
        log_api_debug(f"Aviso na limpeza de fila na inicialização: {e}")

threading.Thread(target=_recover_stale_queue_on_startup, daemon=True).start()


def _async_queue_processor(
    queue_item_id: str,
    raw_rows: List[Dict[str, Any]] = None,
    record_id: str = None,
    selected_fields: List[str] = None,
    queue_number: str = "",
    requester_name: str = "",
    deduplicate_flag: bool = True
):
    """Motor assincrono de particionamento e formatacao de base em sub-lotes para Subir Fila."""
    try:
        item = db_get_queue_item(queue_item_id)
        if not item or item.get("status") == "cancelled":
            return

        logs_list = item.setdefault("logs", [])

        # 1. Recupera as linhas brutas se nao foram fornecidas em memoria
        if not raw_rows:
            effective_rec_id = record_id or item.get("record_id")
            record = db_get_record(effective_rec_id) if effective_rec_id else None
            if record:
                raw_rows = record.get("rows")
                if (not raw_rows or len(raw_rows) == 0) and record.get("stored_file"):
                    stored_file_path = Path(record["stored_file"])
                    if stored_file_path.exists():
                        extension = stored_file_path.suffix.lower()
                        logs_list.append("Carregando e validando dados da planilha no plano de fundo...")
                        db_update_queue_item(queue_item_id, {"logs": logs_list[-300:]})
                        try:
                            file_bytes = stored_file_path.read_bytes()
                            if extension == ".xlsx":
                                raw_rows = read_excel_rows(file_bytes)
                            elif extension == ".csv":
                                raw_rows = read_csv_rows(file_bytes)
                            if raw_rows:
                                raw_rows = clean_dataframe(raw_rows)
                        except Exception as e:
                            log_api_debug(f"Erro ao carregar linhas do arquivo original {stored_file_path}: {e}")

        if not raw_rows:
            logs_list.append("Erro: Nenhum registro recuperado para processar.")
            db_update_queue_item(queue_item_id, {"status": "error", "logs": logs_list[-300:]})
            return

        # 2. Desduplicacao inteligente
        if deduplicate_flag:
            raw_rows, dup_count = deduplicate_rows(raw_rows)
            if dup_count > 0:
                logs_list.append(f"Desduplicacao ativa: {dup_count} registros duplicados removidos. {len(raw_rows)} contatos unicos prontos.")
                db_update_queue_item(queue_item_id, {"logs": logs_list[-300:], "total_rows": len(raw_rows)})

        total_rows = len(raw_rows)
        CHUNK_SIZE = int(os.environ.get("QUEUE_CHUNK_SIZE", os.environ.get("ENRICHMENT_CHUNK_SIZE", "5000")))
        total_chunks = max(1, math.ceil(total_rows / CHUNK_SIZE)) if total_rows > 0 else 1
        is_chunked = total_chunks > 1

        chunk_info = {
            "is_chunked": is_chunked,
            "total_chunks": total_chunks,
            "current_chunk": 1,
            "chunk_size": CHUNK_SIZE,
            "total_rows": total_rows
        }

        if is_chunked:
            logs_list.append(f"Base massiva detectada ({total_rows} registros). Particionamento ativo: dividindo em {total_chunks} sub-lotes de ate {CHUNK_SIZE} registros para estabilidade e controle de memoria.")
        else:
            logs_list.append(f"Iniciando formatacao e validacao de {total_rows} contatos para importacao na fila.")

        db_update_queue_item(queue_item_id, {
            "status": "processing",
            "total_rows": total_rows,
            "chunk_info": chunk_info,
            "current_chunk": 1,
            "total_chunks": total_chunks,
            "logs": logs_list[-300:]
        })

        effective_fields = normalize_selected_fields(selected_fields or item.get("selected_fields"))
        if not effective_fields:
            original_columns = []
            rec = db_get_record(record_id or item.get("record_id", ""))
            if rec:
                original_columns = rec.get("columns") or []
            effective_fields = original_columns or ["Nome", "Telefone"]

        concatenator = FastConcatenator(effective_fields)
        clean_queue = str(queue_number or item.get("queue_number") or "importar").strip()
        if clean_queue == "TEMP":
            clean_queue = "importar"

        original_keys = list(dict.fromkeys(key for row in raw_rows for key in row.keys())) if raw_rows else []
        fail_cols = [c for c in original_keys if not c.startswith("enriquecido_") and c not in ("_status_enriquecimento", "tipo_enriquecimento")]
        if not fail_cols:
            fail_cols = list(original_keys) if original_keys else ["Falha"]

        csv_path = EXPORT_DIR / f"raw_{queue_item_id}.csv"
        with open(csv_path, "w", encoding="utf-8", newline="") as f:
            csv.DictWriter(f, fieldnames=original_keys, delimiter=";").writeheader()

        all_success_rows = []
        all_failed_rows = []
        processed_count = 0
        success_count = 0
        error_count = 0

        for chunk_idx in range(total_chunks):
            current_item = db_get_queue_item(queue_item_id)
            if not current_item or current_item.get("status") == "cancelled":
                log_api_debug(f"[QueueProcessor] Job {queue_item_id} cancelado pelo usuario.")
                return

            chunk_num = chunk_idx + 1
            start_idx = chunk_idx * CHUNK_SIZE
            end_idx = min(start_idx + CHUNK_SIZE, total_rows)
            chunk_raw = raw_rows[start_idx:end_idx]
            chunk_len = len(chunk_raw)

            chunk_info["current_chunk"] = chunk_num
            if is_chunked:
                logs_list.append(f"[Sub-lote {chunk_num}/{total_chunks}] Formatando e validando {chunk_len} contatos (linhas {start_idx + 1} a {end_idx})...")

            chunk_success = []
            chunk_failed = []
            for row in chunk_raw:
                if is_valid_queue_row(row):
                    chunk_success.append({"importar": concatenator.format(row)})
                    success_count += 1
                else:
                    chunk_failed.append(dict(row))
                    error_count += 1
                processed_count += 1

            all_success_rows.extend(chunk_success)
            all_failed_rows.extend(chunk_failed)

            if is_chunked:
                # 1. Fatia bruta
                part_csv_path = EXPORT_DIR / f"raw_{queue_item_id}_part{chunk_num}.csv"
                with open(part_csv_path, "w", encoding="utf-8", newline="") as pf:
                    pwriter = csv.DictWriter(pf, fieldnames=original_keys, delimiter=";")
                    pwriter.writeheader()
                    pwriter.writerows(chunk_raw)

                # 2. Fatia tratada pronta para importar
                part_treated_path = EXPORT_DIR / f"fila_{clean_queue}_{queue_item_id}_part{chunk_num}.csv"
                with open(part_treated_path, "w", encoding="utf-8", newline="") as ptf:
                    pt_writer = csv.DictWriter(ptf, fieldnames=["importar"], delimiter=";")
                    pt_writer.writeheader()
                    pt_writer.writerows(chunk_success)

            with open(csv_path, "a", encoding="utf-8", newline="") as f:
                cwriter = csv.DictWriter(f, fieldnames=original_keys, delimiter=";")
                cwriter.writerows(chunk_raw)

            db_update_queue_item(queue_item_id, {
                "processed_count": processed_count,
                "success_count": success_count,
                "error_count": error_count,
                "current_chunk": chunk_num,
                "chunk_info": chunk_info,
                "logs": logs_list[-300:]
            })

        # Pre-geracao do Excel final sob lock
        filepath = EXPORT_DIR / f"fila_{clean_queue}_{queue_item_id}.xlsx"
        sheets = [
            {"title": "Importar", "columns": ["importar"], "rows": all_success_rows},
            {"title": "Falha", "columns": fail_cols, "rows": all_failed_rows}
        ]
        with _get_export_lock(queue_item_id):
            write_export_workbook(filepath=filepath, sheets=sheets)

        completed_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        logs_list.append(f"Formatacao da fila concluida com sucesso! Todos os {total_chunks} sub-lotes foram unificados no arquivo mestre ({processed_count} registros validados: {success_count} aptos para importacao, {error_count} falhas).")

        final_requester = requester_name or item.get("requester_name") or "Sistema"
        db_update_queue_item(queue_item_id, {
            "status": "completed",
            "completed_at": completed_time,
            "processed_count": processed_count,
            "success_count": success_count,
            "error_count": error_count,
            "queue_number": clean_queue,
            "selected_fields": effective_fields,
            "logs": logs_list[-300:]
        })

        db_save_audit_event("QUEUE_FINALIZED", final_requester, {
            "request_id": queue_item_id,
            "queue_number": clean_queue,
            "mode": "queue_only",
            "total_rows": total_rows,
            "success_count": success_count,
            "error_count": error_count,
            "chunks_count": total_chunks,
        })

        if clean_queue != "importar":
            db_save_notification({
                "id": str(uuid.uuid4()),
                "title": f"Fila {clean_queue} tratada e gerada",
                "message": f"A fila {clean_queue} foi tratada e concatenada para {final_requester} ({total_rows} registros em {total_chunks} sub-lotes).",
                "time": completed_time,
            })

    except Exception as exc:
        log_api_debug(f"[QueueProcessor] Erro no processamento do lote {queue_item_id}: {exc}")
        try:
            it = db_get_queue_item(queue_item_id)
            if it:
                it.setdefault("logs", []).append(f"Erro no processamento da fila: {exc}")
                db_update_queue_item(queue_item_id, {"status": "error", "logs": it["logs"][-300:]})
        except Exception:
            pass


@app.post("/api/request-import")
@login_required
def create_import_request():
    payload = request.get_json(silent=True) or {}
    record_id = payload.get("record_id")
    query_fields = payload.get("query_fields", [])
    mode = (payload.get("mode") or "enrich_and_queue").strip()
    deduplicate_flag = payload.get("deduplicate", True)

    if not record_id:
        return jsonify({"error": "ID do registro não informado."}), 400

    record = db_get_record(record_id)
    if record is None:
        return jsonify({"error": "Registro não encontrado."}), 404

    raw_rows = record.get("rows")
    stored_file_path = Path(record["stored_file"]) if record.get("stored_file") else None

    # No modo queue_only, se nao temos raw_rows em memoria, evitamos parse sincrono bloqueante
    # para nao esgotar o timeout do proxy/tunnel (previne erro 503). O worker carrega em background.
    if mode == "queue_only":
        if (not raw_rows or len(raw_rows) == 0) and (not stored_file_path or not stored_file_path.exists()):
            return jsonify({"error": "Erro ao recuperar dados para processamento (arquivo não encontrado)."}), 400
    else:
        if (not raw_rows or len(raw_rows) == 0) and stored_file_path and stored_file_path.exists():
            extension = stored_file_path.suffix.lower()
            try:
                file_bytes = stored_file_path.read_bytes()
                if extension == ".xlsx":
                    raw_rows = read_excel_rows(file_bytes)
                elif extension == ".csv":
                    raw_rows = read_csv_rows(file_bytes)
                if raw_rows:
                    raw_rows = clean_dataframe(raw_rows)
            except Exception as e:
                log_api_debug(f"Erro ao carregar linhas do arquivo original {stored_file_path}: {e}")

        if not raw_rows:
            return jsonify({"error": "Erro ao recuperar dados para processamento (linhas vazias)."}), 400

    # Aplicacao de desduplicacao previa inteligente (apenas se raw_rows ja esta em memoria)
    dup_count = 0
    if deduplicate_flag and raw_rows:
        raw_rows, dup_count = deduplicate_rows(raw_rows)

    requester_name = payload.get("requester_name", "").strip()
    current_user_name = session.get("user_name", "Usuário")
    
    if requester_name:
        final_requester_name = f"{current_user_name} ({requester_name})"
    else:
        final_requester_name = "Aguardando confirmação" if mode == "queue_only" else current_user_name

    initial_logs = []
    if dup_count > 0:
        initial_logs.append(f"Desduplicação ativada: {dup_count} registros duplicados removidos. {len(raw_rows)} registros únicos prontos.")
    if mode != "queue_only":
        initial_logs.append("Solicitação inserida na fila global de processamento. Aguardando sua vez...")
    else:
        initial_logs.append("Fila recebida. Processando contatos em sub-lotes em segundo plano...")

    export_format = (payload.get("export_format") or "original").strip().lower()

    if mode == "queue_only":
        queue_number_param = (payload.get("queue_number") or "").strip()
        selected_fields_param = normalize_selected_fields(payload.get("selected_fields") or [])
        total_rows_estimate = len(raw_rows) if raw_rows else (record.get("total_in") or 0)
        final_queue_num = queue_number_param if queue_number_param else "TEMP"
        status_init = "processing"

        queue_item = {
            "id": str(uuid.uuid4()),
            "record_id": record_id,
            "queue_number": final_queue_num,
            "requester_name": final_requester_name,
            "observacoes": payload.get("observacoes", ""),
            "selected_fields": selected_fields_param,
            "query_fields": query_fields,
            "request_time": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "status": status_init,
            "total_rows": total_rows_estimate,
            "processed_count": 0,
            "success_count": 0,
            "error_count": 0,
            "logs": initial_logs,
            "rows": [],
            "completed_at": None,
            "mode": mode,
            "export_format": export_format,
        }
        db_save_queue_item(queue_item)

        db_save_audit_event(
            "REQUEST_IMPORT",
            final_requester_name,
            {
                "queue_id": queue_item["id"],
                "mode": mode,
                "total_rows": total_rows_estimate,
                "duplicates_removed": dup_count,
                "query_fields": query_fields,
            }
        )

        # Inicia a esteira assincrona particionada em sub-lotes imediatamente
        threading.Thread(
            target=_async_queue_processor,
            args=(queue_item["id"], raw_rows, record_id, selected_fields_param, queue_number_param, final_requester_name, deduplicate_flag),
            daemon=True,
            name=f"QueueProcessor-{queue_item['id'][:8]}"
        ).start()

        return jsonify({
            "message": "Solicitação inserida na esteira de processamento da fila.",
            "request_id": queue_item["id"],
            "status": "processing",
            "mode": mode,
            "duplicates_removed": dup_count,
        })

    queue_item = {
        "id": str(uuid.uuid4()),
        "record_id": record_id,
        "queue_number": "",
        "requester_name": final_requester_name,
        "observacoes": "",
        "selected_fields": [],
        "query_fields": query_fields,
        "request_time": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "status": "queued",
        "total_rows": len(raw_rows),
        "processed_count": 0,
        "success_count": 0,
        "error_count": 0,
        "logs": initial_logs,
        "rows": [],
        "completed_at": None,
        "mode": mode,
        "export_format": export_format,
    }
    db_save_queue_item(queue_item)

    # Log de auditoria para criação da solicitação
    db_save_audit_event(
        "REQUEST_IMPORT",
        final_requester_name,
        {
            "queue_id": queue_item["id"],
            "mode": mode,
            "total_rows": len(raw_rows),
            "duplicates_removed": dup_count,
            "query_fields": query_fields,
        }
    )

    rec_type = record.get("type") or (detect_enrichment_type(list(raw_rows[0].keys())) if raw_rows else "NOME_UF")
    ENRICHMENT_QUEUE_MANAGER.enqueue(queue_item["id"], raw_rows, rec_type, query_fields)
    pos = ENRICHMENT_QUEUE_MANAGER.get_position(queue_item["id"])

    return jsonify({
        "message": "Solicitação inserida na fila de processamento.",
        "request_id": queue_item["id"],
        "status": "queued",
        "queue_position": pos,
        "mode": mode,
        "duplicates_removed": dup_count,
    })



def pregenerate_export_file(request_id: str):
    """Pré-gera o arquivo Excel em segundo plano assim que a fila é finalizada."""
    def _worker():
        try:
            item = db_get_queue_item(request_id)
            if not item:
                return
            is_enrich_only = (item.get("mode") == "enrich_only") or (item.get("queue_number") == "ENRICH")
            is_queue_only = (item.get("mode") == "queue_only")

            original_columns = []
            record = db_get_record(item.get("record_id", ""))
            if record:
                original_columns = record.get("columns") or []

            raw_source_rows = []
            csv_path = EXPORT_DIR / f"raw_{request_id}.csv"
            if csv_path.exists():
                with open(csv_path, "r", encoding="utf-8") as f:
                    reader = csv.DictReader(f, delimiter=";")
                    for row in reader:
                        raw_source_rows.append(dict(row))
            if not raw_source_rows:
                raw_source_rows = item.get("rows") or []

            fail_cols = [c for c in original_columns if not c.startswith("enriquecido_") and c not in ("_status_enriquecimento", "tipo_enriquecimento")]
            if not fail_cols and raw_source_rows:
                fail_cols = [k for k in raw_source_rows[0].keys() if not k.startswith("enriquecido_") and not k.startswith("_") and k != "tipo_enriquecimento"]
            if not fail_cols:
                fail_cols = list(original_columns) if original_columns else ["Falha"]

            if is_enrich_only:
                # 1. Modelo Layout Original (planilha única idêntica à anexada, mesma aba e ordem original com enriquecimento)
                filepath_orig = EXPORT_DIR / f"enriquecimento_{request_id}_original.xlsx"
                with _get_export_lock(request_id):
                    if not filepath_orig.exists() or filepath_orig.stat().st_size == 0:
                        orig_sheet_name = get_original_sheet_title(record)
                        selected = item.get("selected_fields") or []
                        enrich_cols = [
                            c for c in selected 
                            if c not in original_columns 
                            and not c.startswith("_") 
                            and c != "tipo_enriquecimento"
                        ]
                        if not enrich_cols and raw_source_rows:
                            sample = raw_source_rows[0]
                            enrich_cols = [
                                k for k in sample.keys() 
                                if k.startswith("enriquecido_") and k not in original_columns and not k.startswith("_")
                            ]
                        orig_layout_cols = list(original_columns) + enrich_cols
                        sheets_orig = [{
                            "title": orig_sheet_name,
                            "columns": orig_layout_cols,
                            "rows": raw_source_rows,
                        }]
                        write_export_workbook(filepath=filepath_orig, sheets=sheets_orig, include_summary=False)

                # 2. Modelo 3 Abas (Enriquecido, Falha e Resumo)
                filepath_3abas = EXPORT_DIR / f"enriquecimento_{request_id}_3abas.xlsx"
                with _get_export_lock(request_id):
                    if not filepath_3abas.exists() or filepath_3abas.stat().st_size == 0:
                        export_columns = item.get("selected_fields") or list(original_columns)
                        success_rows = []
                        failed_rows = []
                        for row in raw_source_rows:
                            if is_row_success(row):
                                success_rows.append(dict(row))
                            else:
                                failed_rows.append(dict(row))
                        sheets_3abas = [{"title": "Enriquecido", "columns": export_columns, "rows": success_rows}]
                        if failed_rows:
                            sheets_3abas.append({"title": "Falha", "columns": fail_cols if fail_cols else export_columns, "rows": failed_rows})
                        write_export_workbook(filepath=filepath_3abas, sheets=sheets_3abas, include_summary=True)
            elif is_queue_only:
                filepath = EXPORT_DIR / f"fila_{item.get('queue_number') or 'importar'}_{request_id}.xlsx"
                with _get_export_lock(request_id):
                    if not filepath.exists() or filepath.stat().st_size == 0:
                        concatenator = FastConcatenator(item.get("selected_fields") or [])
                        success_import_rows = []
                        failed_original_rows = []
                        for row in raw_source_rows:
                            if is_valid_queue_row(row):
                                success_import_rows.append({"importar": concatenator.format(row)})
                            else:
                                failed_original_rows.append(dict(row))
                        sheets = [
                            {"title": "Importar", "columns": ["importar"], "rows": success_import_rows},
                            {"title": "Falha", "columns": fail_cols, "rows": failed_original_rows}
                        ]
                        write_export_workbook(filepath=filepath, sheets=sheets)
            else:
                filepath = EXPORT_DIR / f"fila_{item.get('queue_number') or 'importar'}_{request_id}.xlsx"
                with _get_export_lock(request_id):
                    if not filepath.exists() or filepath.stat().st_size == 0:
                        selected_enrich_fields = item.get("selected_fields") or []
                        if not selected_enrich_fields:
                            selected_enrich_fields = [c for c in (original_columns + PREFIXED_ENRICH_FIELDS) if not c.startswith("_")]
                        concatenator = FastConcatenator(selected_enrich_fields)
                        success_import_rows = []
                        failed_original_rows = []
                        enriched_success_rows = []
                        for row in raw_source_rows:
                            row_dict = dict(row)
                            r_succ = is_row_success(row_dict)
                            r_val = is_valid_queue_row(row_dict)
                            if r_succ:
                                enriched_success_rows.append(row_dict)
                            if r_succ and r_val:
                                success_import_rows.append({"importar": concatenator.format(row_dict)})
                            else:
                                failed_original_rows.append(row_dict)
                        sheets = [
                            {"title": "Importar", "columns": ["importar"], "rows": success_import_rows},
                            {"title": "Falha", "columns": fail_cols, "rows": failed_original_rows},
                            {"title": "Enriquecido", "columns": selected_enrich_fields, "rows": enriched_success_rows},
                        ]
                        write_export_workbook(filepath=filepath, sheets=sheets)
        except Exception as e:
            log_api_debug(f"Erro em pregenerate_export_file({request_id}): {e}")

    threading.Thread(target=_worker, daemon=True).start()


@app.post("/api/request/<request_id>/finalize")
@login_required
def finalize_import_request(request_id: str):
    payload = request.get_json(silent=True) or {}
    queue_number = payload.get("queue_number")
    requester_name_input = payload.get("requester_name", "").strip()
    observations = payload.get("observacoes", "")
    selected_fields = payload.get("selected_fields", [])
    mode = (payload.get("mode") or "enrich_and_queue").strip()
    
    current_user_name = session.get("user_name", "Usuário")
    if requester_name_input:
        requester_name = f"{current_user_name} ({requester_name_input})"
    else:
        requester_name = current_user_name

    item = db_get_queue_item(request_id)
    if item is None:
        return jsonify({"error": "Solicitação não encontrada."}), 404

    selected_fields = normalize_selected_fields(selected_fields)
    if not selected_fields:
        record = db_get_record(item.get("record_id", ""))
        selected_fields = (record.get("columns") if record else []) or []

    export_format = (payload.get("export_format") or item.get("export_format") or "original").strip().lower()
    item["export_format"] = export_format

    if mode == "enrich_only" or item.get("mode") == "enrich_only":
        if not selected_fields:
            selected_fields = ["Nome", "Telefone"]
        item["selected_fields"] = selected_fields
        item["status"] = "completed"
        item["completed_at"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        item["queue_number"] = "ENRICH"
        item["total_rows_enriched"] = item.get("total_rows_enriched", item.get("success_count", 0))
        
        if requester_name_input:
            item["requester_name"] = f"{current_user_name} ({requester_name_input})"
        elif not item.get("requester_name") or item.get("requester_name") == "Aguardando confirmação":
            item["requester_name"] = current_user_name
        requester = item.get("requester_name", "").strip()
        
        db_update_queue_item(request_id, {
            "selected_fields": item["selected_fields"],
            "status": item["status"],
            "completed_at": item["completed_at"],
            "queue_number": item["queue_number"],
            "requester_name": item["requester_name"],
            "export_format": item["export_format"],
        })

        db_save_audit_event("QUEUE_FINALIZED", requester, {
            "request_id": request_id,
            "mode": "enrich_only",
            "selected_fields": selected_fields,
        })

        # Dispara pré-geração em segundo plano
        pregenerate_export_file(request_id)

        if requester:
            notification = {
                "id": str(uuid.uuid4()),
                "title": f"Enriquecimento concluído",
                "message": f"Arquivo enriquecido concluído para {requester} e pronto para download.",
                "time": item["completed_at"],
            }
            db_save_notification(notification)

        return jsonify({
            "message": "Arquivo enriquecido finalizado e pronto para download.",
            "item": item,
        })

    if not queue_number or not requester_name:
        return jsonify({"error": "Número da fila e nome do solicitante são obrigatórios."}), 400

    if not re.fullmatch(r"\d{6}", str(queue_number)):
        return jsonify({"error": "O número da fila deve conter exatamente 6 dígitos numéricos."}), 400

    if not selected_fields:
        selected_fields = ["Nome", "Telefone"]

    is_queue_only = (mode == "queue_only") or (item.get("mode") == "queue_only")
    if is_queue_only:
        clean_queue_num = str(queue_number)
        old_queue = str(item.get("queue_number") or "importar").strip()
        item["queue_number"] = clean_queue_num
        item["requester_name"] = requester_name
        item["observacoes"] = observations
        item["selected_fields"] = selected_fields
        item["mode"] = "queue_only"
        if not item.get("completed_at"):
            item["completed_at"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        db_update_queue_item(request_id, {
            "queue_number": clean_queue_num,
            "requester_name": requester_name,
            "observacoes": observations,
            "selected_fields": selected_fields,
            "completed_at": item["completed_at"],
            "export_format": item["export_format"],
        })

        # Sincroniza o arquivo Excel se a fila mudou de TEMP/importar para o numero definitivo
        old_file = EXPORT_DIR / f"fila_{old_queue}_{request_id}.xlsx"
        new_file = EXPORT_DIR / f"fila_{clean_queue_num}_{request_id}.xlsx"
        if old_file.exists() and not new_file.exists():
            try:
                shutil.copy2(old_file, new_file)
            except Exception as e:
                log_api_debug(f"Erro ao sincronizar arquivo da fila {request_id}: {e}")

        # Se ainda nao existe nenhum arquivo pre-gerado, dispara em background
        if not new_file.exists() and not old_file.exists():
            pregenerate_export_file(request_id)

        db_save_audit_event("QUEUE_FINALIZED", requester_name, {
            "request_id": request_id,
            "queue_number": clean_queue_num,
            "mode": "queue_only",
            "selected_fields": selected_fields,
        })

        notification = {
            "id": str(uuid.uuid4()),
            "title": f"Fila {clean_queue_num} tratada e gerada",
            "message": f"A fila {clean_queue_num} foi tratada e concatenada por {requester_name} para importação direta.",
            "time": item["completed_at"],
        }
        db_save_notification(notification)

        return jsonify({
            "message": "Planilha tratada e concatenada com sucesso! O download foi disponibilizado para importação.",
            "item": item,
        })

    item["queue_number"] = str(queue_number)
    item["requester_name"] = requester_name
    item["observacoes"] = observations
    item["selected_fields"] = selected_fields
    item["mode"] = mode or item.get("mode") or "enrich_and_queue"
    item["status"] = "completed"
    item["completed_at"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    item["total_rows_enriched"] = item.get("total_rows_enriched", item.get("success_count", 0))

    db_update_queue_item(request_id, {
        "queue_number": item["queue_number"],
        "requester_name": item["requester_name"],
        "observacoes": item["observacoes"],
        "selected_fields": item["selected_fields"],
        "status": item["status"],
        "completed_at": item["completed_at"],
        "total_rows_enriched": item["total_rows_enriched"],
        "export_format": item["export_format"],
    })

    db_save_audit_event("QUEUE_FINALIZED", requester_name, {
        "request_id": request_id,
        "queue_number": item["queue_number"],
        "mode": item["mode"],
        "selected_fields": selected_fields,
    })

    # Dispara pré-geração em segundo plano
    pregenerate_export_file(request_id)

    notification = {
        "id": str(uuid.uuid4()),
        "title": f"Fila {queue_number} tratada e gerada",
        "message": f"A fila {queue_number} foi tratada e concatenada por {requester_name} para importação direta.",
        "time": item["completed_at"],
    }
    db_save_notification(notification)

    return jsonify({
        "message": "Planilha tratada e concatenada com sucesso! O download foi disponibilizado para importação.",
        "item": item,
    })



@app.post("/api/request/<request_id>/cancel")
@login_required
def cancel_import_request(request_id: str):
    item = db_get_queue_item(request_id)
    if item is None:
        return jsonify({"error": "Solicitação não encontrada."}), 404

    if item["status"] not in ("processing", "pending", "queued"):
        return jsonify({"error": "Esta solicitação não pode ser cancelada no estado atual."}), 400

    ENRICHMENT_QUEUE_MANAGER.cancel_job(request_id)
    item["status"] = "cancelled"
    logs = item.setdefault("logs", [])
    logs.append("Solicitação cancelada pelo usuário.")
    db_update_queue_item(request_id, {"status": "cancelled", "logs": logs})
    
    db_save_audit_event("QUEUE_CANCELLED", session.get("user_name"), {"request_id": request_id})
    return jsonify({"message": "Solicitação cancelada com sucesso."})


@app.post("/api/request/<request_id>/refuse")
@admin_required
def refuse_import_request(request_id: str):
    payload = request.get_json(silent=True) or {}
    reason = payload.get("reason", "").strip()
    if not reason:
        return jsonify({"error": "Motivo da recusa é obrigatório."}), 400

    item = db_get_queue_item(request_id)
    if item is None:
        return jsonify({"error": "Solicitação não encontrada."}), 404

    item["status"] = "rejected"
    item["reject_reason"] = reason
    item["logs"].append(f"Solicitação recusada pela Estratégia. Motivo: {reason}")
    db_update_queue_item(request_id, {"status": "rejected", "reject_reason": reason, "logs": item["logs"]})
    
    db_save_audit_event("QUEUE_REJECTED", session.get("user_name"), {
        "request_id": request_id,
        "queue_number": item.get("queue_number"),
        "reason": reason,
    })
    
    return jsonify({"message": "Solicitação recusada com sucesso."})


@app.get("/api/audit-logs")
@login_required
def get_audit_logs():
    """Retorna os logs de auditoria do sistema."""
    events = db_get_audit_events(limit=150)
    return jsonify({"audit_logs": events})


@app.get("/api/analytics")
@login_required
def get_analytics():
    """Retorna métricas consolidadas de produtividade e assertividade."""
    queue = db_get_all_queue()
    total_queues = len(queue)
    completed_queues = sum(1 for q in queue if q.get("status") == "completed")
    rejected_queues = sum(1 for q in queue if q.get("status") == "rejected")
    processing_queues = sum(1 for q in queue if q.get("status") in ("processing", "pending"))
    
    total_records_processed = sum(int(q.get("total_rows") or 0) for q in queue)
    total_records_enriched = sum(int(q.get("total_rows_enriched") or q.get("success_count") or 0) for q in queue)
    
    avg_assertiveness = 0.0
    if total_records_processed > 0:
        avg_assertiveness = round((total_records_enriched / total_records_processed) * 100, 1)

    return jsonify({
        "total_queues": total_queues,
        "completed_queues": completed_queues,
        "rejected_queues": rejected_queues,
        "processing_queues": processing_queues,
        "total_records_processed": total_records_processed,
        "total_records_enriched": total_records_enriched,
        "avg_assertiveness": avg_assertiveness,
    })



@app.get("/api/me")
@login_required
def get_current_user():
    return jsonify({
        "user_id": session.get("user_id"),
        "role": session.get("role"),
        "name": session.get("user_name"),
    })


@app.get("/api/queue")
@login_required
def get_queue():
    queue = db_get_all_queue()
    def _get_queue_sort_key(item: Any) -> int:
        if not isinstance(item, dict):
            return 0
        qn = item.get("queue_number")
        if isinstance(qn, int):
            return qn
        if isinstance(qn, str) and qn.isdigit():
            return int(qn)
        return 0

    ordered = sorted(queue, key=_get_queue_sort_key)
    # Sanitiza e otimiza payload para a listagem da fila: remove logs extensos e linhas brutas
    # Detalhes e logs completos continuam disponíveis sob demanda via /api/request/<id>/logs
    sanitized = []
    for it in ordered:
        if isinstance(it, dict):
            item_copy = {k: v for k, v in it.items() if k not in ("logs", "rows", "preview_row")}
            sanitized.append(item_copy)
        else:
            sanitized.append(it)
    return jsonify({"queue": sanitized})


@app.get("/api/notifications")
@login_required
def get_notifications():
    notifications = db_get_notifications()
    return jsonify({"notifications": notifications})


def _get_or_create_export_file(request_id: str, export_format: str = "", custom_name: str = "") -> tuple[Path, str, str]:
    item = db_get_queue_item(request_id)
    if not item:
        raise ValueError("Solicitação não encontrada.")

    is_enrich_only = (item.get("mode") == "enrich_only") or (item.get("queue_number") == "ENRICH")
    is_queue_only = (item.get("mode") == "queue_only")

    original_columns = []
    record = db_get_record(item.get("record_id", ""))
    if record:
        original_columns = record.get("columns") or []

    raw_source_rows = []
    csv_path = EXPORT_DIR / f"raw_{request_id}.csv"
    if csv_path.exists():
        try:
            with open(csv_path, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f, delimiter=";")
                for row in reader:
                    raw_source_rows.append(dict(row))
        except Exception as e:
            log_api_debug(f"Erro ao ler CSV local para download {request_id}: {e}")

    if not raw_source_rows:
        raw_source_rows = item.get("rows") or []

    completed_time = item.get("completed_at") or item.get("request_time")
    item_mode = "enrich_only" if is_enrich_only else ("queue_only" if is_queue_only else "enrich_and_queue")
    download_filename = build_download_filename(
        requester_name=item.get("requester_name") or "",
        completed_at=completed_time,
        mode=item_mode,
        queue_number=item.get("queue_number") or "",
        suffix="completa" if export_format == "full" else "",
    )

    if custom_name:
        ext = ".csv" if (export_format == "csv" or custom_name.lower().endswith(".csv")) else ".xlsx"
        download_filename = custom_name if custom_name.lower().endswith(ext) else f"{custom_name}{ext}"

    if export_format == "csv" or download_filename.lower().endswith(".csv"):
        if csv_path.exists():
            return csv_path, download_filename.replace(".xlsx", ".csv"), "text/csv"

    fail_cols = [c for c in original_columns if not c.startswith("enriquecido_") and c not in ("_status_enriquecimento", "tipo_enriquecimento")]
    if not fail_cols and raw_source_rows:
        fail_cols = [k for k in raw_source_rows[0].keys() if not k.startswith("enriquecido_") and not k.startswith("_") and k != "tipo_enriquecimento"]
    if not fail_cols:
        fail_cols = list(original_columns) if original_columns else ["Falha"]

    # 1. Modo Apenas Enriquecer (enrich_only)
    if is_enrich_only:
        # Se for solicitado o formato de 3 abas (Enriquecido, Falha e Resumo)
        chosen_format = (export_format or item.get("export_format") or "original").strip().lower()
        if chosen_format in ("standard", "3abas", "resumo"):
            filepath = EXPORT_DIR / f"enriquecimento_{request_id}_3abas.xlsx"
            if not filepath.exists() or filepath.stat().st_size == 0:
                export_columns = item.get("selected_fields") or list(original_columns)
                success_rows = []
                failed_rows = []
                for row in raw_source_rows:
                    if is_row_success(row):
                        success_rows.append(dict(row))
                    else:
                        failed_rows.append(dict(row))

                sheets = [
                    {
                        "title": "Enriquecido",
                        "columns": export_columns,
                        "rows": success_rows,
                    }
                ]
                if failed_rows:
                    sheets.append({
                        "title": "Falha",
                        "columns": fail_cols if fail_cols else export_columns,
                        "rows": failed_rows,
                    })

                write_export_workbook(filepath=filepath, sheets=sheets, include_summary=True)

            dl_name_3abas = download_filename
            if not custom_name and not dl_name_3abas.lower().endswith("_3abas.xlsx"):
                dl_name_3abas = dl_name_3abas.replace(".xlsx", "_3abas.xlsx")
            return filepath, dl_name_3abas, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        # Formato Padrão / Layout Original: Devolve a planilha IGUAL ela foi anexada (mesma aba, mesma ordem, só com colunas enriquecidas adicionadas)
        filepath = EXPORT_DIR / f"enriquecimento_{request_id}_original.xlsx"
        if not filepath.exists() or filepath.stat().st_size == 0:
            orig_sheet_name = get_original_sheet_title(record)
            selected = item.get("selected_fields") or []
            enrich_cols = [
                c for c in selected 
                if c not in original_columns 
                and not c.startswith("_") 
                and c != "tipo_enriquecimento"
            ]
            if not enrich_cols and raw_source_rows:
                sample = raw_source_rows[0]
                enrich_cols = [
                    k for k in sample.keys() 
                    if k.startswith("enriquecido_") and k not in original_columns and not k.startswith("_")
                ]

            orig_layout_cols = list(original_columns) + enrich_cols

            sheets = [
                {
                    "title": orig_sheet_name,
                    "columns": orig_layout_cols,
                    "rows": raw_source_rows,
                }
            ]
            write_export_workbook(filepath=filepath, sheets=sheets, include_summary=False)

        return filepath, download_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # 2. Fila de Importação: Formato COMPLETO (todas as colunas + Resultado Concatenado)
    if export_format == "full":
        filepath = EXPORT_DIR / f"fila_{item.get('queue_number') or 'completa'}_{request_id}_completa.xlsx"
        if not filepath.exists() or filepath.stat().st_size == 0:
            export_columns = list(original_columns)
            for field in item.get("selected_fields") or []:
                if field not in export_columns:
                    export_columns.append(field)
            if "Resultado Concatenado" not in export_columns:
                export_columns.append("Resultado Concatenado")

            concatenator = FastConcatenator(item.get("selected_fields") or [])
            success_rows = []
            failed_rows = []
            for row in raw_source_rows:
                row_copy = dict(row)
                row_copy["Resultado Concatenado"] = concatenator.format(row_copy)
                if is_queue_only:
                    if is_valid_queue_row(row):
                        success_rows.append(row_copy)
                    else:
                        failed_rows.append(dict(row))
                else:
                    if is_row_success(row) and is_valid_queue_row(row):
                        success_rows.append(row_copy)
                    else:
                        failed_rows.append(dict(row))

            sheets = [
                {
                    "title": "Importar",
                    "columns": export_columns,
                    "rows": success_rows,
                },
                {
                    "title": "Falha",
                    "columns": fail_cols,
                    "rows": failed_rows,
                }
            ]
            write_export_workbook(filepath=filepath, sheets=sheets)
        return filepath, download_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # 3. Modo Subir Fila (queue_only)
    if is_queue_only:
        q_num = item.get("queue_number") or "importar"
        filepath = EXPORT_DIR / f"fila_{q_num}_{request_id}.xlsx"
        if not filepath.exists() or filepath.stat().st_size == 0:
            fallback_path = EXPORT_DIR / f"fila_importar_{request_id}.xlsx"
            if fallback_path.exists() and fallback_path.stat().st_size > 0:
                try:
                    shutil.copy2(fallback_path, filepath)
                except Exception:
                    filepath = fallback_path

        if not filepath.exists() or filepath.stat().st_size == 0:
            with _get_export_lock(request_id):
                if not filepath.exists() or filepath.stat().st_size == 0:
                    concatenator = FastConcatenator(item.get("selected_fields") or [])
                    success_import_rows = []
                    failed_original_rows = []
                    for row in raw_source_rows:
                        if is_valid_queue_row(row):
                            concatenated_val = concatenator.format(row)
                            success_import_rows.append({"importar": concatenated_val})
                        else:
                            failed_original_rows.append(dict(row))

                    sheets = [
                        {
                            "title": "Importar",
                            "columns": ["importar"],
                            "rows": success_import_rows,
                        },
                        {
                            "title": "Falha",
                            "columns": fail_cols,
                            "rows": failed_original_rows,
                        }
                    ]
                    write_export_workbook(filepath=filepath, sheets=sheets)
        return filepath, download_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # 4. Modo Enriquecer e Subir (enrich_and_queue)
    filepath = EXPORT_DIR / f"fila_{item.get('queue_number') or 'importar'}_{request_id}.xlsx"
    if not filepath.exists() or filepath.stat().st_size == 0:
        selected_enrich_fields = item.get("selected_fields") or []
        if not selected_enrich_fields:
            selected_enrich_fields = [c for c in (original_columns + PREFIXED_ENRICH_FIELDS) if not c.startswith("_")]

        concatenator = FastConcatenator(selected_enrich_fields)
        success_import_rows = []
        failed_original_rows = []
        enriched_success_rows = []

        for row in raw_source_rows:
            row_dict = dict(row)
            row_is_success = is_row_success(row_dict)
            row_is_valid_queue = is_valid_queue_row(row_dict)

            if row_is_success:
                enriched_success_rows.append(row_dict)

            if row_is_success and row_is_valid_queue:
                concatenated_val = concatenator.format(row_dict)
                success_import_rows.append({"importar": concatenated_val})
            else:
                failed_original_rows.append(row_dict)

        sheets = [
            {
                "title": "Importar",
                "columns": ["importar"],
                "rows": success_import_rows,
            },
            {
                "title": "Falha",
                "columns": fail_cols,
                "rows": failed_original_rows,
            },
            {
                "title": "Enriquecido",
                "columns": selected_enrich_fields,
                "rows": enriched_success_rows,
            }
        ]

        write_export_workbook(filepath=filepath, sheets=sheets)
    return filepath, download_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@app.get("/api/request/<request_id>/download/info")
@login_required
def get_download_info(request_id: str):
    try:
        export_format = request.args.get("format", "").strip().lower()
        custom_name = request.args.get("custom_name", "").strip()
        filepath, download_filename, mime_type = _get_or_create_export_file(request_id, export_format, custom_name)

        filesize = filepath.stat().st_size
        chunk_size = 4 * 1024 * 1024  # 4MB por fatia
        total_chunks = max(1, math.ceil(filesize / chunk_size))

        return jsonify({
            "filename": download_filename,
            "filesize": filesize,
            "chunk_size": chunk_size,
            "total_chunks": total_chunks,
            "mime_type": mime_type,
        })
    except Exception as e:
        log_api_debug(f"Erro em get_download_info({request_id}): {e}")
        return jsonify({"error": str(e)}), 400


@app.get("/api/request/<request_id>/download/chunk")
@login_required
def download_export_chunk(request_id: str):
    try:
        export_format = request.args.get("format", "").strip().lower()
        custom_name = request.args.get("custom_name", "").strip()
        chunk_index = int(request.args.get("chunk_index", 0))
        chunk_size = int(request.args.get("chunk_size", 4 * 1024 * 1024))

        filepath, download_filename, mime_type = _get_or_create_export_file(request_id, export_format, custom_name)
        filesize = filepath.stat().st_size

        start = chunk_index * chunk_size
        if start >= filesize:
            return jsonify({"error": "Índice de fatia fora dos limites."}), 416

        with open(filepath, "rb") as f:
            f.seek(start)
            chunk_bytes = f.read(chunk_size)

        response = Response(chunk_bytes, mimetype="application/octet-stream")
        response.headers["Content-Length"] = str(len(chunk_bytes))
        response.headers["X-Chunk-Index"] = str(chunk_index)
        response.headers["X-Total-Size"] = str(filesize)
        return response
    except Exception as e:
        log_api_debug(f"Erro em download_export_chunk({request_id}): {e}")
        return jsonify({"error": str(e)}), 400


@app.get("/api/request/<request_id>/download")
@login_required
def download_export(request_id: str):
    export_format = request.args.get("format", "").strip().lower()
    custom_name = request.args.get("custom_name", "").strip()
    part = request.args.get("part", "").strip()
    try:
        if part and part.isdigit():
            part_num = int(part)
            item = db_get_queue_item(request_id)
            queue_num = (item.get("queue_number") or "importar") if item else "importar"
            treated_part = EXPORT_DIR / f"fila_{queue_num}_{request_id}_part{part_num}.csv"
            if not treated_part.exists():
                treated_part = EXPORT_DIR / f"fila_importar_{request_id}_part{part_num}.csv"

            if treated_part.exists():
                dl_name = f"fila_{queue_num}_parte_{part_num}_{request_id}.csv"
                return send_file(treated_part, as_attachment=True, download_name=dl_name, mimetype="text/csv")

            part_file = EXPORT_DIR / f"raw_{request_id}_part{part_num}.csv"
            if part_file.exists():
                return send_file(part_file, as_attachment=True, download_name=f"parte_{part_num}_{request_id}.csv", mimetype="text/csv")
            return jsonify({"error": f"Fatia {part_num} não encontrada."}), 404

        filepath, download_filename, mime_type = _get_or_create_export_file(request_id, export_format, custom_name)

        db_save_audit_event("FILE_DOWNLOAD", session.get("user_name"), {
            "request_id": request_id,
            "filename": download_filename,
            "format": export_format or "standard",
        })

        return send_file(filepath, as_attachment=True, download_name=download_filename, mimetype=mime_type)
    except Exception as e:
        log_api_debug(f"Erro em download_export({request_id}): {e}")
        return jsonify({"error": str(e)}), 400


@app.post("/api/batch/download-zip")
@login_required
def batch_download_zip():
    payload = request.get_json(silent=True) or {}
    items = payload.get("items", [])
    zip_name = payload.get("zip_name", "lote_planilhas.zip").strip()
    if not zip_name.lower().endswith(".zip"):
        zip_name += ".zip"

    if not items:
        return jsonify({"error": "Nenhum item informado para download."}), 400

    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, itm in enumerate(items):
            req_id = itm.get("request_id")
            filename = itm.get("filename")
            if not req_id:
                continue
            try:
                filepath, default_name, _ = _get_or_create_export_file(req_id, export_format="", custom_name=filename)
                if filepath and filepath.exists():
                    arcname = filename if filename else default_name
                    zf.write(filepath, arcname=arcname)
            except Exception as e:
                log_api_debug(f"Erro ao incluir {req_id} no zip: {e}")

    bio.seek(0)
    db_save_audit_event("BATCH_ZIP_DOWNLOAD", session.get("user_name"), {
        "zip_name": zip_name,
        "total_files": len(items),
    })

    return send_file(bio, as_attachment=True, download_name=zip_name, mimetype="application/zip")




@app.get("/api/request/<request_id>/logs")
@login_required
def get_request_logs(request_id: str):
    item = db_get_queue_item(request_id)
    if item is None:
        return jsonify({"error": "Solicitação não encontrada."}), 404
        
    preview_row = item.get("preview_row")
    if not preview_row or (not preview_row.get("enriquecido_cpf") and not preview_row.get("enriquecido_telefone")):
        if item.get("rows"):
            for r in item["rows"]:
                if r.get("enriquecido_cpf") or r.get("enriquecido_telefone") or str(r.get("_status_enriquecimento") or "").strip().lower() == "sucesso":
                    preview_row = r
                    item["preview_row"] = preview_row
                    break
        if not preview_row or (not preview_row.get("enriquecido_cpf") and not preview_row.get("enriquecido_telefone")):
            csv_path = EXPORT_DIR / f"raw_{request_id}.csv"
            if csv_path.exists() and item.get("status") != "queued":
                try:
                    with open(csv_path, "r", encoding="utf-8") as f:
                        reader = csv.DictReader(f, delimiter=";")
                        first_row = None
                        for row in reader:
                            if first_row is None:
                                first_row = dict(row)
                            if row.get("enriquecido_cpf") or row.get("enriquecido_telefone") or str(row.get("_status_enriquecimento") or "").strip().lower() == "sucesso":
                                preview_row = dict(row)
                                item["preview_row"] = preview_row
                                break
                        if not preview_row and first_row:
                            preview_row = first_row
                            item["preview_row"] = preview_row
                except Exception:
                    pass
        
    queue_pos = ENRICHMENT_QUEUE_MANAGER.get_position(request_id) if item.get("status") == "queued" else 0
    total_queued = ENRICHMENT_QUEUE_MANAGER.get_total_queued()
        
    return jsonify({
        "logs": item.get("logs", []),
        "processed_count": item.get("processed_count", 0),
        "total_rows": item.get("total_rows", 0),
        "success_count": item.get("success_count", 0),
        "error_count": item.get("error_count", 0),
        "status": item["status"],
        "queue_position": queue_pos,
        "total_queued": total_queued,
        "preview": preview_row or {},
        "chunk_info": item.get("chunk_info", {}),
        "current_chunk": item.get("current_chunk", 1),
        "total_chunks": item.get("total_chunks", 1),
    })


@app.post("/api/request/<request_id>/complete")
@admin_required
def complete_request(request_id: str):
    item = db_get_queue_item(request_id)
    if item is None:
        return jsonify({"error": "Solicitação não encontrada."}), 404

    uploaded_file = request.files.get("summaryImage")
    summary_name = "resumo_importacao.png"
    if uploaded_file and uploaded_file.filename:
        summary_name = uploaded_file.filename
        destination = EXPORT_DIR / f"summary_{request_id}_{summary_name}"
        uploaded_file.save(destination)

    item["status"] = "completed"
    item["summary_name"] = summary_name
    item["completed_at"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    db_update_queue_item(request_id, {
        "status": item["status"],
        "summary_name": item["summary_name"],
        "completed_at": item["completed_at"]
    })

    notification = {
        "id": str(uuid.uuid4()),
        "title": f"Importação da fila {item['queue_number']} concluída",
        "message": f"A importação solicitada por {item['requester_name']} foi concluída com sucesso.",
        "time": item["completed_at"],
    }
    db_save_notification(notification)

    return jsonify({
        "message": "Solicitação concluída com sucesso.",
        "notification": notification,
    })


@app.get("/api/summary/<request_id>")
@login_required
def get_summary_image(request_id: str):
    item = db_get_queue_item(request_id)
    if item is None or not item.get("summary_name"):
        return jsonify({"error": "Resumo não encontrado."}), 404
        
    filename = f"summary_{request_id}_{item['summary_name']}"
    filepath = EXPORT_DIR / filename
    if not filepath.exists():
        return jsonify({"error": "Arquivo não encontrado no servidor."}), 404
        
    return send_file(filepath)


@app.get("/api/health")
def health():
    return jsonify({"status": "ok"})


@app.get("/api/download-sample")
@login_required
def download_sample():
    filepath = BASE_DIR / "modelo_teste.csv"
    return send_file(filepath, as_attachment=True, download_name="modelo_teste.csv")


CONVERT_DIR = TEMP_PARENT / "conversions"
CONVERT_DIR.mkdir(exist_ok=True)
CONVERTED_FILES: Dict[str, Dict[str, Any]] = {}
CONVERSION_TASKS: Dict[str, Dict[str, Any]] = {}


def _save_conversion_task_state(task_id: str, state_data: Dict[str, Any], task_dir: Optional[Path] = None):
    try:
        if task_dir is None:
            task_dir = CONVERT_DIR / task_id
        if task_dir.exists():
            state_file = task_dir / "task_state.json"
            tmp_file = task_dir / "task_state.json.tmp"
            tmp_file.write_text(json.dumps(state_data, ensure_ascii=False), encoding="utf-8")
            tmp_file.replace(state_file)
    except Exception:
        pass


def _async_converter_worker(
    task_id: str, 
    tool: str, 
    saved_inputs: List[Path], 
    options: Dict[str, Any], 
    task_dir: Path, 
    user_name: str
):
    try:
        import converter_service

        def progress_cb(processed: int, total: int, step_desc: str):
            if task_id in CONVERSION_TASKS:
                pct = int((processed / total) * 100) if total > 0 else 0
                CONVERSION_TASKS[task_id].update({
                    "processed_pages": processed,
                    "total_pages": total,
                    "progress": min(99, max(0, pct)),
                    "current_step": step_desc,
                    "page_info": f"Página {processed} de {total}" if total > 0 else ""
                })
                _save_conversion_task_state(task_id, CONVERSION_TASKS[task_id], task_dir)

        num_files = len(saved_inputs)
        output_filepath = None
        download_filename = ""
        mimetype = "application/octet-stream"

        if tool == "pdf_to_word":
            if num_files == 1:
                first_file = saved_inputs[0]
                output_filepath = task_dir / f"{first_file.stem}.docx"
                converter_service.convert_pdf_to_word(first_file, output_filepath, options=options, progress_callback=progress_cb)
                download_filename = f"{first_file.stem}.docx"
                mimetype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            else:
                out_files = []
                for idx, f_in in enumerate(saved_inputs):
                    out_f = task_dir / f"{f_in.stem}.docx"
                    if out_f.exists():
                        out_f = task_dir / f"{f_in.stem}_{idx+1}.docx"
                    converter_service.convert_pdf_to_word(f_in, out_f, options=options, progress_callback=progress_cb)
                    out_files.append(out_f)
                zip_path = task_dir / "documentos_word.zip"
                converter_service.create_zip_from_files(out_files, zip_path)
                output_filepath = zip_path
                download_filename = "documentos_word.zip"
                mimetype = "application/zip"

        elif tool == "pdf_to_excel":
            if num_files == 1:
                first_file = saved_inputs[0]
                output_filepath = task_dir / f"{first_file.stem}.xlsx"
                converter_service.convert_pdf_to_excel(first_file, output_filepath, options=options, progress_callback=progress_cb)
                download_filename = f"{first_file.stem}.xlsx"
                mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            else:
                out_files = []
                for idx, f_in in enumerate(saved_inputs):
                    out_f = task_dir / f"{f_in.stem}.xlsx"
                    if out_f.exists():
                        out_f = task_dir / f"{f_in.stem}_{idx+1}.xlsx"
                    converter_service.convert_pdf_to_excel(f_in, out_f, options=options, progress_callback=progress_cb)
                    out_files.append(out_f)
                zip_path = task_dir / "planilhas_convertidas.zip"
                converter_service.create_zip_from_files(out_files, zip_path)
                output_filepath = zip_path
                download_filename = "planilhas_convertidas.zip"
                mimetype = "application/zip"

        elif tool == "pdf_to_images":
            img_format = options.get("format", "png").lower()
            if num_files == 1:
                first_file = saved_inputs[0]
                out_img_dir = task_dir / "images"
                imgs = converter_service.convert_pdf_to_images(first_file, out_img_dir, img_format=img_format, options=options, progress_callback=progress_cb)
                if len(imgs) == 1:
                    output_filepath = Path(imgs[0])
                    download_filename = output_filepath.name
                    mimetype = f"image/{img_format}"
                else:
                    zip_path = task_dir / f"{first_file.stem}_imagens.zip"
                    converter_service.create_zip_from_files(imgs, zip_path)
                    output_filepath = zip_path
                    download_filename = f"{first_file.stem}_imagens.zip"
                    mimetype = "application/zip"
            else:
                all_imgs = []
                for idx, f_in in enumerate(saved_inputs):
                    out_img_dir = task_dir / f"images_{idx+1}"
                    imgs = converter_service.convert_pdf_to_images(f_in, out_img_dir, img_format=img_format, options=options, progress_callback=progress_cb)
                    all_imgs.extend(imgs)
                zip_path = task_dir / "todas_imagens.zip"
                converter_service.create_zip_from_files(all_imgs, zip_path)
                output_filepath = zip_path
                download_filename = "todas_imagens.zip"
                mimetype = "application/zip"

        elif tool == "pdf_to_text":
            if num_files == 1:
                first_file = saved_inputs[0]
                output_filepath = task_dir / f"{first_file.stem}.txt"
                converter_service.convert_pdf_to_text(first_file, output_filepath, options=options, progress_callback=progress_cb)
                download_filename = f"{first_file.stem}.txt"
                mimetype = "text/plain"
            else:
                out_files = []
                for idx, f_in in enumerate(saved_inputs):
                    out_f = task_dir / f"{f_in.stem}.txt"
                    if out_f.exists():
                        out_f = task_dir / f"{f_in.stem}_{idx+1}.txt"
                    converter_service.convert_pdf_to_text(f_in, out_f, options=options, progress_callback=progress_cb)
                    out_files.append(out_f)
                zip_path = task_dir / "textos_convertidos.zip"
                converter_service.create_zip_from_files(out_files, zip_path)
                output_filepath = zip_path
                download_filename = "textos_convertidos.zip"
                mimetype = "application/zip"

        elif tool == "word_to_pdf":
            if num_files == 1:
                first_file = saved_inputs[0]
                output_filepath = task_dir / f"{first_file.stem}.pdf"
                converter_service.convert_word_to_pdf(first_file, output_filepath)
                download_filename = f"{first_file.stem}.pdf"
                mimetype = "application/pdf"
            else:
                out_files = []
                for idx, f_in in enumerate(saved_inputs):
                    out_f = task_dir / f"{f_in.stem}.pdf"
                    if out_f.exists():
                        out_f = task_dir / f"{f_in.stem}_{idx+1}.pdf"
                    converter_service.convert_word_to_pdf(f_in, out_f)
                    out_files.append(out_f)
                zip_path = task_dir / "documentos_pdf.zip"
                converter_service.create_zip_from_files(out_files, zip_path)
                output_filepath = zip_path
                download_filename = "documentos_pdf.zip"
                mimetype = "application/zip"

        elif tool == "excel_to_pdf":
            if num_files == 1:
                first_file = saved_inputs[0]
                output_filepath = task_dir / f"{first_file.stem}.pdf"
                converter_service.convert_excel_to_pdf(first_file, output_filepath)
                download_filename = f"{first_file.stem}.pdf"
                mimetype = "application/pdf"
            else:
                out_files = []
                for idx, f_in in enumerate(saved_inputs):
                    out_f = task_dir / f"{f_in.stem}.pdf"
                    if out_f.exists():
                        out_f = task_dir / f"{f_in.stem}_{idx+1}.pdf"
                    converter_service.convert_excel_to_pdf(f_in, out_f)
                    out_files.append(out_f)
                zip_path = task_dir / "relatorios_pdf.zip"
                converter_service.create_zip_from_files(out_files, zip_path)
                output_filepath = zip_path
                download_filename = "relatorios_pdf.zip"
                mimetype = "application/zip"

        elif tool == "word_to_excel":
            if num_files == 1:
                first_file = saved_inputs[0]
                output_filepath = task_dir / f"{first_file.stem}.xlsx"
                converter_service.convert_word_to_excel(first_file, output_filepath, options=options, progress_callback=progress_cb)
                download_filename = f"{first_file.stem}.xlsx"
                mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            else:
                out_files = []
                for idx, f_in in enumerate(saved_inputs):
                    out_f = task_dir / f"{f_in.stem}.xlsx"
                    if out_f.exists():
                        out_f = task_dir / f"{f_in.stem}_{idx+1}.xlsx"
                    converter_service.convert_word_to_excel(f_in, out_f, options=options, progress_callback=progress_cb)
                    out_files.append(out_f)
                zip_path = task_dir / "planilhas_word_convertidas.zip"
                converter_service.create_zip_from_files(out_files, zip_path)
                output_filepath = zip_path
                download_filename = "planilhas_word_convertidas.zip"
                mimetype = "application/zip"


        elif tool == "images_to_pdf":
            first_stem = saved_inputs[0].stem
            out_name = f"{first_stem}_convertido.pdf" if num_files == 1 else "imagens_convertidas.pdf"
            output_filepath = task_dir / out_name
            converter_service.convert_images_to_pdf(saved_inputs, output_filepath)
            download_filename = out_name
            mimetype = "application/pdf"

        elif tool == "merge_pdf":
            output_filepath = task_dir / "documento_mesclado.pdf"
            converter_service.merge_pdfs(saved_inputs, output_filepath)
            download_filename = "documento_mesclado.pdf"
            mimetype = "application/pdf"

        elif tool == "split_pdf":
            split_mode = options.get("splitMode") or options.get("mode") or "all"
            page_range = options.get("splitRange") or options.get("range") or options.get("pageRange") or ""
            if num_files == 1:
                first_file = saved_inputs[0]
                out_split_dir = task_dir / "split"
                split_files = converter_service.split_pdf(first_file, out_split_dir, pages_mode=split_mode, custom_range=page_range, options=options, progress_callback=progress_cb)
                if len(split_files) == 1:
                    output_filepath = Path(split_files[0])
                    download_filename = output_filepath.name
                    mimetype = "application/pdf"
                else:
                    zip_path = task_dir / f"{first_file.stem}_dividido.zip"
                    converter_service.create_zip_from_files(split_files, zip_path)
                    output_filepath = zip_path
                    download_filename = f"{first_file.stem}_dividido.zip"
                    mimetype = "application/zip"
            else:
                all_split_files = []
                for idx, f_in in enumerate(saved_inputs):
                    out_split_dir = task_dir / f"split_{idx+1}"
                    s_files = converter_service.split_pdf(f_in, out_split_dir, pages_mode=split_mode, custom_range=page_range, options=options, progress_callback=progress_cb)
                    all_split_files.extend(s_files)
                zip_path = task_dir / "paginas_divididas.zip"
                converter_service.create_zip_from_files(all_split_files, zip_path)
                output_filepath = zip_path
                download_filename = "paginas_divididas.zip"
                mimetype = "application/zip"

        elif tool == "compress_pdf":
            if num_files == 1:
                first_file = saved_inputs[0]
                output_filepath = task_dir / f"{first_file.stem}_comprimido.pdf"
                converter_service.compress_pdf(first_file, output_filepath)
                download_filename = f"{first_file.stem}_comprimido.pdf"
                mimetype = "application/pdf"
            else:
                out_files = []
                for idx, f_in in enumerate(saved_inputs):
                    out_f = task_dir / f"{f_in.stem}_comprimido.pdf"
                    if out_f.exists():
                        out_f = task_dir / f"{f_in.stem}_comprimido_{idx+1}.pdf"
                    converter_service.compress_pdf(f_in, out_f)
                    out_files.append(out_f)
                zip_path = task_dir / "pdfs_comprimidos.zip"
                converter_service.create_zip_from_files(out_files, zip_path)
                output_filepath = zip_path
                download_filename = "pdfs_comprimidos.zip"
                mimetype = "application/zip"

        elif tool == "rotate_pdf":
            angle = int(options.get("angle", 90))
            if num_files == 1:
                first_file = saved_inputs[0]
                output_filepath = task_dir / f"{first_file.stem}_girado.pdf"
                converter_service.rotate_pdf(first_file, output_filepath, angle=angle, options=options)
                download_filename = f"{first_file.stem}_girado.pdf"
                mimetype = "application/pdf"
            else:
                out_files = []
                for idx, f_in in enumerate(saved_inputs):
                    out_f = task_dir / f"{f_in.stem}_girado.pdf"
                    if out_f.exists():
                        out_f = task_dir / f"{f_in.stem}_girado_{idx+1}.pdf"
                    converter_service.rotate_pdf(f_in, out_f, angle=angle, options=options)
                    out_files.append(out_f)
                zip_path = task_dir / "pdfs_girados.zip"
                converter_service.create_zip_from_files(out_files, zip_path)
                output_filepath = zip_path
                download_filename = "pdfs_girados.zip"
                mimetype = "application/zip"

        elif tool == "protect_pdf":
            password = options.get("password", "")
            if not password:
                raise ValueError("Por favor, defina uma senha para proteger o arquivo.")
            if num_files == 1:
                first_file = saved_inputs[0]
                output_filepath = task_dir / f"{first_file.stem}_protegido.pdf"
                converter_service.protect_pdf(first_file, output_filepath, password=password)
                download_filename = f"{first_file.stem}_protegido.pdf"
                mimetype = "application/pdf"
            else:
                out_files = []
                for idx, f_in in enumerate(saved_inputs):
                    out_f = task_dir / f"{f_in.stem}_protegido.pdf"
                    if out_f.exists():
                        out_f = task_dir / f"{f_in.stem}_protegido_{idx+1}.pdf"
                    converter_service.protect_pdf(f_in, out_f, password=password)
                    out_files.append(out_f)
                zip_path = task_dir / "pdfs_protegidos.zip"
                converter_service.create_zip_from_files(out_files, zip_path)
                output_filepath = zip_path
                download_filename = "pdfs_protegidos.zip"
                mimetype = "application/zip"

        elif tool == "unlock_pdf":
            password = options.get("password", "")
            if num_files == 1:
                first_file = saved_inputs[0]
                output_filepath = task_dir / f"{first_file.stem}_desbloqueado.pdf"
                converter_service.unlock_pdf(first_file, output_filepath, password=password)
                download_filename = f"{first_file.stem}_desbloqueado.pdf"
                mimetype = "application/pdf"
            else:
                out_files = []
                for idx, f_in in enumerate(saved_inputs):
                    out_f = task_dir / f"{f_in.stem}_desbloqueado.pdf"
                    if out_f.exists():
                        out_f = task_dir / f"{f_in.stem}_desbloqueado_{idx+1}.pdf"
                    converter_service.unlock_pdf(f_in, out_f, password=password)
                    out_files.append(out_f)
                zip_path = task_dir / "pdfs_desbloqueados.zip"
                converter_service.create_zip_from_files(out_files, zip_path)
                output_filepath = zip_path
                download_filename = "pdfs_desbloqueados.zip"
                mimetype = "application/zip"

        else:
            raise ValueError(f"Ferramenta de conversão '{tool}' desconhecida.")

        if not output_filepath or not output_filepath.exists():
            raise RuntimeError("Falha na geração do arquivo convertido.")

        file_size = output_filepath.stat().st_size

        def format_size(size_bytes):
            if size_bytes < 1024:
                return f"{size_bytes} B"
            elif size_bytes < 1024 * 1024:
                return f"{size_bytes / 1024:.1f} KB"
            else:
                return f"{size_bytes / (1024 * 1024):.2f} MB"

        CONVERTED_FILES[task_id] = {
            "path": str(output_filepath),
            "download_name": download_filename,
            "mimetype": mimetype,
            "created_at": time.time()
        }

        try:
            db_save_audit_event("CONVERTER_EXECUTE", user_name, {
                "tool": tool,
                "filename": download_filename,
                "filesize": format_size(file_size),
            })
        except Exception:
            pass

        if task_id not in CONVERSION_TASKS:
            CONVERSION_TASKS[task_id] = {}
        CONVERSION_TASKS[task_id].update({
            "status": "completed",
            "progress": 100,
            "current_step": "Conversão concluída com sucesso!",
            "completed_at": time.time(),
            "filename": download_filename,
            "filesize": format_size(file_size),
            "download_id": task_id
        })
        _save_conversion_task_state(task_id, CONVERSION_TASKS[task_id], task_dir)

    except Exception as exc:
        import traceback
        traceback.print_exc()
        if task_id not in CONVERSION_TASKS:
            CONVERSION_TASKS[task_id] = {}
        CONVERSION_TASKS[task_id].update({
            "status": "error",
            "error": str(exc),
            "current_step": f"Erro: {str(exc)}"
        })
        _save_conversion_task_state(task_id, CONVERSION_TASKS[task_id], task_dir)


@app.post("/api/converter/chunk")
@login_required
def converter_upload_chunk():
    try:
        upload_id = request.form.get("upload_id")
        file_index_str = request.form.get("file_index", "0")
        chunk_index_str = request.form.get("chunk_index")
        total_chunks_str = request.form.get("total_chunks")
        filename = request.form.get("filename") or "arquivo.pdf"
        chunk_file = request.files.get("chunk")

        if not upload_id or chunk_index_str is None or total_chunks_str is None or not chunk_file:
            return jsonify({"error": "Parâmetros do fragmento incompletos."}), 400

        file_index = int(file_index_str)
        chunk_index = int(chunk_index_str)
        total_chunks = int(total_chunks_str)

        upload_tmp_dir = CHUNK_UPLOAD_DIR / upload_id
        upload_tmp_dir.mkdir(parents=True, exist_ok=True)

        chunk_path = upload_tmp_dir / f"f{file_index:03d}_{chunk_index:06d}.part"
        chunk_file.save(str(chunk_path))

        meta_file = upload_tmp_dir / f"f{file_index:03d}.meta"
        if not meta_file.exists():
            meta_file.write_text(json.dumps({"filename": filename, "total_chunks": total_chunks}), encoding="utf-8")

        return jsonify({
            "success": True,
            "file_index": file_index,
            "chunk_received": chunk_index,
            "total_chunks": total_chunks
        }), 200

    except Exception as exc:
        log_api_debug(f"Erro em converter_upload_chunk: {exc}")
        return jsonify({"error": f"Erro no processamento do fragmento: {str(exc)}"}), 400


@app.post("/api/converter/process")
@login_required
def converter_process():
    try:
        tool = request.form.get("tool", "").strip()
        options_json = request.form.get("options", "{}")
        upload_id = request.form.get("upload_id", "").strip()
        try:
            options = json.loads(options_json)
        except Exception:
            options = {}

        task_id = str(uuid.uuid4())
        task_dir = CONVERT_DIR / task_id
        task_dir.mkdir(parents=True, exist_ok=True)

        saved_inputs = []

        if upload_id:
            # Reconstroi os arquivos recebidos em fragmentos (chunks)
            upload_tmp_dir = CHUNK_UPLOAD_DIR / upload_id
            if not upload_tmp_dir.exists():
                return jsonify({"error": "Sessão de upload fragmentado não encontrada ou expirada."}), 400

            meta_files = sorted(list(upload_tmp_dir.glob("f*.meta")))
            if not meta_files:
                return jsonify({"error": "Nenhum arquivo encontrado no upload fragmentado."}), 400

            for m_file in meta_files:
                meta = json.loads(m_file.read_text(encoding="utf-8"))
                f_name = meta.get("filename", "arquivo.pdf")
                tot_c = meta.get("total_chunks", 1)
                f_prefix = m_file.stem

                safe_name = f_name.replace(" ", "_")
                dest = task_dir / safe_name
                if dest.exists():
                    s_stem = Path(safe_name).stem
                    s_suf = Path(safe_name).suffix
                    dest = task_dir / f"{s_stem}_{len(saved_inputs)+1}{s_suf}"

                with open(dest, "wb") as outfile:
                    for c_idx in range(tot_c):
                        part_file = upload_tmp_dir / f"{f_prefix}_{c_idx:06d}.part"
                        if not part_file.exists():
                            return jsonify({"error": f"Fatia {c_idx+1} de {tot_c} do arquivo {f_name} ausente no servidor."}), 400
                        outfile.write(part_file.read_bytes())

                saved_inputs.append(dest)

            try:
                import shutil
                shutil.rmtree(upload_tmp_dir, ignore_errors=True)
            except Exception:
                pass

        else:
            # Fluxo tradicional para arquivos diretos
            files = request.files.getlist("files")
            if not files or not files[0].filename:
                return jsonify({"error": "Nenhum arquivo enviado para conversão."}), 400

            for f in files:
                if not f.filename:
                    continue
                safe_name = f.filename.replace(" ", "_")
                dest = task_dir / safe_name
                if dest.exists():
                    s_stem = Path(safe_name).stem
                    s_suf = Path(safe_name).suffix
                    dest = task_dir / f"{s_stem}_{len(saved_inputs)+1}{s_suf}"
                f.save(dest)
                saved_inputs.append(dest)

        if not saved_inputs:
            return jsonify({"error": "Nenhum arquivo válido enviado para conversão."}), 400

        CONVERSION_TASKS[task_id] = {
            "status": "processing",
            "progress": 0,
            "current_step": "Iniciando leitura do documento...",
            "page_info": "",
            "total_pages": 0,
            "processed_pages": 0,
            "started_at": time.time(),
            "completed_at": None,
            "filename": None,
            "filesize": None,
            "download_id": task_id,
            "error": None
        }
        _save_conversion_task_state(task_id, CONVERSION_TASKS[task_id], task_dir)

        user_name = session.get("user_name") or "usuario"
        threading.Thread(
            target=_async_converter_worker,
            args=(task_id, tool, saved_inputs, options, task_dir, user_name),
            daemon=True
        ).start()

        return jsonify({
            "success": True,
            "task_id": task_id,
            "status": "processing",
            "message": "Conversão iniciada no plano de fundo."
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Erro ao iniciar processamento: {str(e)}"}), 500


@app.get("/api/converter/status/<task_id>")
@login_required
def converter_status(task_id: str):
    task = CONVERSION_TASKS.get(task_id)
    task_dir = CONVERT_DIR / task_id

    # Fallback de persistencia em disco caso o processo tenha reiniciado
    if not task and task_dir.exists():
        state_file = task_dir / "task_state.json"
        if state_file.exists():
            try:
                task = json.loads(state_file.read_text(encoding="utf-8"))
                CONVERSION_TASKS[task_id] = task
            except Exception:
                pass

    if not task:
        # Se o diretorio da tarefa existe no disco, retornar status processing para nao quebrar polling
        if task_dir.exists():
            return jsonify({
                "task_id": task_id,
                "status": "processing",
                "progress": 0,
                "current_step": "Processando documento no servidor...",
                "page_info": "",
                "total_pages": 0,
                "processed_pages": 0,
                "elapsed_seconds": 0
            }), 200
        return jsonify({"error": "Tarefa de conversão não encontrada."}), 404

    elapsed = time.time() - task.get("started_at", time.time())
    res = dict(task)
    res["elapsed_seconds"] = int(elapsed)
    return jsonify(res)


@app.get("/api/converter/download/<download_id>")
@login_required
def converter_download(download_id: str):
    custom_name = request.args.get("custom_name", "").strip()
    file_info = CONVERTED_FILES.get(download_id)
    candidate_dir = CONVERT_DIR / download_id

    # Restaura metadados do arquivo a partir de task_state.json se ausente na memoria
    if not file_info and candidate_dir.exists():
        state_file = candidate_dir / "task_state.json"
        if state_file.exists():
            try:
                t_state = json.loads(state_file.read_text(encoding="utf-8"))
                fn = t_state.get("filename")
                if fn and (candidate_dir / fn).exists():
                    file_info = {
                        "path": str(candidate_dir / fn),
                        "download_name": fn,
                        "mimetype": "application/octet-stream",
                        "created_at": t_state.get("completed_at", time.time())
                    }
                    CONVERTED_FILES[download_id] = file_info
            except Exception:
                pass

    if not file_info:
        if candidate_dir.exists():
            excluded_suffixes = {".part", ".meta", ".tmp", ".json"}
            for p in candidate_dir.glob("*"):
                if p.is_file() and not p.name.startswith("input_") and p.suffix.lower() not in excluded_suffixes and p.name != "task_state.json":
                    dl_name = p.name
                    if custom_name:
                        ext = p.suffix
                        dl_name = custom_name if custom_name.lower().endswith(ext.lower()) else f"{custom_name}{ext}"
                    return send_file(p, as_attachment=True, download_name=dl_name)
        return jsonify({"error": "Arquivo não encontrado ou expirado."}), 404

    target_path = Path(file_info["path"])
    if not target_path.exists():
        return jsonify({"error": "Arquivo não encontrado no servidor."}), 404

    final_name = file_info["download_name"]
    if custom_name:
        ext = Path(file_info["download_name"]).suffix
        final_name = custom_name if custom_name.lower().endswith(ext.lower()) else f"{custom_name}{ext}"

    return send_file(
        target_path, 
        as_attachment=True, 
        download_name=final_name, 
        mimetype=file_info.get("mimetype")
    )



if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3969))
    from waitress import serve

    # Console do Windows pode ser cp1252: evita UnicodeEncodeError nos logs com acentos
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    print("=" * 70, flush=True)
    print(f" Servidor de enriquecimento ATIVO em http://localhost:{port}", flush=True)
    print(f" Concorrencia: {get_enrichment_concurrency()} | Logs abaixo em tempo real", flush=True)
    print(" Ctrl+C para parar. (Nao clique dentro desta janela: o Windows pausa o processo.)", flush=True)
    print("=" * 70, flush=True)

    serve(app, host="0.0.0.0", port=port)
