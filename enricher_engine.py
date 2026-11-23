# -*- coding: utf-8 -*-
"""
Enricher Engine - Motor Modular de Enriquecimento Nova Vida (API JSON 2026)
Compatível com a documentação Postman do WebService WSLocalizador (PESQUISAR 2421/2420).
"""

from __future__ import annotations

import argparse
import asyncio
import csv
import io
import json
import os
import re
import sys
import time
import unicodedata
from datetime import datetime, date, time as dt_time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import aiohttp
from openpyxl import Workbook, load_workbook
try:
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
except ImportError:
    ILLEGAL_CHARACTERS_RE = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")

# Constantes de campos enriquecidos
ENRICHED_FIELD_NAMES = [
    "enriquecido_cpf",
    "enriquecido_nome",
    "enriquecido_sexo",
    "enriquecido_nome_mae",
    "enriquecido_data_nascimento",
    "enriquecido_idade",
    "enriquecido_telefone",
    "enriquecido_telefone_2",
    "enriquecido_telefone_3",
    "enriquecido_telefone_4",
    "enriquecido_telefone_5",
    "enriquecido_whatsapp",
    "enriquecido_email",
    "enriquecido_logradouro",
    "enriquecido_numero",
    "enriquecido_complemento",
    "enriquecido_bairro",
    "enriquecido_cidade",
    "enriquecido_estado",
    "enriquecido_cep",
    "enriquecido_renda_estimada",
    "enriquecido_score",
    "enriquecido_perfil",
    "enriquecido_cbo",
    "enriquecido_profissao",
    "enriquecido_obito",
    "enriquecido_status",
    "tipo_enriquecimento",
    "_status_enriquecimento",
]


def normalize_text_key(text: str) -> str:
    """Remove acentos, pontuacoes e espacos extras para comparacao de chaves."""
    if not text:
        return ""
    nfkd = unicodedata.normalize("NFKD", str(text))
    clean = "".join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r"[^a-zA-Z0-9]", "", clean).lower()


def clean_name_string(name: Any) -> str:
    """Limpa e formata o nome para consulta na API."""
    if not name:
        return ""
    s = str(name).strip().upper()
    nfkd = unicodedata.normalize("NFKD", s)
    no_accents = "".join(c for c in nfkd if not unicodedata.combining(c))
    clean = re.sub(r"[^A-Z\s]", " ", no_accents)
    return re.sub(r"\s+", " ", clean).strip()


def normalize_uf(uf: Any) -> str:
    """Valida e normaliza a sigla do estado (UF)."""
    if not uf:
        return ""
    s = re.sub(r"[^a-zA-Z]", "", str(uf)).strip().upper()
    valid_ufs = {
        "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA",
        "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN",
        "RS", "RO", "RR", "SC", "SP", "SE", "TO"
    }
    return s if s in valid_ufs else ""


def clean_digits(value: Any) -> str:
    """Extrai apenas dígitos de um valor."""
    if not value:
        return ""
    return re.sub(r"\D", "", str(value))


class NovaVidaClient:
    """Cliente assíncrono para autenticação e consultas à API Nova Vida."""

    def __init__(
        self,
        base_url: Optional[str] = None,
        usuario: Optional[str] = None,
        senha: Optional[str] = None,
        cliente: Optional[str] = None,
    ):
        # Carrega de config.json se existir
        cfg = {}
        cfg_path = Path("config.json")
        if cfg_path.exists():
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                pass

        self.base_url = (base_url or cfg.get("NOVA_VIDA_API_URL") or "https://wsnv.novavidati.com.br").rstrip("/")
        self.usuario = (usuario or cfg.get("NOVA_VIDA_USER") or os.getenv("NOVA_VIDA_USER", ""))
        self.senha = (senha or cfg.get("NOVA_VIDA_PASSWORD") or os.getenv("NOVA_VIDA_PASSWORD", ""))
        self.cliente = (cliente or cfg.get("NOVA_VIDA_CLIENT") or os.getenv("NOVA_VIDA_CLIENT", ""))
        self.token: Optional[str] = None
        self.token_timestamp: float = 0.0
        self._lock = asyncio.Lock()

    async def get_token_async(self, session: aiohttp.ClientSession, force_refresh: bool = False) -> str:
        """Obtém ou renova o ws_token assincronamente via GerarTokenJson / GerarToken."""
        async with self._lock:
            now = time.time()
            if self.token and not force_refresh and (now - self.token_timestamp < 240.0):
                return self.token

            token_url_json = f"{self.base_url}/WSLocalizador.asmx/GerarTokenJson"
            payload = {
                "credencial": {
                    "usuario": self.usuario,
                    "senha": self.senha,
                    "cliente": self.cliente
                }
            }

            try:
                async with session.post(
                    token_url_json,
                    json=payload,
                    headers={"Content-Type": "application/json", "Accept": "application/json"},
                    timeout=aiohttp.ClientTimeout(total=8.0)
                ) as resp:
                    if resp.status == 200:
                        body = (await resp.text()).strip()
                        extracted = self._extract_token_from_body(body)
                        if extracted:
                            self.token = extracted
                            self.token_timestamp = time.time()
                            return self.token
            except Exception:
                pass

            # Fallback para GerarToken tradicional (form-urlencoded)
            token_url_legacy = f"{self.base_url}/WSLocalizador.asmx/GerarToken"
            form_data = {
                "usuario": self.usuario,
                "senha": self.senha,
                "cliente": self.cliente
            }
            try:
                async with session.post(
                    token_url_legacy,
                    data=form_data,
                    headers={"Content-Type": "application/x-www-form-urlencoded"},
                    timeout=aiohttp.ClientTimeout(total=8.0)
                ) as resp:
                    if resp.status == 200:
                        body = (await resp.text()).strip()
                        extracted = self._extract_token_from_body(body)
                        if extracted:
                            self.token = extracted
                            self.token_timestamp = time.time()
                            return self.token
            except Exception:
                pass

            return self.token or ""

    @staticmethod
    def _extract_token_from_body(body: str) -> Optional[str]:
        if not body:
            return None
        body = body.strip()
        token = None
        if body.startswith('"') and body.endswith('"'):
            try:
                token = json.loads(body)
            except Exception:
                pass
        elif body.startswith("{") or body.startswith("["):
            try:
                obj = json.loads(body)
                if isinstance(obj, dict):
                    token = obj.get("token") or obj.get("d") or obj.get("result")
            except Exception:
                pass
        else:
            m = re.search(r"<string[^>]*>([^<]+)</string>", body, re.IGNORECASE)
            if m:
                token = m.group(1).strip()
            elif len(body) > 15 and "<" not in body:
                token = body

        if token:
            clean_token = re.sub(r"[\r\n\t\s]", "", str(token))
            return clean_token if clean_token else None
        return None

    async def pesquisar_async(
        self,
        session: aiohttp.ClientSession,
        cpf: str = "",
        telefone: str = "",
        nome: str = "",
        uf: str = "",
        consulta: str = "2421"
    ) -> Optional[Dict[str, Any]]:
        """
        Executa a pesquisa individual no endpoint PESQUISAR.
        Retorna o dicionário parseado ou None.
        """
        clean_c = clean_digits(cpf)
        clean_p = clean_digits(telefone)
        clean_n = clean_name_string(nome)
        clean_u = normalize_uf(uf)

        if not (clean_c or clean_p or clean_n or clean_u):
            return None

        payload = {
            "CPF": clean_c,
            "TELEFONE": clean_p,
            "NOME": clean_n,
            "UF": clean_u
        }

        endpoint = f"{self.base_url}/WSLocalizador.asmx/PESQUISAR"
        backoffs = [0.20, 0.45, 0.90]

        for attempt in range(3):
            token = await self.get_token_async(session, force_refresh=(attempt > 0))
            if not token:
                await asyncio.sleep(0.3)
                continue

            headers = {
                "Content-Type": "application/json",
                "Accept": "application/json",
                "token": token,
                "consulta": str(consulta)
            }

            try:
                async with session.post(
                    endpoint,
                    json=payload,
                    headers=headers,
                    timeout=aiohttp.ClientTimeout(connect=3.5, sock_read=12.0)
                ) as resp:
                    if resp.status in (429, 500, 502, 503, 504):
                        await asyncio.sleep(backoffs[attempt] if attempt < len(backoffs) else 0.5)
                        continue

                    if resp.status == 200:
                        raw_text = (await resp.text()).strip()
                        if not raw_text or "token" in raw_text.lower() and "expir" in raw_text.lower():
                            await self.get_token_async(session, force_refresh=True)
                            continue
                        parsed_dict = self._parse_json_response(raw_text)
                        if parsed_dict:
                            return parsed_dict

                    elif resp.status == 204:
                        return None

                    return None
            except (aiohttp.ClientError, asyncio.TimeoutError):
                if attempt < 2:
                    await asyncio.sleep(backoffs[attempt])
                else:
                    return None

        return None

    @staticmethod
    def _parse_json_response(raw_text: str) -> Optional[Dict[str, Any]]:
        """Faz o parsing resiliente do JSON mesmo quando acompanhado de `{"d":null}`."""
        if not raw_text:
            return None
        try:
            decoder = json.JSONDecoder()
            obj, _ = decoder.raw_decode(raw_text)
            if isinstance(obj, list) and obj:
                obj = obj[0]
            if isinstance(obj, dict):
                return obj
        except Exception:
            pass
        return None


class EnrichmentNormalizer:
    """Mapeia e extrai a estrutura padronizada a partir da resposta da API Nova Vida."""

    @staticmethod
    def extract_enriched_data(api_dict: Optional[Dict[str, Any]]) -> Dict[str, Any]:
        """Converte a resposta da API 2421/2420 nos campos oficiais de enriquecimento."""
        if not api_dict or not isinstance(api_dict, dict):
            return {f: "" for f in ENRICHED_FIELD_NAMES if f != "_status_enriquecimento"} | {"_status_enriquecimento": "falha"}

        cpf = clean_digits(api_dict.get("CPF") or api_dict.get("DOCUMENTO") or "")
        nome = str(api_dict.get("NOME") or "").strip()
        sexo = str(api_dict.get("SEXO") or "").strip().upper()
        nome_mae = str(api_dict.get("NOMEMAE") or api_dict.get("NOME_MAE") or "").strip()
        
        nasc_raw = str(api_dict.get("NASCIMENTO") or api_dict.get("NASC") or "").strip()
        nasc_fmt = ""
        idade = ""
        if nasc_raw:
            # Formata data caso venha YYYY-MM-DD ou DD/MM/YYYY
            if re.match(r"^\d{4}-\d{2}-\d{2}", nasc_raw):
                parts = nasc_raw[:10].split("-")
                nasc_fmt = f"{parts[2]}/{parts[1]}/{parts[0]}"
            else:
                nasc_fmt = nasc_raw[:10]
            
            ym = re.search(r"\b\d{4}\b", nasc_fmt)
            if ym:
                try:
                    idade = str(datetime.now().year - int(ym.group(0)))
                except Exception:
                    idade = ""

        cbo = str(api_dict.get("CBO") or "").strip()
        profissao = str(api_dict.get("POSSIVEL_PROFISSAO") or api_dict.get("PROFISSAO") or "").strip()
        renda = str(api_dict.get("RENDA") or "").strip()
        status = str(api_dict.get("STATUS") or "REGULAR").strip()

        # Endereço estruturado
        end_obj = api_dict.get("ENDERECO") or api_dict.get("ENDERECOS") or {}
        if isinstance(end_obj, list) and end_obj:
            end_dict = end_obj[0] if isinstance(end_obj[0], dict) else {}
        elif isinstance(end_obj, dict):
            end_dict = end_obj
        else:
            end_dict = {}

        def _clean_str(val: Any) -> str:
            s = str(val or "").strip()
            return "" if s.upper() in ("NULL", "NONE") else s

        logradouro = _clean_str(end_dict.get("LOGRADOURO"))
        numero = _clean_str(end_dict.get("NUMERO"))
        complemento = _clean_str(end_dict.get("COMPLEMENTO"))
        bairro = _clean_str(end_dict.get("BAIRRO"))
        cidade = _clean_str(end_dict.get("CIDADE"))
        uf = _clean_str(end_dict.get("UF"))
        cep = _clean_str(end_dict.get("CEP"))

        # Contatos / Telefones
        contatos = api_dict.get("CONTATOS") or api_dict.get("TELEFONES") or api_dict.get("CELULARES") or []
        distinct_phones: List[str] = []
        if isinstance(contatos, list):
            for c in contatos:
                if isinstance(c, dict):
                    ddd = clean_digits(c.get("DDD"))
                    num = clean_digits(c.get("TELEFONE") or c.get("CELULAR") or c.get("NUMERO"))
                    full = ddd + num
                    if 10 <= len(full) <= 11 and full not in distinct_phones:
                        distinct_phones.append(full)
                elif isinstance(c, str):
                    digits = clean_digits(c)
                    if 10 <= len(digits) <= 11 and digits not in distinct_phones:
                        distinct_phones.append(digits)

        main_phone = distinct_phones[0] if distinct_phones else ""
        has_success = bool(cpf or nome or main_phone)

        return {
            "enriquecido_cpf": cpf,
            "enriquecido_nome": nome,
            "enriquecido_sexo": sexo,
            "enriquecido_nome_mae": nome_mae,
            "enriquecido_data_nascimento": nasc_fmt,
            "enriquecido_idade": idade,
            "enriquecido_telefone": main_phone,
            "enriquecido_telefone_2": distinct_phones[1] if len(distinct_phones) > 1 else "",
            "enriquecido_telefone_3": distinct_phones[2] if len(distinct_phones) > 2 else "",
            "enriquecido_telefone_4": distinct_phones[3] if len(distinct_phones) > 3 else "",
            "enriquecido_telefone_5": distinct_phones[4] if len(distinct_phones) > 4 else "",
            "enriquecido_whatsapp": "SIM" if main_phone else "NÃO",
            "enriquecido_email": "",
            "enriquecido_logradouro": logradouro,
            "enriquecido_numero": numero,
            "enriquecido_complemento": complemento,
            "enriquecido_bairro": bairro,
            "enriquecido_cidade": cidade,
            "enriquecido_estado": uf,
            "enriquecido_cep": cep,
            "enriquecido_renda_estimada": renda,
            "enriquecido_score": "",
            "enriquecido_perfil": "",
            "enriquecido_cbo": cbo,
            "enriquecido_profissao": profissao,
            "enriquecido_obito": "0",
            "enriquecido_status": status,
            "tipo_enriquecimento": "2421_COMPLETO" if "CONTATOS" in api_dict or "ENDERECO" in api_dict else "2420_SIMPLIFICADO",
            "_status_enriquecimento": "sucesso" if has_success else "falha"
        }


class FileFragmenter:
    """Lê, fragmenta em chunks e escreve planilhas Excel e CSV com baixo consumo de RAM."""

    @staticmethod
    def read_file(file_path: Path) -> List[Dict[str, Any]]:
        """Lê arquivo .xlsx ou .csv retornando lista de dicionários com as linhas."""
        ext = file_path.suffix.lower()
        if ext == ".csv":
            rows = []
            for encoding in ("utf-8-sig", "utf-8", "latin1", "cp1252"):
                try:
                    with open(file_path, "r", encoding=encoding) as f:
                        sample = f.read(4096)
                        f.seek(0)
                        delimiter = ";" if sample.count(";") >= sample.count(",") else ","
                        reader = csv.DictReader(f, delimiter=delimiter)
                        for r in reader:
                            rows.append({k.strip(): v for k, v in r.items() if k})
                    if rows:
                        return rows
                except Exception:
                    continue
            return rows

        elif ext == ".xlsx":
            wb = load_workbook(file_path, read_only=True, data_only=True)
            sheet = wb.active
            iter_rows = sheet.iter_rows(values_only=True)
            headers = None
            rows = []
            for r in iter_rows:
                if not headers:
                    if any(v is not None and str(v).strip() != "" for v in r):
                        headers = [str(v).strip() if v is not None else f"col_{i}" for i, v in enumerate(r)]
                    continue
                if not any(v is not None and str(v).strip() != "" for v in r):
                    continue
                row_dict = {}
                for idx, val in enumerate(r):
                    key = headers[idx] if idx < len(headers) else f"col_{idx}"
                    if isinstance(val, (datetime, date)):
                        row_dict[key] = val.strftime("%d/%m/%Y")
                    elif val is not None:
                        row_dict[key] = str(val).strip()
                    else:
                        row_dict[key] = ""
                rows.append(row_dict)
            wb.close()
            return rows

        raise ValueError(f"Formato de arquivo não suportado: {ext}")

    @staticmethod
    def chunk_list(lst: List[Any], chunk_size: int = 500) -> List[List[Any]]:
        """Divide uma lista em lotes menores."""
        return [lst[i:i + chunk_size] for i in range(0, len(lst), chunk_size)]

    @staticmethod
    def write_excel(output_path: Path, rows: List[Dict[str, Any]]) -> None:
        """Grava as linhas enriquecidas em um novo arquivo Excel .xlsx formatado."""
        if not rows:
            return

        wb = Workbook(write_only=True)
        ws = wb.create_sheet(title="Enriquecido")

        headers = list(rows[0].keys())
        clean_headers = [ILLEGAL_CHARACTERS_RE.sub("", str(h)).strip()[:32767] for h in headers]
        ws.append(clean_headers)

        for r in rows:
            row_vals = []
            for h in headers:
                v = r.get(h, "")
                if isinstance(v, str):
                    v = ILLEGAL_CHARACTERS_RE.sub("", v)
                    if len(v) > 32767:
                        v = v[:32767]
                row_vals.append(v)
            ws.append(row_vals)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(str(output_path))


class FastEnricherEngine:
    """Orquestrador principal de enriquecimento assíncrono em lote."""

    def __init__(
        self,
        concurrency: int = 25,
        consulta_mode: str = "2421",
        chunk_size: int = 500
    ):
        self.concurrency = max(1, min(100, concurrency))
        self.consulta_mode = consulta_mode
        self.chunk_size = chunk_size
        self.client = NovaVidaClient()

    async def process_rows_async(
        self,
        raw_rows: List[Dict[str, Any]],
        on_progress: Optional[Callable[[int, int, int, int], None]] = None
    ) -> List[Dict[str, Any]]:
        """Processa todas as linhas com controle estrito de concorrência e semáforo."""
        total = len(raw_rows)
        sem = asyncio.Semaphore(self.concurrency)
        connector = aiohttp.TCPConnector(limit=self.concurrency * 2, ssl=False, ttl_dns_cache=600)
        
        success_count = 0
        error_count = 0
        processed_count = 0
        enriched_results: List[Dict[str, Any]] = []

        async with aiohttp.ClientSession(connector=connector) as session:
            # Pré-aquecimento do token
            await self.client.get_token_async(session)

            async def process_one(row: Dict[str, Any]) -> Dict[str, Any]:
                nonlocal success_count, error_count, processed_count
                async with sem:
                    # Identifica os campos presentes na linha
                    cpf = ""
                    nome = ""
                    uf = ""
                    tel = ""
                    for k, v in row.items():
                        norm = normalize_text_key(k)
                        if norm in ("cpf", "documento", "doc") and not cpf:
                            cpf = str(v)
                        elif norm in ("nome", "cliente", "razaosocial", "titular") and not nome:
                            nome = str(v)
                        elif norm in ("uf", "estado", "siglauf") and not uf:
                            uf = str(v)
                        elif norm in ("telefone", "tel", "celular", "fone") and not tel:
                            tel = str(v)

                    # 1. Consulta principal (2421 Completo)
                    res = await self.client.pesquisar_async(
                        session=session,
                        cpf=cpf,
                        telefone=tel if not nome else "",
                        nome=nome,
                        uf=uf,
                        consulta=self.consulta_mode
                    )

                    # 2. Fallback de busca simplificada (2420) se não localizou
                    if not res and self.consulta_mode == "2421":
                        res = await self.client.pesquisar_async(
                            session=session,
                            cpf=cpf,
                            telefone=tel if not nome else "",
                            nome=nome,
                            uf=uf,
                            consulta="2420"
                        )

                    # 3. Fallback: Se tinha telefone e não achou por Nome+UF, tenta por Telefone isolado
                    if not res and tel:
                        res = await self.client.pesquisar_async(
                            session=session,
                            telefone=tel,
                            consulta="2421"
                        )

                    # Normalização dos dados de retorno
                    parsed = EnrichmentNormalizer.extract_enriched_data(res)
                    combined = dict(row) | parsed

                    if parsed.get("_status_enriquecimento") == "sucesso":
                        success_count += 1
                    else:
                        error_count += 1

                    processed_count += 1
                    if on_progress:
                        on_progress(processed_count, total, success_count, error_count)

                    return combined

            # Executa todos os registros em paralelo controlado
            tasks = [process_one(r) for r in raw_rows]
            enriched_results = await asyncio.gather(*tasks)

        return enriched_results

    def run(self, input_file: Path, output_file: Path) -> Dict[str, Any]:
        """Executa o pipeline completo: Leitura -> Chunks -> Enriquecimento -> Escrita Excel."""
        t0 = time.time()
        print(f"Lendo arquivo: {input_file}")
        raw_rows = FileFragmenter.read_file(input_file)
        total = len(raw_rows)
        print(f"Total de registros carregados: {total}")

        if total == 0:
            print("Arquivo vazio. Finalizando.")
            return {"total": 0, "sucessos": 0, "falhas": 0, "tempo": 0.0}

        def print_progress(proc: int, tot: int, succ: int, err: int):
            if proc % 20 == 0 or proc == tot:
                pct = (proc / tot) * 100
                print(f"Progresso: {proc}/{tot} ({pct:.1f}%) | Sucessos: {succ} | Falhas: {err}", end="\r")

        print(f"Iniciando enriquecimento (Concorrência: {self.concurrency}, Modo: {self.consulta_mode})...")
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            enriched_rows = loop.run_until_complete(
                self.process_rows_async(raw_rows, on_progress=print_progress)
            )
        finally:
            loop.close()

        print("\nGerando arquivo Excel final...")
        FileFragmenter.write_excel(output_file, enriched_rows)
        elapsed = time.time() - t0
        speed = total / elapsed if elapsed > 0 else 0

        success_total = sum(1 for r in enriched_rows if r.get("_status_enriquecimento") == "sucesso")
        error_total = total - success_total
        assertiveness = (success_total / total) * 100 if total > 0 else 0.0

        summary = {
            "total": total,
            "sucessos": success_total,
            "falhas": error_total,
            "assertividade_pct": round(assertiveness, 2),
            "tempo_segundos": round(elapsed, 2),
            "velocidade_linhas_segundo": round(speed, 2),
            "arquivo_saida": str(output_file)
        }
        return summary


def main():
    parser = argparse.ArgumentParser(description="Motor Modular de Enriquecimento Nova Vida (2026)")
    parser.add_argument("--arquivo", "-a", type=str, required=True, help="Caminho do arquivo de entrada (.xlsx ou .csv)")
    parser.add_argument("--saida", "-s", type=str, default="resultado_enriquecido.xlsx", help="Caminho do arquivo de saída (.xlsx)")
    parser.add_argument("--concorrencia", "-c", type=int, default=25, help="Quantidade de requisições simultâneas (padrão: 25)")
    parser.add_argument("--modo", "-m", type=str, default="2421", choices=["2421", "2420"], help="Modo de consulta (2421=Completo, 2420=Simplificado)")

    args = parser.parse_args()
    engine = FastEnricherEngine(concurrency=args.concorrencia, consulta_mode=args.modo)
    res = engine.run(Path(args.arquivo), Path(args.saida))
    print("\n" + "=" * 50)
    print("RESUMO DA EXECUÇÃO:")
    for k, v in res.items():
        print(f" - {k}: {v}")
    print("=" * 50)


if __name__ == "__main__":
    main()
