#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Utilitario CLI de Observabilidade e Visualizacao de Logs.
Exibe de forma organizada logs de auditoria (acessos/IPs) e chamadas a APIs (consultas/erros).
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Generator, List, Optional

# Habilita suporte a codigos ANSI no Windows CMD/PowerShell
if sys.platform == "win32":
    os.system("")

# Diretorios e caminhos padrao
BASE_DIR = Path(__file__).resolve().parent
AUDIT_LOG_PATH = BASE_DIR / "audit_logs.jsonl"
API_DEBUG_PATH = BASE_DIR / "api_debug.log"

# Cores ANSI
class Cores:
    RESET = "\033[0m"
    BOLD = "\033[1m"
    DIM = "\033[2m"
    
    # Texto
    VERMELHO = "\033[91m"
    VERDE = "\033[92m"
    AMARELO = "\033[93m"
    AZUL = "\033[94m"
    MAGENTA = "\033[95m"
    CIANO = "\033[96m"
    BRANCO = "\033[97m"
    CINZA = "\033[90m"


def normalizar_timestamp(ts_str: str) -> Optional[datetime]:
    """Converte diferentes formatos de data e hora para datetime."""
    if not ts_str:
        return None
    formatos = [
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S,%f",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d",
    ]
    for fmt in formatos:
        try:
            return datetime.strptime(ts_str.strip(), fmt)
        except ValueError:
            continue
    return None


def formatar_timestamp_exibicao(dt: Optional[datetime], raw_str: str) -> str:
    """Formata data/hora para exibicao compacta e legivel."""
    if dt:
        return dt.strftime("%d/%m %H:%M:%S")
    return raw_str[:14] if len(raw_str) > 14 else raw_str


def ler_eventos_auditoria(audit_file: Path) -> List[Dict[str, Any]]:
    """Carrega e normaliza registros do arquivo audit_logs.jsonl."""
    eventos = []
    if not audit_file.exists():
        return eventos

    try:
        with open(audit_file, "r", encoding="utf-8", errors="replace") as f:
            for num_linha, linha in enumerate(f, 1):
                linha = linha.strip()
                if not linha:
                    continue
                try:
                    data = json.loads(linha)
                except Exception:
                    continue

                action = data.get("action", "EVENTO")
                user = data.get("user_name") or data.get("user_id") or "Desconhecido"
                ip = data.get("ip") or "-"
                ts_raw = data.get("timestamp", "")
                dt = normalizar_timestamp(ts_raw)
                details = data.get("details", {})

                status = "OK"
                desc = ""
                if action == "LOGIN_SUCCESS":
                    status = "OK"
                    desc = f"Login realizado com sucesso ({data.get('role', 'user')})"
                elif action == "LOGIN_FAILED":
                    status = "ERRO"
                    desc = f"Tentativa de login invalida para '{details.get('username', user)}'"
                elif action == "LOGOUT":
                    status = "INFO"
                    desc = "Encerrou a sessao"
                elif action == "UPLOAD_FILE":
                    status = "OK"
                    nome_arq = details.get("filename", "arquivo")
                    linhas = details.get("total_rows", 0)
                    dups = details.get("duplicates", 0)
                    score = details.get("quality_score", "")
                    desc = f"Upload: {nome_arq} ({linhas} linhas, {dups} dups, score={score}%)"
                elif action == "REQUEST_IMPORT":
                    status = "OK"
                    modo = details.get("mode", "")
                    total = details.get("total_rows", 0)
                    desc = f"Solicitacao de enriquecimento ({modo}): {total} linhas"
                elif action == "QUEUE_FINALIZED":
                    status = "OK"
                    fila = details.get("queue_number", "")
                    desc = f"Fila finalizada: #{fila}"
                elif action == "QUEUE_CANCELLED":
                    status = "AVISO"
                    desc = f"Fila cancelada: {details.get('request_id', '')[:8]}"
                elif action == "QUEUE_REJECTED":
                    status = "AVISO"
                    desc = f"Fila rejeitada: {details.get('request_id', '')[:8]}"
                elif action == "FILE_DOWNLOAD":
                    status = "OK"
                    desc = f"Download de exportacao: {details.get('filename', '')}"
                else:
                    status = "INFO"
                    desc = json.dumps(details, ensure_ascii=False) if details else action

                eventos.append({
                    "tipo": "ACESSO",
                    "origem": "AUDITORIA",
                    "timestamp": dt or datetime.min,
                    "timestamp_str": ts_raw,
                    "usuario": user,
                    "ip": ip,
                    "acao": action,
                    "status": status,
                    "detalhes": desc,
                    "raw": data
                })
    except Exception as e:
        print(f"Erro ao ler audit_logs: {e}", file=sys.stderr)

    return eventos


RE_API_LOG = re.compile(
    r"^\[(?P<ts>\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\]\s*"
    r"(?:\[(?P<level>[A-Z_]+)\]\s*)?"
    r"(?:\[IP:(?P<ip>[^\]]+)\]\s*)?"
    r"(?P<msg>.*)$"
)

def ler_eventos_api_debug(debug_file: Path) -> List[Dict[str, Any]]:
    """Carrega e normaliza registros do arquivo api_debug.log."""
    eventos = []
    if not debug_file.exists():
        return eventos

    try:
        with open(debug_file, "r", encoding="utf-8", errors="replace") as f:
            for linha in f:
                linha = linha.strip()
                if not linha:
                    continue

                m = RE_API_LOG.match(linha)
                if not m:
                    continue

                ts_raw = m.group("ts")
                level_raw = m.group("level")
                ip_raw = m.group("ip") or "-"
                msg = m.group("msg").strip()

                dt = normalizar_timestamp(ts_raw)

                status = "INFO"
                if level_raw:
                    lvl = level_raw.upper()
                    if lvl in ("ERRO", "ERROR", "CRITICAL"):
                        status = "ERRO"
                    elif lvl in ("AVISO", "WARN", "WARNING"):
                        status = "AVISO"
                    elif lvl in ("OK", "SUCCESS"):
                        status = "OK"
                else:
                    msg_low = msg.lower()
                    if any(k in msg_low for k in ["erro", "falha", "exception", "failed", "critical"]):
                        status = "ERRO"
                    elif any(k in msg_low for k in ["inacessivel", "aviso", "warning", "alerta"]):
                        status = "AVISO"
                    elif any(k in msg_low for k in ["sucesso", "concluído", "concluido", "iniciado"]):
                        status = "OK"

                msg_low = msg.lower()
                acao = "CONSULTA"
                if "token" in msg_low:
                    acao = "TOKEN"
                elif "motor" in msg_low:
                    acao = "MOTOR_ASYNC"
                elif "supabase" in msg_low:
                    acao = "SUPABASE"
                elif "enriquecer" in msg_low or "enriquecimento" in msg_low:
                    acao = "ENRIQUECIMENTO"
                elif "nvcheck" in msg_low:
                    acao = "NVCHECK"
                elif "upload" in msg_low:
                    acao = "UPLOAD"
                elif "download" in msg_low:
                    acao = "DOWNLOAD"

                eventos.append({
                    "tipo": "CONSULTA",
                    "origem": "API_DEBUG",
                    "timestamp": dt or datetime.min,
                    "timestamp_str": ts_raw,
                    "usuario": "Sistema",
                    "ip": ip_raw,
                    "acao": acao,
                    "status": status,
                    "detalhes": msg,
                    "raw": {"raw_line": linha}
                })
    except Exception as e:
        print(f"Erro ao ler api_debug.log: {e}", file=sys.stderr)

    return eventos


def filtrar_eventos(
    eventos: List[Dict[str, Any]],
    apenas_erros: bool = False,
    apenas_consultas: bool = False,
    apenas_acessos: bool = False,
    filtro_ip: Optional[str] = None,
    filtro_usuario: Optional[str] = None,
    busca: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Aplica os filtros especificados pelo usuario."""
    resultado = []
    
    ip_query = filtro_ip.strip().lower() if filtro_ip else None
    user_query = filtro_usuario.strip().lower() if filtro_usuario else None
    busca_query = busca.strip().lower() if busca else None

    for ev in eventos:
        if apenas_erros and ev["status"] not in ("ERRO", "AVISO"):
            continue

        if apenas_consultas and ev["tipo"] != "CONSULTA":
            continue

        if apenas_acessos and ev["tipo"] != "ACESSO":
            continue

        if ip_query and ip_query not in ev["ip"].lower():
            continue

        if user_query and user_query not in ev["usuario"].lower():
            continue

        if busca_query:
            texto_completo = f"{ev['acao']} {ev['detalhes']} {ev['usuario']} {ev['ip']} {ev['status']}".lower()
            if busca_query not in texto_completo:
                continue

        resultado.append(ev)

    return resultado


def formatar_linha_evento(ev: Dict[str, Any]) -> str:
    """Monta linha colorida e alinhada para saida no terminal."""
    ts_fmt = formatar_timestamp_exibicao(ev["timestamp"], ev["timestamp_str"])
    ts_col = f"{Cores.CINZA}[{ts_fmt}]{Cores.RESET}"

    orig = ev.get("origem", "SISTEMA")
    if orig == "AUDITORIA":
        orig_col = f"{Cores.CIANO}[ACESSO]{Cores.RESET}"
    else:
        orig_col = f"{Cores.AZUL}[API   ]{Cores.RESET}"

    st = ev.get("status", "INFO")
    if st == "OK":
        st_col = f"{Cores.VERDE}[OK   ]{Cores.RESET}"
    elif st == "AVISO":
        st_col = f"{Cores.AMARELO}[AVISO]{Cores.RESET}"
    elif st == "ERRO":
        st_col = f"{Cores.VERMELHO}{Cores.BOLD}[ERRO ]{Cores.RESET}"
    else:
        st_col = f"{Cores.CINZA}[INFO ]{Cores.RESET}"

    ip_raw = ev.get("ip", "-")
    ip_col = f"{Cores.MAGENTA}{ip_raw:<15}{Cores.RESET}"

    usr = ev.get("usuario", "Sistema")
    if len(usr) > 18:
        usr = usr[:17] + "."
    usr_col = f"{Cores.BRANCO}{usr:<18}{Cores.RESET}"

    acao = ev.get("acao", "ACAO")
    acao_col = f"{Cores.BOLD}{acao:<14}{Cores.RESET}"

    detalhes = ev.get("detalhes", "")
    if st == "ERRO":
        detalhes_col = f"{Cores.VERMELHO}{detalhes}{Cores.RESET}"
    elif st == "AVISO":
        detalhes_col = f"{Cores.AMARELO}{detalhes}{Cores.RESET}"
    else:
        detalhes_col = f"{detalhes}"

    return f"{ts_col} {orig_col} {st_col} {ip_col} {usr_col} {acao_col} {detalhes_col}"


def imprimir_cabecalho(filtros_ativos: List[str]) -> None:
    """Imprime cabecalho informativo das colunas."""
    sep = "=" * 110
    print(f"{Cores.CINZA}{sep}{Cores.RESET}")
    print(f"{Cores.BOLD}{Cores.BRANCO}PAINEL DE LOGS E OBSERVABILIDADE - CONSULTAS, ACESSOS E IPS{Cores.RESET}")
    if filtros_ativos:
        print(f"{Cores.AMARELO}Filtros ativos: {', '.join(filtros_ativos)}{Cores.RESET}")
    print(f"{Cores.CINZA}{sep}{Cores.RESET}")
    header_str = (
        f"{Cores.CINZA}[DATA/HORA]  [TIPO]   [STATUS] [IP CLIENTE]    [USUARIO]          [ACAO]         DETALHES{Cores.RESET}"
    )
    print(header_str)
    print(f"{Cores.CINZA}{'-' * 110}{Cores.RESET}")


def exportar_para_json(eventos: List[Dict[str, Any]], destino: Path) -> None:
    """Exporta lista de eventos normalizados para arquivo JSON identado."""
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        serializaveis = []
        for ev in eventos:
            item = dict(ev)
            if isinstance(item.get("timestamp"), datetime):
                item["timestamp"] = item["timestamp"].strftime("%Y-%m-%d %H:%M:%S")
            serializaveis.append(item)

        with open(destino, "w", encoding="utf-8") as f:
            json.dump({
                "gerado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "total_registros": len(serializaveis),
                "registros": serializaveis
            }, f, indent=2, ensure_ascii=False)
        print(f"\n{Cores.VERDE}[SUCESSO] Relatorio JSON exportado com sucesso para: {destino}{Cores.RESET}")
    except Exception as e:
        print(f"\n{Cores.VERMELHO}[ERRO] Falha ao exportar JSON: {e}{Cores.RESET}", file=sys.stderr)


def modo_tempo_real(
    audit_file: Path,
    debug_file: Path,
    apenas_erros: bool = False,
    apenas_consultas: bool = False,
    apenas_acessos: bool = False,
    filtro_ip: Optional[str] = None,
    filtro_usuario: Optional[str] = None,
    busca: Optional[str] = None,
) -> None:
    """Monitora continuamente em tempo real os novos registros dos arquivos."""
    filtros = []
    if apenas_erros:
        filtros.append("Apenas Erros/Falhas")
    if apenas_consultas:
        filtros.append("Apenas Consultas de API")
    if apenas_acessos:
        filtros.append("Apenas Acessos de Usuarios")
    if filtro_ip:
        filtros.append(f"IP={filtro_ip}")
    if filtro_usuario:
        filtros.append(f"Usuario={filtro_usuario}")
    if busca:
        filtros.append(f"Busca='{busca}'")
    filtros.append("Modo Tempo Real (Ctrl+C para encerrar)")

    imprimir_cabecalho(filtros)

    pos_audit = audit_file.stat().st_size if audit_file.exists() else 0
    pos_debug = debug_file.stat().st_size if debug_file.exists() else 0

    print(f"{Cores.CINZA}Aguardando novos eventos...{Cores.RESET}")

    try:
        while True:
            novos_eventos: List[Dict[str, Any]] = []

            # Checar audit_logs.jsonl
            if audit_file.exists():
                tam_atual = audit_file.stat().st_size
                if tam_atual > pos_audit:
                    with open(audit_file, "r", encoding="utf-8", errors="replace") as f:
                        f.seek(pos_audit)
                        for linha in f:
                            linha = linha.strip()
                            if not linha:
                                continue
                            try:
                                data = json.loads(linha)
                                action = data.get("action", "EVENTO")
                                user = data.get("user_name") or data.get("user_id") or "Desconhecido"
                                ip = data.get("ip") or "-"
                                ts_raw = data.get("timestamp", "")
                                dt = normalizar_timestamp(ts_raw)
                                details = data.get("details", {})
                                status = "OK" if "FAIL" not in action else "ERRO"
                                desc = json.dumps(details, ensure_ascii=False) if details else action
                                novos_eventos.append({
                                    "tipo": "ACESSO",
                                    "origem": "AUDITORIA",
                                    "timestamp": dt or datetime.now(),
                                    "timestamp_str": ts_raw,
                                    "usuario": user,
                                    "ip": ip,
                                    "acao": action,
                                    "status": status,
                                    "detalhes": desc,
                                    "raw": data
                                })
                            except Exception:
                                pass
                        pos_audit = f.tell()
                elif tam_atual < pos_audit:
                    pos_audit = 0

            # Checar api_debug.log
            if debug_file.exists():
                tam_atual = debug_file.stat().st_size
                if tam_atual > pos_debug:
                    with open(debug_file, "r", encoding="utf-8", errors="replace") as f:
                        f.seek(pos_debug)
                        for linha in f:
                            linha = linha.strip()
                            if not linha:
                                continue
                            m = RE_API_LOG.match(linha)
                            if not m:
                                continue
                            ts_raw = m.group("ts")
                            lvl_raw = m.group("level")
                            ip_raw = m.group("ip") or "-"
                            msg = m.group("msg").strip()
                            dt = normalizar_timestamp(ts_raw)

                            st = "INFO"
                            if lvl_raw:
                                l = lvl_raw.upper()
                                if l in ("ERRO", "ERROR", "CRITICAL"):
                                    st = "ERRO"
                                elif l in ("AVISO", "WARN", "WARNING"):
                                    st = "AVISO"
                                elif l in ("OK", "SUCCESS"):
                                    st = "OK"
                            else:
                                if any(k in msg.lower() for k in ["erro", "falha", "failed", "critical"]):
                                    st = "ERRO"
                                elif any(k in msg.lower() for k in ["aviso", "warning", "inacessivel"]):
                                    st = "AVISO"

                            novos_eventos.append({
                                "tipo": "CONSULTA",
                                "origem": "API_DEBUG",
                                "timestamp": dt or datetime.now(),
                                "timestamp_str": ts_raw,
                                "usuario": "Sistema",
                                "ip": ip_raw,
                                "acao": "CONSULTA",
                                "status": st,
                                "detalhes": msg,
                                "raw": {"raw_line": linha}
                            })
                        pos_debug = f.tell()
                elif tam_atual < pos_debug:
                    pos_debug = 0

            if novos_eventos:
                filtrados = filtrar_eventos(
                    novos_eventos,
                    apenas_erros=apenas_erros,
                    apenas_consultas=apenas_consultas,
                    apenas_acessos=apenas_acessos,
                    filtro_ip=filtro_ip,
                    filtro_usuario=filtro_usuario,
                    busca=busca
                )
                for ev in filtrados:
                    print(formatar_linha_evento(ev))

            time.sleep(1.0)
    except KeyboardInterrupt:
        print(f"\n{Cores.AMARELO}Monitoramento em tempo real encerrado pelo usuario.{Cores.RESET}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Painel CLI de Logs de Consultas, Acessos e IPs",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Exemplos de uso:
  python logs_cli.py                     (ultimos 30 eventos unificados)
  python logs_cli.py -e                  (apenas erros e falhas)
  python logs_cli.py -a                  (apenas acessos de usuarios)
  python logs_cli.py -c                  (apenas consultas de enriquecimento/API)
  python logs_cli.py --ip 192.168.1.10   (filtrar por IP especifico)
  python logs_cli.py -s "Paulo"          (pesquisar por texto/CPF/usuario)
  python logs_cli.py -f                  (modo tempo real continuo - live tail)
  python logs_cli.py --exportar relatorio.json (exportar registros filtrados para JSON)
"""
    )

    parser.add_argument("-e", "--erros", action="store_true", help="Exibe apenas erros e falhas")
    parser.add_argument("-c", "--consultas", action="store_true", help="Exibe apenas consultas e operacoes de API")
    parser.add_argument("-a", "--acessos", action="store_true", help="Exibe apenas acessos e auditoria de usuarios")
    parser.add_argument("--ip", type=str, help="Filtra eventos por endereco IP especifico")
    parser.add_argument("-u", "--usuario", type=str, help="Filtra eventos por nome ou id de usuario")
    parser.add_argument("-s", "--busca", type=str, help="Pesquisa textual por CPF, arquivo, ID ou termo")
    parser.add_argument("-n", "--linhas", type=int, default=30, help="Quantidade de registros a exibir (padrao: 30)")
    parser.add_argument("-f", "--seguir", action="store_true", help="Acompanha os logs em tempo real (estilo tail -f)")
    parser.add_argument("--exportar", type=str, help="Caminho do arquivo para exportar os logs filtrados em formato JSON")

    args = parser.parse_args()

    if args.seguir:
        modo_tempo_real(
            AUDIT_LOG_PATH,
            API_DEBUG_PATH,
            apenas_erros=args.erros,
            apenas_consultas=args.consultas,
            apenas_acessos=args.acessos,
            filtro_ip=args.ip,
            filtro_usuario=args.usuario,
            busca=args.busca,
        )
        return

    # Modo estatico: carregar eventos existentes
    eventos_audit = ler_eventos_auditoria(AUDIT_LOG_PATH)
    eventos_api = ler_eventos_api_debug(API_DEBUG_PATH)

    todos_eventos = eventos_audit + eventos_api
    todos_eventos.sort(key=lambda x: x["timestamp"])

    filtrados = filtrar_eventos(
        todos_eventos,
        apenas_erros=args.erros,
        apenas_consultas=args.consultas,
        apenas_acessos=args.acessos,
        filtro_ip=args.ip,
        filtro_usuario=args.usuario,
        busca=args.busca,
    )

    total_encontrado = len(filtrados)
    exibir = filtrados[-args.linhas:] if args.linhas > 0 else filtrados

    filtros_ativos = []
    if args.erros:
        filtros_ativos.append("Erros/Falhas")
    if args.consultas:
        filtros_ativos.append("Consultas/API")
    if args.acessos:
        filtros_ativos.append("Acessos/Usuarios")
    if args.ip:
        filtros_ativos.append(f"IP={args.ip}")
    if args.usuario:
        filtros_ativos.append(f"Usuario={args.usuario}")
    if args.busca:
        filtros_ativos.append(f"Busca='{args.busca}'")

    imprimir_cabecalho(filtros_ativos)

    if not exibir:
        print(f"{Cores.AMARELO}Nenhum registro encontrado para os filtros informados.{Cores.RESET}")
    else:
        for ev in exibir:
            print(formatar_linha_evento(ev))

    print(f"{Cores.CINZA}{'-' * 110}{Cores.RESET}")
    print(f"Exibindo {len(exibir)} de {total_encontrado} eventos encontrados.")

    if args.exportar:
        exportar_para_json(filtrados, Path(args.exportar))


if __name__ == "__main__":
    main()
