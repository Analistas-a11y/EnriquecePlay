# -*- coding: utf-8 -*-
"""
Teste automatizado para validar o particionamento em sub-lotes (batch chunking)
e a recombinação automática no motor assíncrono de enriquecimento.
"""
import asyncio
import csv
import os
import sys
import uuid
from pathlib import Path

# Configura tamanho de sub-lote pequeno para testar particionamento com múltiplas fatias
os.environ["ENRICHMENT_CHUNK_SIZE"] = "50"

from app import (
    app,
    _async_enrichment_engine,
    db_save_queue_item,
    db_get_queue_item,
    EXPORT_DIR,
)

def run_test():
    print("[TEST] Iniciando teste de particionamento e recombinação em sub-lotes...")
    test_id = f"test_chunk_{uuid.uuid4().hex[:8]}"

    # 120 registros sintéticos (50 + 50 + 20 -> 3 sub-lotes)
    total_records = 120
    raw_rows = []
    for i in range(total_records):
        raw_rows.append({
            "Nome": f"Pessoa Teste {i+1}",
            "CPF": f"{10000000000 + i:011d}",
            "Telefone": f"1198{1000000 + i:07d}",
            "UF": "SP",
            "Cidade": "Sao Paulo"
        })

    queue_item = {
        "id": test_id,
        "queue_number": "999991",
        "requester_name": "Teste Automatizado Chunking",
        "observacoes": "Teste de Sub-lotes",
        "selected_fields": ["Nome", "Telefone", "CPF", "UF"],
        "query_fields": ["Nome", "UF", "Telefone", "CPF"],
        "request_time": "15/10/2026 10:00:00",
        "status": "processing",
        "total_rows": total_records,
        "processed_count": 0,
        "success_count": 0,
        "error_count": 0,
        "logs": [],
        "rows": [],
        "completed_at": None,
        "mode": "enrich_and_queue",
    }
    db_save_queue_item(queue_item)

    print(f"[TEST] Executando motor assíncrono para {total_records} registros com ENRICHMENT_CHUNK_SIZE=50...")
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        loop.run_until_complete(
            _async_enrichment_engine(
                test_id,
                raw_rows,
                "NOME_UF",
                ["Nome", "UF", "Telefone", "CPF"]
            )
        )
    finally:
        loop.close()

    # Validações pós-execução
    item = db_get_queue_item(test_id)
    assert item is not None, "Item de fila não encontrado após execução."
    print(f"[OK] Status final do item: {item.get('status')}")
    assert item.get("status") == "pending", f"Esperava status pending, obteve {item.get('status')}"

    chunk_info = item.get("chunk_info", {})
    print(f"[OK] Informações de sub-lote registradas: {chunk_info}")
    assert chunk_info.get("is_chunked") is True, "Deveria estar marcado como is_chunked."
    assert chunk_info.get("total_chunks") == 3, f"Esperava 3 sub-lotes, obteve {chunk_info.get('total_chunks')}"
    assert item.get("processed_count") == total_records, f"Esperava {total_records} processados, obteve {item.get('processed_count')}"

    # Validação do arquivo CSV mestre consolidado
    csv_master = EXPORT_DIR / f"raw_{test_id}.csv"
    assert csv_master.exists(), f"Arquivo CSV mestre consolidado não encontrado em {csv_master}"

    with open(csv_master, "r", encoding="utf-8") as f:
        reader = list(csv.DictReader(f, delimiter=";"))
        print(f"[OK] Linhas no arquivo CSV consolidado mestre: {len(reader)}")
        assert len(reader) == total_records, f"Esperava {total_records} linhas no consolidado, obteve {len(reader)}"

    # Validação dos arquivos de partes (sub-lotes)
    part1 = EXPORT_DIR / f"raw_{test_id}_part1.csv"
    part2 = EXPORT_DIR / f"raw_{test_id}_part2.csv"
    part3 = EXPORT_DIR / f"raw_{test_id}_part3.csv"

    assert part1.exists(), f"Sub-lote parte 1 ausente: {part1}"
    assert part2.exists(), f"Sub-lote parte 2 ausente: {part2}"
    assert part3.exists(), f"Sub-lote parte 3 ausente: {part3}"

    with open(part1, "r", encoding="utf-8") as f:
        rows_p1 = list(csv.DictReader(f, delimiter=";"))
        assert len(rows_p1) == 50, f"Parte 1 deveria ter 50 linhas, tem {len(rows_p1)}"

    with open(part2, "r", encoding="utf-8") as f:
        rows_p2 = list(csv.DictReader(f, delimiter=";"))
        assert len(rows_p2) == 50, f"Parte 2 deveria ter 50 linhas, tem {len(rows_p2)}"

    with open(part3, "r", encoding="utf-8") as f:
        rows_p3 = list(csv.DictReader(f, delimiter=";"))
        assert len(rows_p3) == 20, f"Parte 3 deveria ter 20 linhas, tem {len(rows_p3)}"

    print("[OK] Todas as 3 partes validadas com os tamanhos exatos (50, 50, 20).")

    # Testa endpoint de download de parte individual e consolidada via test_client
    with app.test_client() as client:
        with client.session_transaction() as sess:
            sess["user_id"] = "test_user"
            sess["user_name"] = "Usuario Teste"
            sess["role"] = "admin"

        # Download da parte 1
        resp_p1 = client.get(f"/api/request/{test_id}/download?part=1")
        assert resp_p1.status_code == 200, f"Falha no download da parte 1: {resp_p1.status_code}"
        assert resp_p1.headers.get("Content-Type") == "text/csv; charset=utf-8" or "text/csv" in resp_p1.headers.get("Content-Type", "")

        # Logs detalhados
        resp_logs = client.get(f"/api/request/{test_id}/logs")
        assert resp_logs.status_code == 200
        logs_data = resp_logs.get_json()
        assert logs_data.get("chunk_info", {}).get("total_chunks") == 3
        print(f"[OK] Logs retornam chunk_info correto com 3 sub-lotes.")

    # Limpeza dos arquivos de teste
    for p in [csv_master, part1, part2, part3]:
        try:
            if p.exists():
                p.unlink()
        except Exception:
            pass

    print("[TEST SUCESSO] Particionamento em sub-lotes e recombinação automática 100% aprovados!")

if __name__ == "__main__":
    run_test()
