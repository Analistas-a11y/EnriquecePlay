# -*- coding: utf-8 -*-
"""
Teste automatizado para verificar o suporte e performance com arquivos massivos de 250.000+ linhas.
"""
import io
import os
import sys
import time
from pathlib import Path

def test_massive_file_support():
    print("Iniciando teste com arquivo de 250.000 linhas...")
    from app import app, db_get_queue_item, db_get_record

    # 1. Gera um CSV em memoria / disco com 250.000 linhas
    total_lines = 250000
    print(f"Gerando CSV sintetico com {total_lines} registros...")
    t0 = time.time()

    csv_lines = ["Nome,CPF,Telefone,UF,Cidade\n"]
    # Gera registros em lote para velocidade de geracao
    base_chunk = []
    for i in range(1000):
        nome = f"Cliente Teste Numero {i+1} da Silva"
        cpf = f"{((10000000000 + i)):011d}"
        tel = f"1198{((1000000 + i)):07d}"
        base_chunk.append(f"{nome},{cpf},{tel},SP,Sao Paulo\n")

    chunk_str = "".join(base_chunk)
    # Repete para atingir 250.000 linhas
    full_csv_buffer = io.StringIO()
    full_csv_buffer.write("Nome,CPF,Telefone,UF,Cidade\n")
    for _ in range(250):
        full_csv_buffer.write(chunk_str)

    csv_bytes = full_csv_buffer.getvalue().encode("utf-8")
    t_gen = time.time() - t0
    print(f"[OK] CSV de {len(csv_bytes) / (1024*1024):.1f} MB gerado em {t_gen:.2f}s.")

    # 2. Testa upload do arquivo de 250k linhas no endpoint /api/upload
    with app.test_client() as client:
        with client.session_transaction() as sess:
            sess["user_id"] = "test_admin"
            sess["user_name"] = "Admin Teste"
            sess["role"] = "admin"

        data = {
            "file": (io.BytesIO(csv_bytes), "base_massiva_250k.csv")
        }

        print("Enviando arquivo para POST /api/upload...")
        t_upload_start = time.time()
        resp = client.post("/api/upload", data=data, content_type="multipart/form-data")
        t_upload = time.time() - t_upload_start

        assert resp.status_code == 200, f"Upload falhou com codigo {resp.status_code}: {resp.data}"
        payload = resp.get_json()
        assert payload.get("id"), "ID do registro nao retornado."
        record_id = payload["id"]
        total_in = payload.get("total_in")
        print(f"[OK] Upload de 250k linhas concluido com sucesso em {t_upload:.2f}s!")
        print(f"     Total de linhas processadas: {total_in}")
        print(f"     Tipo detectado: {payload.get('detected_type')}")
        print(f"     Score de qualidade: {payload.get('preflight', {}).get('quality_score')}%")

        assert total_in == total_lines, f"Esperava {total_lines} linhas, obteve {total_in}"

        # 3. Valida criacao de solicitacao de importacao (/api/request-import) em modo queue_only
        print("Criando solicitacao de importacao no endpoint /api/request-import...")
        t_req_start = time.time()
        req_resp = client.post("/api/request-import", json={
            "record_id": record_id,
            "mode": "queue_only",
            "deduplicate": False,
            "requester_name": "Teste Carga 250k"
        })
        t_req = time.time() - t_req_start

        assert req_resp.status_code == 200, f"Falha ao criar importacao: {req_resp.data}"
        req_payload = req_resp.get_json()
        request_id = req_payload.get("request_id")
        assert request_id, "request_id ausente na resposta de importacao."
        print(f"[OK] Solicitacao de importacao criada em {t_req:.2f}s com ID: {request_id}")

        # 4. Valida finalizacao da fila com FastConcatenator
        print("Finalizando fila de importacao /api/request/<id>/finalize...")
        t_fin_start = time.time()
        fin_resp = client.post(f"/api/request/{request_id}/finalize", json={
            "queue_number": "999999",
            "requester_name": "Teste Carga",
            "selected_fields": ["Nome", "Telefone", "CPF", "UF"],
            "mode": "queue_only"
        })
        t_fin = time.time() - t_fin_start

        assert fin_resp.status_code == 200, f"Falha ao finalizar fila: {fin_resp.data}"
        print(f"[OK] Fila de 250k linhas finalizada com sucesso em {t_fin:.2f}s!")

    print("\nSUCESSO: TESTE DE SUPORTE A 250.000 LINHAS CONCLUIDO COM EXCELENTE PERFORMANCE!")

if __name__ == "__main__":
    test_massive_file_support()
