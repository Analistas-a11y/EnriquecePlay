# -*- coding: utf-8 -*-
import asyncio
import aiohttp
import time
from app import get_nova_vida_token, NOVA_VIDA_API_URL

async def test_concurrency(num_requests=30, concurrency=25):
    token = get_nova_vida_token()
    print(f"Token obtido: {token[:30]}...")
    endpoint = f"{NOVA_VIDA_API_URL.rstrip('/')}/WSLocalizador.asmx/PesquisaAtributosTK"
    nvcheck = f"{NOVA_VIDA_API_URL.rstrip('/')}/WSLocalizador.asmx/NVCHECK"
    
    connector = aiohttp.TCPConnector(limit=concurrency, ssl=False)
    timeout = aiohttp.ClientTimeout(connect=3.0, sock_read=8.0)
    
    sem = asyncio.Semaphore(concurrency)
    
    successes = 0
    errors = 0
    timeouts = 0
    
    async with aiohttp.ClientSession(connector=connector, timeout=timeout) as session:
        async def query_one(i):
            nonlocal successes, errors, timeouts
            async with sem:
                t0 = time.time()
                data = {
                    "nome": f"MARIA SILVA {i}",
                    "endereco": "",
                    "cidade": "",
                    "uf": "SP",
                    "telefone": "11987654321",
                    "celular": "11987654321",
                    "email": "",
                    "nascimento": "",
                    "token": token,
                }
                try:
                    async with session.post(endpoint, data=data) as resp:
                        dt = time.time() - t0
                        if resp.status == 200:
                            txt = await resp.text()
                            successes += 1
                            if i % 10 == 0:
                                print(f"Req {i}: HTTP 200 em {dt:.2f}s (len: {len(txt)})")
                        else:
                            errors += 1
                            print(f"Req {i}: HTTP {resp.status} em {dt:.2f}s")
                except asyncio.TimeoutError:
                    timeouts += 1
                    print(f"Req {i}: TIMEOUT após {time.time()-t0:.2f}s")
                except Exception as e:
                    errors += 1
                    print(f"Req {i}: ERRO: {e}")
                    
        t_start = time.time()
        tasks = [query_one(i) for i in range(num_requests)]
        await asyncio.gather(*tasks)
        t_total = time.time() - t_start
        print(f"--- RESULTADO COM CONCORRENCIA {concurrency} ---")
        print(f"Total: {num_requests} | Sucessos: {successes} | Erros: {errors} | Timeouts: {timeouts} | Tempo: {t_total:.2f}s ({num_requests/t_total:.1f} req/s)")

if __name__ == "__main__":
    asyncio.run(test_concurrency(num_requests=30, concurrency=25))
