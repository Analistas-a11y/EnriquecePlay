"""
Scanner assíncrono ultra-rápido de métodos e códigos de consulta da Nova Vida.
"""
import asyncio
import aiohttp
import re
import app

async def scan():
    token = app.get_nova_vida_token(force_refresh=True)
    print("Token gerado:", token[:35])

    # Extrai códigos do token
    codes_str = app.base64.b64decode(token).decode("utf-8", errors="ignore")
    m = re.search(r"\|((?:\d+_True;)+)\|", codes_str)
    active_codes = []
    if m:
        for p in m.group(1).split(";"):
            if "_True" in p:
                c = p.split("_")[0]
                if c.isdigit():
                    active_codes.append(int(c))
    print(f"Códigos ativos ({len(active_codes)}): {active_codes}")

    connector = aiohttp.TCPConnector(ssl=False, limit=100)
    async with aiohttp.ClientSession(connector=connector) as session:
        # 1. Testar CONSULTAR com documento="SILVIO SANTOS" para todos os códigos em paralelo
        url_consultar = "https://wsnv.novavidati.com.br/WSLocalizador.asmx/CONSULTAR"
        
        async def test_code_nome(code):
            data = {"consulta": str(code), "documento": "SILVIO SANTOS", "token": token}
            try:
                async with session.post(url_consultar, data=data, timeout=aiohttp.ClientTimeout(total=5.0)) as resp:
                    text = await resp.text()
                    if "CONSULTA N" not in text and "NÃO DISPONIVEL" not in text and "NAO DISPONIVEL" not in text and len(text) > 100:
                        print(f"\n[SUCESSO] Código {code} com NOME 'SILVIO SANTOS' retornou:\n{text[:300]}")
            except Exception:
                pass

        async def test_code_help(code):
            url_help = "https://wsnv.novavidati.com.br/WSLocalizador.asmx/PESQUISAR_HELP"
            data = {"consulta": str(code)}
            try:
                async with session.post(url_help, data=data, timeout=aiohttp.ClientTimeout(total=4.0)) as resp:
                    text = await resp.text()
                    clean = re.sub(r"<[^>]+>", " ", text).strip()
                    if clean and "System." not in clean and len(clean) > 3:
                        print(f"Help Consulta {code}: {clean}")
            except Exception:
                pass

        print("\n--- Testando PESQUISAR_HELP para identificar nome de cada consulta ---")
        await asyncio.gather(*(test_code_help(c) for c in active_codes))

        print("\n--- Testando CONSULTAR com NOME ---")
        await asyncio.gather(*(test_code_nome(c) for c in active_codes))

if __name__ == "__main__":
    asyncio.run(scan())
