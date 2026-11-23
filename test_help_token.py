import asyncio
import aiohttp
from app import get_nova_vida_token

async def test_help_with_token():
    token = get_nova_vida_token(force_refresh=True)
    base = "https://wsnv.novavidati.com.br/WSLocalizador.asmx"
    connector = aiohttp.TCPConnector(ssl=False)
    
    async with aiohttp.ClientSession(connector=connector) as session:
        # Testar com token no header e no body
        for c in ["2145", "2140", "2141", "2142", "2143", "2144", "2146", "2184", "1", "2", "3", "4", "5", "10", "20", "50", "100"]:
            async with session.post(
                f"{base}/PESQUISAR_HELP",
                data={"consulta": c},
                headers={"token": token}
            ) as r:
                t = await r.text()
                if "Token obrigat" not in t and "inv" not in t.lower():
                    print(f"\n==================== HELP {c} ====================")
                    print(t[:600])

asyncio.run(test_help_with_token())
