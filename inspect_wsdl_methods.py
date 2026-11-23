"""
Inspeção detalhada de todas as operações e parâmetros no WSDL do WSLocalizador Nova Vida.
"""
import urllib.request
import ssl
import re
import xml.etree.ElementTree as ET

def inspect():
    context = ssl._create_unverified_context()
    wsdl_url = "https://wsnv.novavidati.com.br/WSLocalizador.asmx?WSDL"
    
    with urllib.request.urlopen(wsdl_url, timeout=15, context=context) as resp:
        wsdl_content = resp.read().decode("utf-8", errors="ignore")

    # Salva o WSDL para análise se necessário
    print("WSDL carregado com sucesso. Tamanho:", len(wsdl_content))

    # Parse XML
    root = ET.fromstring(wsdl_content)
    
    # Namespaces comuns no WSDL
    ns = {
        "wsdl": "http://schemas.xmlsoap.org/wsdl/",
        "s": "http://www.w3.org/2001/XMLSchema",
        "soap": "http://schemas.xmlsoap.org/wsdl/soap/",
        "soap12": "http://schemas.xmlsoap.org/wsdl/soap12/",
        "http": "http://schemas.xmlsoap.org/wsdl/http/",
        "mime": "http://schemas.xmlsoap.org/wsdl/mime/",
    }

    # Extrai todos os complexType / element
    elements = {}
    for elem in root.findall(".//s:element", ns):
        name = elem.get("name")
        if not name:
            continue
        # Busca parâmetros filhos
        params = []
        for p in elem.findall(".//s:element", ns):
            params.append((p.get("name"), p.get("type")))
        elements[name] = params

    print(f"Total de elementos definidos: {len(elements)}")

    # Filtra operações relevantes a Nome, Telefone, Atributos, Pesquisa, Consulta, etc.
    keywords = ["NOME", "TELEFONE", "CELULAR", "ATRIBUTOS", "PESQUISA", "LOCALIZA", "CONSULTA", "CADASTRO", "CHECK", "AFINIDADE", "VINCULO"]
    for el_name, params in sorted(elements.items()):
        if any(k in el_name.upper() for k in keywords) and not el_name.endswith("Response"):
            print(f"\nOperação/Elemento: {el_name}")
            print(f"  Parâmetros: {params}")

if __name__ == "__main__":
    inspect()
