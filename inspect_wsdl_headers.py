"""
Inspeciona a definição do Header token no WSDL para PESQUISAR.
"""
import urllib.request
import ssl
import re

def check():
    context = ssl._create_unverified_context()
    with urllib.request.urlopen("https://wsnv.novavidati.com.br/WSLocalizador.asmx?WSDL", timeout=10, context=context) as resp:
        wsdl = resp.read().decode("utf-8", errors="ignore")
    
    # Procura headers do SOAP
    headers = re.findall(r'<soap:header\b[^>]*>', wsdl)
    print("Headers SOAP:", headers)

    # Procura definição de tipo de token no schema
    token_defs = re.findall(r'<s:element name="token"[^>]*>.*?</s:element>', wsdl, re.DOTALL)
    if not token_defs:
        token_defs = re.findall(r'<s:element name="[^"]*token[^"]*"[^>]*>.*?</s:element>', wsdl, re.DOTALL | re.IGNORECASE)
    print("Definições de token element:", token_defs)

    # Procura binding de PESQUISAR
    pesquisar_binding = re.findall(r'<wsdl:operation name="PESQUISAR">.*?</wsdl:operation>', wsdl, re.DOTALL)
    print("PESQUISAR binding:", pesquisar_binding)

if __name__ == "__main__":
    check()
