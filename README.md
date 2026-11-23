# Nova Vida - Enriquecimento

Aplicação MVP para validar planilhas, detectar o tipo de enriquecimento, permitir escolha dos campos a importar, montar uma fila de solicitação e acompanhar a conclusão da importação pela equipe de estratégia.

## Como rodar no Windows

1. Se ainda não existir `.venv`, crie o ambiente:
   `uv venv --python 3.12 --seed --no-project .venv`
2. Instale as dependências no ambiente virtual:
   `.\.venv\Scripts\python.exe -m pip install -r requirements.txt`
3. Inicie a aplicação pelo terminal:
   `.\.venv\Scripts\python.exe app.py`
4. Acesse `http://localhost:9376/login` (a porta é lida de `PORT` no `.env`).

A aplicação usa Waitress e escuta em `0.0.0.0` para permitir encaminhamento da porta.
No VS Code, abra **Ports**, encaminhe a porta configurada em `PORT` e escolha **Public**
na visibilidade quando quiser compartilhar o acesso.

As consultas à Nova Vida são processadas pela fila global. A concorrência por lote vem de
`ENRICHMENT_CONCURRENCY` (padrão 80, teto 200); respostas HTTP e falhas de rede são
resumidas em `api_debug.log` (linha `Respostas da API Nova Vida`), e o motivo de exceções
aparece como `[DIAG]`.

Credenciais da Nova Vida: o `config.json` tem **prioridade** sobre o `.env`. Para trocar a
senha, altere só o `.env` (não mantenha `NOVA_VIDA_PASSWORD` no `config.json`) e reinicie o app.

## Usuários e senhas

Os logins (`admin`, `saf`, `estrategia`) usam as senhas definidas no `.env`
(`ADMIN_PASSWORD`, `SAF_PASSWORD`, `ESTRATEGIA_PASSWORD`) e a sessão usa `SECRET_KEY`.
Se alguma não estiver definida, o app gera um valor aleatório temporário e mostra no terminal.
Use `env.template` como modelo do seu `.env`.

## Fluxo principal

- Usuário 1 envia uma planilha em Excel/CSV.
- O backend identifica o tipo de enriquecimento com base no nome das colunas.
- O sistema valida o arquivo e simula o enriquecimento.
- O usuário escolhe os campos a importar por drag-and-drop.
- A solicitação é enviada para fila com número de fila, nome do solicitante e observações.
- A equipe de estratégia visualiza a fila, faz o download da exportação e conclui a importação anexando o resumo.
- A notificação retorna ao usuário 1 como confirmação.

## Integração com a API da Nova Vida

O backend agora tenta chamar a API da Nova Vida quando as variáveis de ambiente estiverem configuradas:

- `NOVA_VIDA_API_URL` — URL base da API, por exemplo `https://api.novavida.com.br`
- `NOVA_VIDA_API_ENDPOINT` — endpoint do enriquecimento, por exemplo `/api/enrichment`
- `NOVA_VIDA_API_KEY` ou `NOVA_VIDA_TOKEN` — token de autenticação

Exemplo:

```bash
set NOVA_VIDA_API_URL=https://api.novavida.com.br
set NOVA_VIDA_API_ENDPOINT=/api/enrichment
set NOVA_VIDA_API_KEY=sua-chave
python app.py
```

Se a API não estiver configurada ou responder com erro, o sistema mantém o comportamento de fallback em mock para não quebrar o fluxo de teste local.

## Observações

- A integração com a API da Nova Vida foi preparada em código para uso real, com fallback seguro para ambiente local.
- Os dados são mantidos em memória durante a execução da aplicação.
