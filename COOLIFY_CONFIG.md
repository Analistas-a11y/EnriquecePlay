# Configuração na Coolify: Enriquecimento Nova Vida (EnriquecePlay)

Gerado em 2026-11-23 pelo agente. Branch com as mudanças: `deploy/coolify` (publicada só nesta branch; a `main` não foi alterada).

## 1. Resumo (30 segundos)
- **Padrão usado: A (front na Coolify, backend no servidor).** O sistema é Flask com HTML gerado no servidor (Jinja), então não existe um "front estático" separado. Por isso a Coolify roda um **nginx** que serve a pasta `static/` (CSS, JS, logos) e **repassa todo o resto** (páginas, login e `/api/...`) para o backend pelo **link do túnel** (`BACKEND_URL`). O backend continua no servidor, com os dados em disco, sem mudança.
- O que mudou no código:
  - `Dockerfile`, `.dockerignore`, `.gitattributes` e pasta `docker/` (nginx).
  - O endereço do Estratégia OS deixou de ser fixo: padrão agora é `https://estrategiaos.bookplay.com.br` (o da Vercel continua aceito na transição). Arquivos: `estrategia_integration.py`, `templates/login.html`, `app.py`.
  - O backend passou a ler o IP real do visitante de um cabeçalho do nginx (`CLIENT_IP_HEADER`). Sem a variável, tudo funciona como antes.
  - Senhas e `SECRET_KEY` não têm mais valor padrão no código (já tratado no commit anterior).
- O que o dono precisa fazer: **(1)** deixar o backend acessível por um túnel com endereço estável e anotar o link; **(2)** criar o recurso na Coolify com as tabelas das seções 2 e 3; **(3)** ajustar o `.env` do backend (seção 4) e reiniciar o `py app.py`; **(4)** apontar o domínio na Cloudflare (seção 6) e ajustar o Estratégia OS (seção 7).

## 2. Recurso na Coolify (tela "Repository configuration")
| Campo | Valor |
|---|---|
| Repository URL | `https://github.com/Analistas-a11y/EnriquecePlay.git` |
| Branch | `deploy/coolify` (depois de aprovada, trocar para `main`) |
| Build pack | Dockerfile |
| Base directory | `/` |
| Dockerfile | `Dockerfile` |
| Ports Exposes | `80` |
| Healthcheck (opcional) | `/nginx-health` |

## 3. Environment Variables da Coolify (só estas duas)
O contêiner da Coolify **não guarda senha nenhuma**. Todas as senhas e chaves ficam no `.env` do servidor (seção 4).

| Nome | Valor | Tipo | Obrigatória | Segredo | Onde pegar o valor |
|---|---|---|---|---|---|
| `BACKEND_URL` | `PENDENTE: link do túnel do backend, com https:// e SEM barra no fim` | Runtime | sim | não | o link do túnel que aponta para a porta 9376 do servidor (ex.: `https://api-enriquece.bookplay.com.br`) |
| `FRAME_ANCESTORS` | `'self' https://estrategiaos.bookplay.com.br` | Runtime | não (já é o padrão da imagem) | não | com `'self'` entre aspas simples. Usado só nos arquivos de `/static`; nas páginas quem envia o cabeçalho é o backend |

Sem `BACKEND_URL` o contêiner **não sobe** e escreve o motivo no log (de propósito). Mudou o valor: basta reiniciar (é variável de execução).

## 4. O que ajustar no `.env` do servidor (backend)
Hoje o `.env` ainda aponta só para o OS da Vercel e **tem prioridade sobre os padrões novos do código**. Para o login único funcionar com o domínio novo, edite estas linhas e reinicie o app (Ctrl+C e `py app.py`):

| Variável | Valor | Observação |
|---|---|---|
| `ESTRATEGIA_OS_URL` | `https://estrategiaos.bookplay.com.br` | de onde o backend busca a chave pública. Só troque quando o OS já estiver respondendo nesse domínio |
| `ESTRATEGIA_OS_ALLOWED_ORIGINS` | `https://estrategiaos.bookplay.com.br,https://estrategia-os.vercel.app,http://localhost:47821` | mantém a Vercel durante a transição; tire-a depois de alguns dias |
| `CLIENT_IP_HEADER` | `X-Original-Client-IP` | passa a registrar o IP real do visitante no histórico/auditoria. Sem isso, aparece o IP do servidor da Coolify |
| `SESSION_COOKIE_SECURE` / `SESSION_COOKIE_SAMESITE` | `1` / `None` | já estão assim; manter |
| `INTEGRATION_API_KEY` | **não mudar** | o Estratégia OS guarda o mesmo valor |
| `OS_PUBLIC_KEY_URL` | (opcional) `https://estrategiaos.bookplay.com.br/api/auth/public-key.pem` | só se a chave não estiver em `<ESTRATEGIA_OS_URL>/api/auth/public-key.pem` |

Demais variáveis lidas pelo código (já existem no seu `.env`, sem mudança): `PORT`, `SECRET_KEY`, `ADMIN_PASSWORD`, `SAF_PASSWORD`, `ESTRATEGIA_PASSWORD`, `NOVA_VIDA_USER`, `NOVA_VIDA_PASSWORD`, `NOVA_VIDA_CLIENT`, `NOVA_VIDA_BASE_URL`, `NOVA_VIDA_API_URL`, `NOVA_VIDA_API_KEY`/`NOVA_VIDA_TOKEN`, `ENRICHMENT_CONCURRENCY`, `ENRICHMENT_CHUNK_SIZE`, `QUEUE_CHUNK_SIZE`, `SUPABASE_URL`, `SUPABASE_KEY`. O modelo está em `env.template`.

## 5. Armazenamento persistente (Persistent Storage)
**Nenhum na Coolify.** O contêiner só serve arquivos estáticos e repassa requisições: não grava nada. Confirmado por leitura: o `Dockerfile` copia só `static/` e `docker/`. Todos os dados (`data/*.json`, `uploads/`, `exports/`, `chunk_uploads/`) continuam **no servidor**, onde o backend roda.

## 6. Domínio, túnel e Cloudflare
- **Backend (servidor):** manter `py app.py` rodando na porta 9376 (Waitress, `0.0.0.0`).
- **Túnel:** hoje o acesso externo é o encaminhamento de porta **Public** do VS Code (devtunnel). Isso funciona, mas o endereço depende do VS Code aberto e pode mudar. **Recomendado:** um **Cloudflare Tunnel nomeado** com endereço fixo, por exemplo `api-enriquece.<seu-domínio>` apontando para `http://localhost:9376`. Esse endereço é o `BACKEND_URL`.
- **Front (Coolify):** em **Domains**, `https://<subdomínio>.bookplay.com.br`. **Recomendado usar subdomínio de `bookplay.com.br`**, o mesmo domínio do Estratégia OS, para o cookie de sessão funcionar dentro do iframe sem depender de cookies de terceiros.
- **Cloudflare:** registro DNS do subdomínio do front apontando para o servidor da Coolify, proxy ligado. Se o proxy da Cloudflare ficar ligado na frente da Coolify, o nginx já lê o IP real pelo `CF-Connecting-IP`.
- Limites a lembrar: upload de **100 MB por requisição** no plano grátis da Cloudflare. O site já envia arquivos grandes em pedaços (`/api/upload/chunk`); confirme com um arquivo grande real. Conexões ociosas caem em ~100 s (o site usa consultas curtas, sem SSE).
- O nginx envia `X-Tunnel-Skip-AntiPhishing-Page`, para o devtunnel não mostrar a página de aviso no meio do fluxo.

## 7. O que mudar no Estratégia OS (feito pelo dono do OS, não pelo agente)
| Onde | O quê |
|---|---|
| `src/config/appsConfig.js` (campo `url` do app de enriquecimento) | trocar o endereço atual por `https://<novo domínio do front>/` |
| `.env` do servidor do OS | `<ID_DO_APP>_API_BASE_URL=https://<novo domínio do front>` (a `_API_KEY` **não muda**). O OS chama `/api/integracao/consultar` e agora essa chamada passa pelo nginx, que repassa o cabeçalho `X-API-Key` |
| `.env` do servidor do OS, linha `CORS_ORIGINS` | só se o navegador chama a API do OS a partir do novo domínio; reiniciar depois |

## 8. Como conferir depois de subir
| Teste | Esperado |
|---|---|
| `https://<domínio>/nginx-health` | `ok` |
| `https://<domínio>/` | redireciona para `/login` e mostra a tela de login (vem do backend) |
| `https://<domínio>/static/style.css` | o CSS, direto do nginx |
| `curl -I https://<domínio>/login` | traz `content-security-policy: frame-ancestors ...` incluindo `https://estrategiaos.bookplay.com.br` |
| Abrir o sistema **dentro do Estratégia OS** | carrega no iframe e entra sozinho (login único) |
| Entrar com usuário e senha | entra; ao recarregar continua logado |
| Subir uma planilha pequena nos 3 modos (subir fila / apenas enriquecer / enriquecer + subir fila) e baixar o resultado | funciona como hoje |
| Subir uma planilha grande (acima de 100 MB) | funciona (upload em pedaços) |
| `502 Bad Gateway` com `/nginx-health` ok | o nginx está de pé, mas não alcança o backend: `BACKEND_URL` errada, `py app.py` parado ou túnel fora |
| Histórico de auditoria mostra o IP | IP do visitante, não o do servidor (precisa de `CLIENT_IP_HEADER`) |

## 9. Diagnóstico e riscos
- **Stack:** Flask + Waitress (`app.py`), templates Jinja (`templates/`), JS e CSS em `static/`. Sem banco SQL: dados em `data/*.json` no disco do servidor; Supabase opcional (o projeto atual está fora do ar e o app já opera em modo local).
- **Streaming/tempo real:** não há SSE nem WebSocket; downloads e uploads em pedaços; timeouts do nginx em 900 s.
- **Cron/tarefas agendadas:** nenhuma externa (a fila roda dentro do processo do backend).
- **Verificado (executado de verdade):**
  - `app.py` e `enricher_engine.py` com sintaxe correta (`ast.parse`).
  - Numa cópia isolada do código, com o Flask test client: `GET /login` responde 200 com `Content-Security-Policy: frame-ancestors 'self' ... https://estrategiaos.bookplay.com.br` e sem `X-Frame-Options`; o `login.html` lista o domínio novo; `POST /login` responde 200 com cookie `Secure; SameSite=None; Partitioned; HttpOnly`; `/api/health` responde 200; com `CLIENT_IP_HEADER=X-Original-Client-IP` o IP lido é o do cabeçalho do nginx, e sem a variável o comportamento anterior é mantido.
- **Não verificado (e por quê):**
  - **`docker build` e `docker run`**: esta máquina não tem Docker. O `Dockerfile` e a configuração do nginx foram revisados só por leitura (cada `COPY` aponta para um arquivo que existe; não há `ARG` de build).
  - **`nginx -t`**: não há nginx aqui. O ponto de maior risco é a configuração de mapas do IP do cliente (`map` com regex) e o `try_files ... @backend` do `/static/`. Se o contêiner reiniciar em loop, olhe o log da Coolify: o `nginx` indica a linha do erro.
  - **Cabeçalho `frame-headers.inc`** (CSP dos arquivos estáticos): o guia também o marca como não testado.
  - **Fluxo completo atrás da Cloudflare/Traefik** (login único dentro do OS, cookie no iframe, upload grande): só dá para validar com o domínio no ar.
- **Endereços fixos que sobraram:**
  - `estrategia_integration.py` (`LEGACY_ESTRATEGIA_URL`): endereço antigo da Vercel mantido de propósito como plano de volta; remover depois da confirmação do dono.
  - `http://localhost:47821` (OS em desenvolvimento) continua aceito no código; não afeta produção.
  - `vercel.json` e `render.yaml`: não alterados (plano de volta).
- **Decisões que ficaram para o dono:**
  1. Túnel definitivo: devtunnel do VS Code ou Cloudflare Tunnel nomeado (recomendado).
  2. O link do túnel vira público: qualquer pessoa que o conheça acessa o backend sem passar pelo nginx. O login do sistema continua protegendo as telas e a API; se quiser fechar, use Cloudflare Access ou restrinja o túnel.
  3. O `static/` da imagem e o do servidor precisam ficar iguais: após mudar CSS/JS, faça **novo deploy** na Coolify. Se faltar um arquivo na imagem, o nginx pede ao backend (não quebra).
  4. As senhas de login (admin, SAF, Estratégia) ainda são as que estavam no código antigo; troque-as no `.env`.

## 10. Plano de volta
Nada do que existe foi removido: `vercel.json`, `render.yaml` e o acesso direto pelo link do túnel continuam como estão. Voltar é apontar de novo o endereço antigo no `appsConfig.js` do OS e no `<ID_DO_APP>_API_BASE_URL`. Desligar o acesso antigo só depois de alguns dias de uso normal.
