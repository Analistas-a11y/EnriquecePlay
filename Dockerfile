# Front do Enriquecimento (Padrao A): nginx serve /static e repassa o resto para o backend (BACKEND_URL).
# O backend (Flask/Waitress) continua rodando no servidor local e e alcancado pelo link do tunel.
# Nao ha etapa de build: os arquivos de static/ sao copiados como estao.
FROM nginx:1.27-alpine

# faz o entrypoint do nginx achar o DNS do conteiner (o nginx precisa dele para resolver o BACKEND_URL)
ENV NGINX_ENTRYPOINT_LOCAL_RESOLVERS=1
# BACKEND_URL e obrigatoria (sem barra no fim): o conteiner nao sobe sem ela e diz o motivo no log
ENV BACKEND_URL=""
# quem pode embutir este site num iframe (o Estrategia OS)
ENV FRAME_ANCESTORS="'self' https://estrategiaos.bookplay.com.br"

COPY docker/nginx.conf.template /etc/nginx/templates/default.conf.template
COPY docker/frame-headers.inc.template /etc/nginx/templates/frame-headers.inc.template
COPY docker/proxy-backend.conf /etc/nginx/proxy-backend.conf
COPY docker/10-exige-backend-url.sh /docker-entrypoint.d/10-exige-backend-url.sh
# tr tira o fim de linha do Windows, se o arquivo chegou com ele
RUN tr -d '\r' < /docker-entrypoint.d/10-exige-backend-url.sh > /tmp/exige.sh \
  && mv /tmp/exige.sh /docker-entrypoint.d/10-exige-backend-url.sh \
  && chmod +x /docker-entrypoint.d/10-exige-backend-url.sh

COPY static/ /usr/share/nginx/html/static/

EXPOSE 80
HEALTHCHECK --interval=30s --timeout=3s --start-period=10s --retries=3 CMD wget -qO- http://127.0.0.1/nginx-health >/dev/null 2>&1 || exit 1
