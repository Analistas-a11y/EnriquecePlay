#!/bin/sh
set -e

if [ -z "${BACKEND_URL}" ]; then
  echo "ERRO: a variável de ambiente BACKEND_URL não foi definida. Na Coolify: Environment Variables > BACKEND_URL=https://<endereço do backend (link do túnel)>" >&2
  exit 1
fi

case "${BACKEND_URL}" in
  http://*|https://*) ;;
  *)
    echo "ERRO: BACKEND_URL precisa começar com http:// ou https:// (recebi: ${BACKEND_URL})" >&2
    exit 1
    ;;
esac

case "${BACKEND_URL}" in
  */)
    echo "ERRO: BACKEND_URL não pode terminar com barra (recebi: ${BACKEND_URL})" >&2
    exit 1
    ;;
esac

echo "Backend configurado: ${BACKEND_URL}"
