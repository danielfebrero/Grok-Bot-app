#!/usr/bin/env bash
# setup-ca.sh — génère la CA mitmproxy puis l'installe dans le trust store
# système de la box. Idempotent.
#
#   1. lance `mitmdump` quelques secondes pour créer ~/.mitmproxy/mitmproxy-ca-cert.pem ;
#   2. copie le cert vers /usr/local/share/ca-certificates/mitmproxy.crt (sudo) ;
#   3. update-ca-certificates.
#
# Le binaire mitmdump est pris dans le venv dédié (MITM_VENV, défaut
# /workspace/.venvs/mitmproxy) sinon dans le PATH.
set -euo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
VENV="${MITM_VENV:-/workspace/.venvs/mitmproxy}"
MITMDUMP="$VENV/bin/mitmdump"; [ -x "$MITMDUMP" ] || MITMDUMP="$(command -v mitmdump || true)"
[ -n "$MITMDUMP" ] || { echo "ERREUR: mitmdump introuvable (installe mitmproxy dans $VENV)"; exit 1; }
CONFDIR="${MITMPROXY_CONFDIR:-$HOME/.mitmproxy}"
CA_PEM="$CONFDIR/mitmproxy-ca-cert.pem"

if [ ! -f "$CA_PEM" ]; then
  echo "[setup-ca] génération de la CA mitmproxy dans $CONFDIR …"
  mkdir -p "$CONFDIR"
  timeout 8 "$MITMDUMP" --set confdir="$CONFDIR" -p 0 -q >/dev/null 2>&1 || true
fi
[ -f "$CA_PEM" ] || { echo "ERREUR: CA non générée ($CA_PEM absent)"; exit 1; }
echo "[setup-ca] CA présente : $CA_PEM"

DEST="/usr/local/share/ca-certificates/mitmproxy.crt"
if [ -f "$DEST" ] && cmp -s "$CA_PEM" "$DEST"; then
  echo "[setup-ca] CA déjà installée dans le trust store système (idempotent)."
else
  echo "[setup-ca] installation dans le trust store système (sudo)…"
  sudo cp "$CA_PEM" "$DEST"
  sudo update-ca-certificates >/dev/null 2>&1 || sudo update-ca-certificates || true
  echo "[setup-ca] installée : $DEST"
fi
echo "[setup-ca] OK. Pour curl : --cacert $CA_PEM  (ou CURL_CA_BUNDLE=$CA_PEM)"
