#!/usr/bin/env bash
# Installe (idempotent) une entrée crontab locale qui lance check.sh toutes les 2h.
# N'ÉCRASE PAS les crons existants : ajoute la ligne seulement si absente.
# Usage : ./cron.sh install | uninstall | show
set -euo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
TAG="# sand-host-rpc-monitor"
LINE="0 */2 * * * $DIR/check.sh >> $DIR/logs/cron.log 2>&1 $TAG"
case "${1:-show}" in
  install)
    if ! command -v crontab >/dev/null 2>&1; then
      echo "crontab indisponible sur cette box."
      echo "Installe la ligne suivante dans ton planificateur local :"
      echo "  $LINE"
      exit 0
    fi
    cur="$(crontab -l 2>/dev/null || true)"
    if printf '%s\n' "$cur" | grep -qF "$TAG"; then
      echo "déjà installé :"; printf '%s\n' "$cur" | grep -F "$TAG"; exit 0
    fi
    { printf '%s\n' "$cur"; echo "$LINE"; } | crontab -
    echo "installé : $LINE"
    ;;
  uninstall)
    command -v crontab >/dev/null 2>&1 || { echo "crontab indisponible"; exit 0; }
    crontab -l 2>/dev/null | grep -vF "$TAG" | crontab - || true
    echo "désinstallé (lignes $TAG retirées)"
    ;;
  show|*)
    echo "ligne cron gérée :"; echo "  $LINE"
    if command -v crontab >/dev/null 2>&1; then
      echo "--- crontab actuel (lignes du moniteur) ---"
      crontab -l 2>/dev/null | grep -F "$TAG" || echo "(aucune installée)"
    else
      echo "crontab indisponible sur cette box."
    fi
    ;;
esac
