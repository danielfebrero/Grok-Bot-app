#!/usr/bin/env bash
# RETIRÉ. mitmproxy transparent + --allow-hosts n'est plus le chemin host↔backend.
# Capture passive : start-capture-host-backend.sh (lancé par hb-on.sh / hb-setup-keylog.sh).
echo "RETIRÉ: start-mitm-host-backend.sh n'intercepte plus le TLS." >&2
echo "Utiliser: hb-setup-keylog.sh puis hb-on.sh (tcpdump + keylog NSS + tshark)." >&2
exit 1
