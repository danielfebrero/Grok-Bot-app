#!/usr/bin/env bash
# RETIRÉ. La patte host↔backend n'utilise plus mitmproxy ni une CA maison.
# Le chemin actuel est le keylog TLS (format NSS / SSLKEYLOGFILE) :
#   hb-setup-keylog.sh   (un restart, --tls-keylog seulement)
#   hb-on.sh / hb-off.sh (tcpdump + tshark, sans restart)
#   hb-revert.sh
echo "RETIRÉ: hb-setup-ca.sh (mitm + NODE_EXTRA_CA_CERTS) n'arme plus rien." >&2
echo "Utiliser: LI_ARM_CONFIRM=yes-i-understand-brick-risk sudo -E bash hb-setup-keylog.sh" >&2
exit 1
