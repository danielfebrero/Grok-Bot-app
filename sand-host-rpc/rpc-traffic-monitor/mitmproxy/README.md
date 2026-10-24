# mitmproxy — interception TLS (req/resp) du trafic RPC, en **proxy explicite**

Interception transparente du TLS pour journaliser **requête ET réponse** des RPC
Connect/gRPC (host↔backend) et agent↔host, avec les corps en clair. Complément de
`../proxy.py` (proxy inverse) : ici on garde le TLS de bout en bout grâce à une CA
mitmproxy installée dans le trust store.

## ⚠️ Décision de sécurité (appliquée)

On **NE redirige PAS** globalement/transparentement tout l'egress du host sand-host
à travers mitmproxy. L'interception transparente (iptables REDIRECT + `--mode
transparent`) casse facilement l'egress live de l'agent et bute sur le
cert-pinning — on a déjà eu des **flaps d'egress**.

**Par défaut : proxy EXPLICITE, opt-in.** Seuls les process qu'on route
volontairement passent par mitmproxy :

```bash
HTTPS_PROXY=http://127.0.0.1:8899 HTTP_PROXY=http://127.0.0.1:8899 <commande>
```

Le reste de l'egress du host n'est pas touché.

## Composants

| Fichier | Rôle |
|---|---|
| `addon_log.py` | addon mitmproxy (hooks `request`/`response`) ; journalise chaque échange en JSONL, corps dé-framés (Connect) + décompressés (gzip/deflate/br/zstd), **secrets/PII rédigés** (réutilise `../proxy.py`) |
| `setup-ca.sh` | génère la CA mitmproxy (`~/.mitmproxy/mitmproxy-ca-cert.pem`) et l'installe dans le trust store système (`/usr/local/share/ca-certificates` + `update-ca-certificates`). Idempotent |
| `start-mitm.sh` | lance `mitmdump -s addon_log.py` sur le port `MITM_PORT` (défaut **8899**), en arrière-plan, PID dans `mitm.pid` |
| `stop-mitm.sh` | arrête le mitmdump référencé par `mitm.pid` |

## Mise en route

```bash
# venv dédié (mitmproxy n'est pas packagé système sur la box)
python3 -m venv /workspace/.venvs/mitmproxy
/workspace/.venvs/mitmproxy/bin/pip install mitmproxy

./setup-ca.sh        # CA + trust store (sudo pour update-ca-certificates)
./start-mitm.sh      # écoute 127.0.0.1:8899

# router un client volontairement :
HTTPS_PROXY=http://127.0.0.1:8899 \
  curl --cacert ~/.mitmproxy/mitmproxy-ca-cert.pem \
  -s -o /dev/null https://api2.cursor.sh/
./stop-mitm.sh
```

## Où vont les logs

Dans le **Syncthing partagé EG**, `box-vers-vps/` (écrit uniquement par la box) :

```
/workspace/shared-eg/box-vers-vps/sand-host-rpc-traffic/traffic-YYYY-MM-DD.jsonl
```

Surchargeable par `RPC_TRAFFIC_LOG_DIR`. **Rotation** : un fichier par jour, purge
au-delà de `RPC_TRAFFIC_KEEP_DAYS` (défaut 14 j) + cap total `RPC_TRAFFIC_CAP_MB`
(défaut 200 Mo). **Aucun secret** n'est écrit (rédaction systématique) et aucun nom
de fichier ne contient « token »/« secret » (sinon le `.stignore` EG le bloquerait).

## Mode transparent (AVANCÉ — NON activé)

Pour mémoire seulement, **non activé** ici :

```bash
# NE PAS faire par défaut (risque de casser l'egress live / cert-pinning) :
sudo sysctl -w net.ipv4.ip_forward=1
sudo iptables -t nat -A OUTPUT -p tcp --dport 443 -j REDIRECT --to-port 8899
mitmdump --mode transparent -s addon_log.py
```

Cette route détourne tout le 443 du host et peut provoquer des flaps d'egress et des
échecs de pinning. À n'envisager que ponctuellement, sous supervision, puis à défaire
(`iptables -t nat -F OUTPUT`).
