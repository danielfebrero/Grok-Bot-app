# rpc-traffic-monitor

Journalise la **requête ET la réponse** (corps complets) du trafic RPC
Connect/gRPC. Deux implémentations complémentaires :

1. **`proxy.py`** — proxy inverse HTTP/1.1 (sans CA, sans MITM de certificat) ;
2. **`mitmproxy/`** — addon mitmproxy (interception TLS transparente via CA), en
   **mode proxy explicite / opt-in** (cf. `mitmproxy/README.md`).

Les logs ne sont **pas committés** (`.gitignore` exclut `logs/`, `*.jsonl`) ; par
défaut ils vont dans le **Syncthing partagé EG** (voir plus bas).

## 1) proxy.py — proxy inverse en clair

`proxy.py` termine le HTTP du client en clair et ré-origine vers un `--upstream`
(http ou https). Selon la cible qu'on lui donne, il capture l'une des deux pattes :

| Patte | `--upstream` | Comment router le trafic dedans |
|---|---|---|
| **agent ↔ host** | `http://127.0.0.1:1340` (port sand-host) | pointer le client/probe sur `http://127.0.0.1:<port>` |
| **host ↔ backend** | `https://api2.cursor.sh` | pointer la base URL backend sur `http://127.0.0.1:<port>` (p.ex. `GROK_BOT_RPC_BASE_URL=http://127.0.0.1:<port>`) |

Le proxy voit les corps en clair parce que c'est lui qui ouvre la connexion TLS
vers l'upstream : pas de MITM de certificat, pas de CA à installer.

```bash
RTM_LEG=host-backend RTM_UPSTREAM=https://api2.cursor.sh ./start.sh   # défaut
RTM_LEG=agent-host  RTM_UPSTREAM=http://127.0.0.1:1340 RTM_PORT=8788 ./start.sh
./stop.sh
./selftest.sh      # preuve de capture req+resp (affiche la dernière ligne, rédactée)
```

Variables : `RTM_PORT` (défaut `8787`), `RTM_UPSTREAM`
(défaut `https://api2.cursor.sh`), `RTM_LEG` (`agent-host` | `host-backend`).

## 2) mitmproxy/ — interception TLS, proxy explicite

Pour capturer une connexion TLS existante sans ré-pointer la base URL, cf.
`mitmproxy/` : `setup-ca.sh` (CA + trust store), `start-mitm.sh` / `stop-mitm.sh`
(mitmdump + `addon_log.py` sur le port 8899), en **proxy explicite opt-in**
(`HTTPS_PROXY=http://127.0.0.1:8899`). **On ne détourne pas globalement l'egress
du host** (risque de flaps egress / cert-pinning) — détails et mode transparent
AVANCÉ non-activé dans `mitmproxy/README.md`.

## Ce qui est journalisé

Une ligne **JSONL** par échange dans `<RPC_TRAFFIC_LOG_DIR>/traffic-YYYY-MM-DD.jsonl` :

```json
{
  "ts": "2026-10-24T09:54:45.142+02:00",
  "leg": "host-backend",
  "direction": "host->backend",
  "method": "/aiserver.v1.DashboardService/GetTeams",
  "http_method": "POST",
  "status": "400",
  "duree_ms": 41.5,
  "req_headers": { "...": "rédigés" },
  "req_body":  "...",
  "resp_headers": { "...": "rédigés" },
  "resp_body": "..."
}
```

- **Corps requête + réponse** décodés : décompression `gzip`/`deflate`/`br`/`zstd`,
  **dé-framing Connect/gRPC** (trames `[flag][len BE 32 bits][payload]`), et
  **protobuf → JSON** quand le `content-type` est `application/proto` (Connect
  unaire), `application/grpc` ou `application/connect+proto`. Les noms de champs
  viennent du catalogue `../cartographie/data/rpc_catalog.json` (scalaires
  protobuf-es, messages imbriqués `ref#Type`). Un tag absent du message est
  `_unknown_<n>`. Les `int64` hors de l'intervalle sûr JSON sont des chaînes ;
  les `bytes` sont en base64. La rédaction s'applique au JSON produit.
- **Rédaction systématique** avant écriture : en-têtes sensibles
  (`authorization`, `cookie`, `set-cookie`, `x-cursor-access-token`, …) →
  `[REDACTED]` ; corps : JWT (`eyJ…`), `Bearer …`, tokens GitHub, clés privées PEM,
  e-mails, et les **valeurs littérales** des variables d'env
  `*TOKEN*/*SECRET*/*CREDENTIAL*/*KEY*/*PASSWORD*` → `[REDACTED-*]`.

## Où vont les logs (Syncthing partagé EG)

Par défaut : **`/workspace/shared-eg/box-vers-vps/sand-host-rpc-traffic/`** — le
dossier Syncthing partagé EG (écrit uniquement par la box, comme `github-traffic/`).
Surchargeable par l'env **`RPC_TRAFFIC_LOG_DIR`**.

**Rotation / cap** (contrainte shared-eg, 1 Go max indicatif) : un fichier par jour,
purge au-delà de `RPC_TRAFFIC_KEEP_DAYS` (défaut **14 j**) + cap total
`RPC_TRAFFIC_CAP_MB` (défaut **200 Mo**) — les plus vieux sont retirés. Aucun nom de
fichier ne contient « token »/« secret » (sinon `.stignore` EG bloquerait la synchro).

Les logs restent hors repo : `.gitignore` exclut `logs/`, `*.jsonl`.

## Limites

- `proxy.py` capture le trafic **qu'on route explicitement au travers** (base URL /
  variable d'env du client) ; il n'intercepte pas passivement une connexion déjà
  établie. HTTP/1.1 uniquement, `Connection: close` par échange.
- Le décodage protobuf couvre les échanges **unaires**. Un flux
  `application/connect+proto` / gRPC est dé-tramé si l'enveloppe est complète ;
  un stream qui ne se termine pas reste hors périmètre.
  Sans entrée de catalogue pour le path, le corps est quand même parcouru
  (tags `_unknown_<n>`, pas de noms inventés).
- **WebSocket** : `ws_frames.py` parse les trames (texte UTF-8/JSON, binaire
  hex ou proto si le path est au catalogue, contrôle résumé) et réutilise la
  même rédaction. Le relai agent↔host s'en sert après un `101`. La capture
  CDP loopback et le tcp/443 passent par `live-intercept` (tshark). RSV /
  permessage-deflate n'est pas décompressé. noVNC (6080/6081) et le tunnel
  egress sont hors périmètre — détails dans `live-intercept/README.md` §13.
- `mitmproxy/` nécessite la CA dans le trust store ; en **proxy explicite** il ne
  trace que les clients lancés avec `HTTPS_PROXY`. L'interception **transparente**
  (iptables) est documentée mais **non activée** (risque egress/pinning).
