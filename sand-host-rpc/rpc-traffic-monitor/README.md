# rpc-traffic-monitor

Proxy inverse HTTP/1.1 qui **journalise la requête ET la réponse** (corps
complets) du trafic RPC Connect/gRPC. **Les logs restent 100 % locaux** sous
`logs/` (gitignoré).

## Principe

`proxy.py` termine le HTTP du client en clair et ré-origine vers un `--upstream`
(http ou https). Selon la cible qu'on lui donne, il capture l'une des deux
pattes demandées :

| Patte | `--upstream` | Comment router le trafic dedans |
|---|---|---|
| **agent ↔ host** | `http://127.0.0.1:1340` (port sand-host) | pointer le client/probe sur `http://127.0.0.1:<port>` |
| **host ↔ backend** | `https://api2.cursor.sh` | pointer la base URL backend sur `http://127.0.0.1:<port>` (p.ex. `GROK_BOT_RPC_BASE_URL=http://127.0.0.1:<port>`) |

Le proxy voit les corps en clair parce que c'est lui qui ouvre la connexion TLS
vers l'upstream : pas de MITM de certificat, pas de CA à installer.

## Ce qui est journalisé

Une ligne **JSONL** par échange dans `logs/traffic-YYYY-MM-DD.jsonl` :

```json
{
  "ts": "2026-10-24T09:54:45.142+02:00",
  "leg": "host-backend",
  "direction": "client->api2.cursor.sh:443",
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

- **Corps requête + réponse** sont décodés : décompression `gzip`/`deflate`,
  et **dé-framing Connect/gRPC** (trames `[flag][len BE 32 bits][payload JSON]`)
  en liste de messages JSON.
- **Rédaction systématique** avant écriture :
  - en-têtes sensibles (`authorization`, `cookie`, `set-cookie`,
    `x-cursor-access-token`, …) → `[REDACTED]` ;
  - corps : JWT (`eyJ…`), `Bearer …`, tokens GitHub, clés privées PEM, e-mails,
    et les **valeurs littérales** des variables d'env de type
    `*TOKEN*/*SECRET*/*CREDENTIAL*/*KEY*/*PASSWORD*` du process courant →
    `[REDACTED-*]`.

Aucun jeton ni secret ni corps brut n'est committé : `logs/` est gitignoré.

## Usage

```bash
# Patte host ↔ backend (défaut) :
RTM_LEG=host-backend RTM_UPSTREAM=https://api2.cursor.sh ./start.sh
#   puis router la base URL backend sur http://127.0.0.1:8787

# Patte agent ↔ host :
RTM_LEG=agent-host RTM_UPSTREAM=http://127.0.0.1:1340 RTM_PORT=8788 ./start.sh

# Arrêt :
./stop.sh

# Preuve de capture req+resp (lance, appelle, affiche la dernière ligne) :
./selftest.sh
```

Variables : `RTM_PORT` (défaut `8787`), `RTM_UPSTREAM`
(défaut `https://api2.cursor.sh`), `RTM_LEG` (`agent-host` | `host-backend`).

## Limites

- Le proxy capture le trafic **qu'on route explicitement au travers** (via base
  URL / variable d'env du client). Il n'intercepte pas passivement une
  connexion déjà établie : pour la patte host↔backend il faut (re)pointer la
  base URL backend du host sur le proxy.
- HTTP/1.1 uniquement (pas de h2c en clair) ; connexion fermée après chaque
  échange (`Connection: close`) — suffisant pour Connect unaire et la plupart
  des flux ; le streaming long est journalisé à la fermeture.
- Alternative pour une interception transparente de la patte TLS host↔backend :
  `mitmproxy` avec une CA installée dans le trust store du host (plus intrusif,
  non requis ici).
