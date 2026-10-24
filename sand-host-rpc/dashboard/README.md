# Dashboard temps réel — trafic RPC capturé

Visualisation **en temps réel** des échanges RPC loggés par le moniteur de trafic
(`../rpc-traffic-monitor/`). 100 % bibliothèque standard Python côté serveur,
HTML/CSS/JS **vanilla** côté client (graphiques dessinés en `<canvas>` natif,
aucune dépendance npm/CDN).

## Ce que ça lit

Les JSONL rédactés écrits par le moniteur dans :

```
$RPC_TRAFFIC_LOG_DIR/traffic-YYYY-MM-DD.jsonl      # 1 ligne JSON par échange
```

Défaut : `/workspace/shared-eg/box-vers-vps/sand-host-rpc-traffic/`.

> **Lecture seule.** Le dashboard n'écrit JAMAIS dans les logs. Les corps/headers
> sont déjà rédactés par le moniteur (tokens/email/secrets masqués) ; le serveur
> ne fait que tronquer les corps pour l'affichage.

## Lancer

```bash
./start-dashboard.sh                 # http://127.0.0.1:8808
./stop-dashboard.sh
```

Puis ouvrir http://127.0.0.1:8808 .

### Variables d'environnement

| Var                   | Défaut                                                        | Rôle                              |
|-----------------------|---------------------------------------------------------------|-----------------------------------|
| `DASH_PORT`           | `8808`                                                        | port d'écoute                     |
| `DASH_HOST`           | `127.0.0.1`                                                   | hôte d'écoute (local uniquement)  |
| `RPC_TRAFFIC_LOG_DIR` | `/workspace/shared-eg/box-vers-vps/sand-host-rpc-traffic`     | dossier des JSONL                 |
| `DASH_MAX_RECORDS`    | `20000`                                                      | taille du buffer mémoire (tail)   |

## API

- `GET /` · `GET /app.js` · `GET /style.css` — l'UI.
- `GET /api/stats?days=N&last=K` — agrégats JSON sur les N derniers jours
  (défaut 1, max 14), avec les K derniers échanges (défaut 50, max 500) :
  - `total`, `by_leg`, `top_methods`, `status_dist`, `hosts`
  - `latency` : `p50` / `p95` / `max`
  - `error_count`, `error_rate`
  - `rate_now`, `timeline` (échanges/minute, 30 derniers buckets)
  - `last` : derniers échanges avec corps **tronqués** (≤ 2 Ko)

Le serveur fait un **tail incrémental** (il ne relit que les octets ajoutés
depuis le dernier passage, suit la rotation de fichier par inode/offset).

## UI

Rafraîchit toutes les **2 s** (polling `/api/stats`, le plus simple/robuste en
stdlib) :

- cartes : total, échanges/min, latence p95, taux d'erreur ;
- débit par minute en `<canvas>` (aire + ligne) ;
- donut des pattes (agent→host / host→backend) en `<canvas>` ;
- barres top méthodes RPC et distribution des status ;
- tableau live des derniers échanges ; clic = modale req/resp (tronqués).

Thème sombre, autonome, aucune ressource externe.
