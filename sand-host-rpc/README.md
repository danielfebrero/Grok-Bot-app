# sand-host-rpc

Outillage d'observation de la surface **RPC** du host Sand/Cursor (`api2.cursor.sh`,
protocole Connect) : jeton, cartographie exhaustive, suivi de dérive des
méthodes, et journalisation du trafic. Tout en français. Les secrets et les logs
restent **locaux** (voir `.gitignore`).

## Sous-dossiers

| Dossier | Rôle |
|---|---|
| [`token-helper/`](token-helper/) | Helper de jeton du sand-host (copie conforme du helper officiel `get-inference-token.py`). Renouvelle le jeton d'inférence depuis la credential du host et l'imprime seul — **ne logge jamais** la credential ni le jeton. |
| [`cartographie/`](cartographie/) | **Cartographie annotée exhaustive** des services/méthodes RPC : scan de tout `/opt/sand`, **21 services / 2317 méthodes** (tous packages : `aiserver.v1`, `agent.v1`, `origin.v1`, `orbit.v1`), fusionnée avec le balayage lecture (verdicts d'accès, surface IDOR). Sous-ensemble audité : 460 méthodes. |
| [`rpc-monitor/`](rpc-monitor/) | **Snapshots + diff + changelog** des déclarations RPC (sur le modèle de `box-update-monitor`/`mcp-schema-monitor`) : empreinte `I\|O\|kind` par méthode, détection ajout/suppression/modif, entrées horodatées commentées, routine locale (`cron.sh`). |
| [`rpc-traffic-monitor/`](rpc-traffic-monitor/) | Proxy inverse qui **journalise requête ET réponse** (corps complets dé-framés Connect + décompressés) des pattes **agent↔host** et **host↔backend**. Logs **locaux** uniquement, secrets/PII rédigés. |

## Démarrage rapide

```bash
# 1) Jeton d'inférence du host (jeton seul sur stdout, secret jamais loggé) :
cd token-helper && eval "$(./token-helper.sh --export)"

# 2) (Re)générer la cartographie exhaustive :
cd ../cartographie && python3 gen-cartographie.py

# 3) Baseline + détection de dérive des méthodes RPC :
cd ../rpc-monitor && ./snapshot.sh && ./check.sh

# 4) Journaliser le trafic (preuve req+resp) :
cd ../rpc-traffic-monitor && ./selftest.sh
```

## Routine locale (cron)

`rpc-monitor/cron.sh install` ajoute (sans écraser les crons existants) un job
toutes les ~2 h qui lance `check.sh` : en cas de changement de surface RPC il
crée un snapshot, un rapport de diff et une entrée horodatée dans
`rpc-monitor/CHANGELOG.md`. `cron.sh show` / `cron.sh uninstall` pour inspecter
ou retirer.

## Invariants de sécurité

- Aucune valeur de **credential** ni de **jeton** n'est committée ni loggée.
- Les **logs de trafic** (corps requêtes/réponses) restent locaux (`.gitignore`).
- `.gitignore` exclut `*.jwt`, `*.token`, `*.pem/.key`, `logs/`, `*.jsonl`.
- Balayage/audit **responsable** : lecture seule par défaut ; tests IDOR
  uniquement avec des identifiants fournis par l'opérateur, jamais devinés.
