# token-helper

Helper de jeton du **sand-host**. C'est une copie conforme du helper officiel
`/workspace/sand-helpers/get-inference-token.py` : on réutilise l'existant, on
n'en réécrit pas un nouveau.

## Rôle

Renouveler le **jeton d'inférence** du sand-host :

1. localise le process `sand-host` (port `SAND_HOST_PORT`, défaut `1340`) ;
2. lit `SAND_INFERENCE_RENEWAL_CREDENTIAL` dans l'environnement de ce process
   (`/proc/<pid>/environ`) ;
3. POST `https://api2.cursor.sh/sand-box/inference-credential` avec cette
   credential (via le proxy d'egress du sandbox si présent) ;
4. imprime **UNIQUEMENT** le jeton (`accessToken`) sur stdout.

## Sécurité — invariants

- **Ne logge jamais** la credential `SAND_INFERENCE_RENEWAL_CREDENTIAL`.
- **Ne logge jamais** le jeton ailleurs que là où l'appelant le demande
  explicitement (`--out`, stdout, `--export`).
- **Aucune valeur de secret n'est committée** : le script ne lit les secrets
  qu'au runtime depuis l'env du host ; rien de sensible n'est écrit dans le
  dépôt. Les jetons produits (`*.jwt`, `*.token`, `/tmp/...`) sont hors du
  dépôt et couverts par le `.gitignore` du dossier `sand-host-rpc/`.

## Usage

```bash
# Exporter les variables pour les autres outils RPC :
eval "$(./token-helper.sh --export)"
#   -> export GROK_BOT_RPC_AUTH="Bearer …"
#      export GROK_BOT_RPC_BASE_URL="https://api2.cursor.sh"

# Récupérer le jeton sur stdout :
TOKEN=$(./token-helper.sh)

# Écrire le jeton dans un fichier 0600 (hors dépôt) :
./token-helper.sh --out /tmp/sand-inference.token

# JSON {accessToken, expiresAtMs} :
./token-helper.sh --json
```

`token-helper.sh` est un simple wrapper qui `exec` le script officiel
`get-inference-token.py` (toutes les options lui sont passées telles quelles).

## Options principales (`get-inference-token.py`)

| Option | Effet |
|---|---|
| *(aucune)* | jeton brut sur stdout |
| `--export` | lignes `export GROK_BOT_RPC_AUTH=… / _BASE_URL=…` pour `eval` |
| `--json` | `{"accessToken": …, "expiresAtMs": …}` |
| `--out PATH` | écrit le jeton dans `PATH` en `0600` |
| `--host-port N` | port du sand-host (défaut `SAND_HOST_PORT` ou `1340`) |
| `--pid N` | PID du sand-host (saute la détection `ss`) |
| `--backend URL` | backend (défaut `GROK_BOT_RPC_BASE_URL` ou `api2.cursor.sh`) |
| `--no-proxy` | ne pas utiliser le proxy d'egress |

## Note sur les deux jetons

- **Jeton d'inférence sand-host** (ce helper) : jeton machine renouvelé via la
  credential du host, utilisé pour les appels RPC host→backend d'inférence.
- **Jeton `sand_mobile`** (hors de ce helper) : le JWT porteur utilisé pour
  l'audit/balayage des méthodes de la cartographie est chargé séparément
  (jamais committé, cf. `.gitignore`).
