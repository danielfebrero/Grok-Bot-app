# mcp-schema-monitor

Suit chaque jour les **schémas des tools MCP** exposés à l'agent Grok Bot (nom, description, `inputSchema`)
et le **statut** de chaque serveur (ready / needsAuth, nombre de tools). Calqué sur `box-update-monitor`
(même verrou, même scan PII bloquant, même publication).

Seuls les schémas et descriptions sont publiés — **jamais** de résultats d'appels d'outils, de tokens,
de valeurs de compte ni d'identifiants personnels.

## Pourquoi un dossier `inbox/`
Les schémas ne sont lisibles **que par un appel d'outil de l'agent** (`GetDynamicTools`), pas par un script.
Le flux est donc en deux temps :
1. l'agent appelle `GetDynamicTools` et écrit le JSON brut dans `inbox/` (local, gitignoré) ;
2. `check.sh` normalise, compare et écrit les rapports ; `publish.sh` pousse.

## Arborescence
```
inbox/_catalog.json      sortie de GetDynamicTools pattern="." : {"namespaces": {ns: {"status", "tools": [noms]}}, "absent": {...}}
inbox/<namespace>.json   sortie brute de GetDynamicTools namespace=<ns> (schémas complets)
lib/msm.py               moteur (python3 stdlib)
check.sh                 normalise + compare (verrou .lock)
publish.sh               publication GitHub, bloquée si le scan n'est pas propre
scan-personal-data.sh    scan données perso / secrets (copie de box-update-monitor)
snapshots/<ts>/index.json                 statut, nb de tools, sha256 de chaque schéma, niveau de capture
snapshots/<ts>/<namespace>/<tool>.json    {tool, description, inputSchema}, clés triées
reports/<ts>.md          rapport (ajoutés / retirés / modifiés + diffs JSON unifiés)
CHANGELOG.md             une ligne par snapshot
```
Niveau de capture par namespace : `full` (schéma complet → diffs de schéma) ou `names`
(seulement la liste des tools issue du catalogue → ajouts/retraits détectés, pas les modifs de schéma).
Passer un namespace de `names` à `full` apparaît dans le rapport comme « niveau de capture ».

## Commandes
```bash
bash /workspace/mcp-schema-monitor/check.sh
#   exit 0  + "BASELINE snapshot=<ts> ..."      premier snapshot
#   exit 0  + "NO_CHANGE last=<ts>"             rien n'a changé (aucun snapshot créé)
#   exit 10 + "CHANGED old=.. new=.. report=<md> added=.. removed=.. modified=.."
#   exit 3 verrou occupé ; exit 1 erreur (ex. inbox/_catalog.json absent)
bash /workspace/mcp-schema-monitor/publish.sh [--dry-run]
#   0 publié / NOTHING_TO_PUBLISH / dry-run propre · 2 bloqué par le scan · 3 verrou · 1 erreur
```

## Données perso
- `inbox/`, `logs/`, `.pii-denylist`, `.publish.conf` sont gitignorés et exclus de la copie.
- `msm.py` ne garde que `tool`, `description`, `inputSchema`.
- Le namespace first-party `cursor` n'est inclus que si son fichier d'entrée passe le scan.
- `publish.sh` scanne exactement les fichiers à pousser ; si non propre : rien n'est poussé,
  `reports/PUBLISH-BLOCKED.flag` (chemins + n° de règle, jamais la valeur) est écrit, exit 2.
