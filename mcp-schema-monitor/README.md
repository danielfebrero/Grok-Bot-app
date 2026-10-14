# mcp-schema-monitor

Suit chaque jour les **schémas des tools MCP** exposés à l'agent Grok Bot (nom, `title`, description, `inputSchema`, `outputSchema`, `annotations`)
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
snapshots/<ts>/<namespace>/<tool>.json    {tool, title?, description, inputSchema, outputSchema, annotations?}, clés triées
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
- `msm.py` ne garde que `tool`, `title`, `description`, `inputSchema`, `outputSchema`, `annotations`.

## outputSchema
- Conservé tel quel quand le descripteur l'a ; sinon écrit explicitement `"outputSchema": null`.
- Il fait partie du sha256 du tool : un changement d'outputSchema est un « Modifié » avec diff.
- `index.json` donne par namespace `output_schema_count` (null au niveau `names`, inconnu) ;
  le rapport l'affiche en colonne `outputSchema` (n/total).
- Le namespace first-party `cursor` n'est inclus que si son fichier d'entrée passe le scan.
- `publish.sh` scanne exactement les fichiers à pousser ; si non propre : rien n'est poussé,
  `reports/PUBLISH-BLOCKED.flag` (chemins + n° de règle, jamais la valeur) est écrit, exit 2.

## Fixtures et schémas de sortie inférés
Aucun serveur ne déclare d'`outputSchema` aujourd'hui. On l'infère donc à partir de vraies réponses.
- `fixtures/<ns>.json` : une entrée par tool, `category` :
  - `a` lecture seule (`status` executed / planned / failed + raison) ;
  - `b` révocable : création suivie de son `pair` (delete/undo) dans le même run. **Non exécuté** tant que
    Dani n'a pas validé explicitement ces écritures (`planned-awaiting-approval`) ;
  - `c` ignoré + `reason` : envoi de message/email/DM/SMS, publication, crédits/argent, merge/suppression,
    création d'inbox/domaine/repo, validation de campagne, données de tiers, `get_users_me`.
- Les valeurs propres à Dani (owner/repo…) sont des placeholders `<OWN_REPO_OWNER>` ; les vraies valeurs
  sont dans `fixtures/values.local.json` (gitignoré, exclu de la copie).
- `lib/infer.py <ns> <tool> <réponse_brute>` : infère un JSON Schema (types, clés, `required`, tableaux,
  types multiples/null) dans `inbox/_inferred.json` (local), puis **supprime la réponse brute**.
  Aucune valeur n'est conservée ni publiée.
- `check.sh` écrit `snapshots/<ts>/_inferred/<ns>/<tool>.json`
  (`outputSchemaInferred` : `source: "inferred"`, `inferredAt`, `responseKind`, `schema`), séparé de
  l'`outputSchema` déclaré. Le hash d'index ne porte que sur le schéma (pas l'horodatage). Un changement
  apparaît dans le rapport (section « outputSchemaInferred », ligne « Inférés modifiés »).
  Sans `inbox/_inferred.json`, les inférences du snapshot précédent sont reprises telles quelles.
- Cadence recommandée pour rejouer les fixtures : **hebdomadaire** (les sorties bougent peu, les lectures X
  consomment des crédits et Exa/Tavily limitent le débit). La capture des descripteurs reste quotidienne.

## Procédure de la routine (quotidienne)
1. `GetDynamicTools` pattern `.` → `inbox/_catalog.json` (statut + noms par namespace).
2. `GetDynamicTools` namespace=<ns> pour chaque namespace ready → `inbox/<ns>.json` (réponse telle quelle).
3. (Hebdo) Lire `user-X/get_usage_credits` ; puis pour chaque fixture `a` executed/planned : `CallDynamicTool`
   avec `write_to_file`, arguments de `fixtures/<ns>.json` (placeholders remplacés depuis
   `fixtures/values.local.json`), puis `python3 lib/infer.py <ns> <tool> <fichier>` (supprime le brut).
   Jamais de fixture `b` ou `c` sans accord explicite.
4. `check.sh` : NO_CHANGE → s'arrêter sans bruit ; BASELINE/CHANGED → `publish.sh`.
5. PUBLISHED → résumer le rapport à Dani, supprimer `reports/NOUVEAU-DIFF.flag` ; exit 2 → signaler les chemins
   de `reports/PUBLISH-BLOCKED.flag` sans contourner le scan.
6. Vider `inbox/` (sauf si l'on veut garder `_inferred.json` pour la semaine).
