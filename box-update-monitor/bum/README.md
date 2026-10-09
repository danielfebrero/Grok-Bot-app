# bum/ — changelog de l'ordinateur cloud (box)

`bum` = **b**ox-**u**pdate-**m**onitor. Ce dossier traduit les diffs bruts de `../reports/` en un
changelog lisible : **snapshot par snapshot, qu'est-ce qui a changé et de quel type de changement il s'agit**.

## Fichiers
- `CHANGELOG.md` — une section par snapshot qui a produit un vrai diff (les passages `NO_CHANGE` ne sont pas listés).
  Entrées les plus récentes **en haut**.

## Format d'une entrée
```
## YYYY-MM-DD HH:MM (Paris) — snapshot `<précédent>` → `<nouveau>`
- Commit GitHub : lien relatif `../../../../commit/<sha>` (pas de slug de repo en dur : le scan données perso le bloquerait)
- Rapport brut : ../reports/<snapshot>.md
- Versions touchées : sand-host, image, agent-store-fuse, skills gérées, dépendances…
### Bugfix        corrections / durcissements visibles dans un diff lisible
### Change        comportement modifié, retrait, renommage, montée de version sans détail
### Features      capacité nouvelle, étayée par le diff
### Perf          optimisation visible dans le code
### Autre / infra rebuilds, binaires, pins de version
```
Règles de classement :
- On ne classe que ce que le rapport ou un diff **lisible** montre. Rien n'est inventé.
- Un gros bundle minifié reconstruit (`host-main.cjs`, `sand-eval-runner.cjs`) ou un binaire sans diff texte
  va sous **Autre / infra** avec la mention « rebuild opaque, classement hypothétique ». Les indices
  repérés en cherchant des mots-clés dans le diff minifié sont signalés comme indices, pas comme faits.
- Une rubrique vide est notée « — » (gardée pour que le format reste stable).
- Les heures sont celles du nom de snapshot (heure de Paris, CEST).

## Mise à jour
La veille horaire (`check-update.sh` → `publish.sh`) ajoute une entrée en haut de `CHANGELOG.md`
à chaque nouveau snapshot avec diff, puis le dossier est publié avec le snapshot (publish.sh recopie tout
`box-update-monitor/`, `bum/` compris, sur le repo GitHub de publication (`.publish.conf`) → `box-update-monitor/bum/`).
Le lien de commit d'une entrée est celui du push du snapshot correspondant. Il peut être ajouté à l'exécution suivante
s'il n'était pas encore connu au moment de l'écriture.
