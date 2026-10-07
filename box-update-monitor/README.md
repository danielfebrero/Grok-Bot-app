# box-update-monitor

Surveille ce qui change dans l'app Grok Bot / runtime agent de la box entre deux mises à jour :
- **Mise à jour de l'ordinateur** (Settings → Computer → Update) : nouveau pod, nouvelle image → `image_sha` change
  (`/etc/sand-box-image-sha`), et tout ce qui est sous `/usr`, `/opt`, `/exec-daemon` est réinstallé.
- **Hot-update du host bundle** (sans Update de l'ordinateur) : le sand-supervisor télécharge un nouveau
  `sand-host` depuis S3 et le remplace à chaud → `sand_host_version` change (`/opt/sand/sand-host/version`).
  Déjà observé : image `a6e2020` → host `f95dbfb` (cf. `/tmp/sand-supervisor.log`).
- **orbitd** (daemon root) s'auto-met à jour dans `/var/lib/sand-daemon-supervisor/orbitd/releases/<rev>`.

`/workspace` survit à l'Update de l'ordinateur, donc les snapshots aussi.

## Arborescence
```
perimeter.conf          périmètre surveillé (full = hash + copie texte, hash = hash seul, exclude)
snapshot.sh             crée snapshots/<YYYY-MM-DD_HH-MM>/
diff.sh                 compare 2 snapshots -> reports/<snapshot>.md (+ reports/<snapshot>-diffs/*.diff)
check-update.sh         détection auto (routine) -> snapshot + diff + reports/NOUVEAU-DIFF.flag
scan-personal-data.sh   scan données perso / secrets à lancer AVANT toute publication (exit 0 = propre)
VERSION-actuelle        versions/build courantes (réécrit à chaque snapshot)
lib/bum.py              moteur (python3 stdlib)
objects/                store dédupliqué par sha256 des fichiers texte (local seulement, gitignoré)
snapshots/<ts>/
  manifest.tsv          path  size  sha256  mtime  mode  kind  perimeter
  files/<chemin absolu> copies texte (hardlinks vers objects/ ; *.gz pour bundles > 2 Mo)
  meta/                 versions.txt, dpkg-packages.txt, python-dists.txt, npm-global.txt, env-var-names.txt
  fingerprint.txt       empreinte rapide utilisée par check-update.sh
  SNAPSHOT-INFO.txt     stats (nb fichiers, durée, label)
reports/                rapports Markdown + diffs complets
logs/                   monitor.log, check-update.log
```

## Commandes
```bash
# Snapshot manuel (ex. juste avant de cliquer sur Update)
bash /workspace/box-update-monitor/snapshot.sh --label "avant update"

# Diff manuel : 2 snapshots les plus récents
bash /workspace/box-update-monitor/diff.sh
# Diff entre deux snapshots précis / lister
bash /workspace/box-update-monitor/diff.sh 2026-10-07_16-42 2026-10-08_09-00
bash /workspace/box-update-monitor/diff.sh --list

# Cas "j'ai fait une update, compare" en une commande :
bash /workspace/box-update-monitor/check-update.sh --deep

# Routine auto
bash /workspace/box-update-monitor/check-update.sh
#   exit 0  + "NO_CHANGE last=<snap>"                         -> rien à faire
#   exit 10 + "CHANGED old=.. new=.. report=<chemin .md> ..." -> lire le rapport, notifier,
#             puis consommer le flag : rm /workspace/box-update-monitor/reports/NOUVEAU-DIFF.flag
#   exit 0  + "BASELINE snapshot=<snap>"                      -> premier snapshot créé
#   exit 3 verrou occupé ; exit 1 erreur
```
`check-update.sh` (rapide, ~2 s) compare : les versions (image_sha, box-scripts, sand-host, build exec-daemon,
canvas-sdk, orbitd, agent-store-fuse, Chrome, node, debian) + sha256 de fichiers clés (`host-main.cjs`,
`exec-daemon/index.js`, `sand-supervisor*.mjs`, `start-sand-box`, `box-doctor`, `sand-daemon-supervisor`…)
+ hash de contenu des arbres `reference/`, `managed-skills/`, `plugin-skills/`, `orbitd/releases/`.
`--deep` refait en plus un manifeste complet (~6 s) si l'empreinte rapide est identique ; snapshot jeté si rien n'a changé.

Fichiers minifiés (lignes > 1000 car.) : découpés sur `;` `{` `}` avant `diff -u` pour un diff lisible.

## Hors périmètre (bruit / données perso)
`/tmp` (logs, état runtime), `/home/box/sand-data` hors `managed-skills/` (transcripts, secrets, index, statsig,
et `plugins/` + `plugin-skills/` = connecteurs installés par l'utilisateur, donc propres au compte),
`/home/box/.cursor/projects` (sorties d'outils des agents), `chrome-profile`, `.cache`, `.npm`, `/workspace`,
`/opt/sand-managed/assignment.json` (affectation propre à la box), `orbitd/staging` + `state.json` (root-only),
`__pycache__`, `*.pyc`.

## Publication sur GitHub (données perso)
Le repo est **public**. Règles :
- Le périmètre ne contient que des fichiers app/système **génériques**, identiques d'une box à l'autre
  (runtime exec-daemon, sand-host, scripts de l'image, skills livrés par l'app, versions/paquets).
  Aucune donnée utilisateur : pas de transcripts, secrets, profil Chrome, connecteurs installés, `/workspace`.
- `meta/env-var-names.txt` ne contient que des **noms** de variables, jamais de valeurs.
- Avant chaque commit : `bash scan-personal-data.sh` doit répondre `PROPRE`. Le scanner lit ses motifs
  perso dans `.pii-denylist` (**local, gitignoré**, donc jamais publié), cherche les motifs de tokens/clés
  (ghp_, github_pat_, AKIA, JWT, clés privées, bearer…) et cherche littéralement les valeurs des variables
  d'env sensibles et des fichiers secrets de la box (sans jamais les afficher).
- `.gitignore` exclut `objects/` (doublon), `logs/`, le flag, `.pii-denylist` et, par sécurité, tout chemin
  utilisateur (`sand-data` sauf managed-skills, chrome-profile, .cursor/projects, .ssh, workspace, *secret*, *.key…).
- Les adresses e-mail/IP présentes dans les copies sont celles des **paquets open source** (auteurs npm,
  exemples de doc, liste des suffixes publics) ou des **templates officekit** (noms fictifs) — pas des données perso.
