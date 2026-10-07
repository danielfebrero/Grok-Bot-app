# Grok-Bot-app

Filesystem dump of the non-stock box tree used by Grok Bot.

See **[non-stock/](non-stock/)** for the preserved path architecture (`exec-daemon/`, `home/`, `usr/`, `workspace/`) plus manifests and notes.

Personal media and secrets were excluded from this dump.

## box-update-monitor

**[box-update-monitor/](box-update-monitor/)** surveille ce qui change dans l'app Grok Bot sur la box d'une mise à jour à l'autre :
une « Mise à jour de l'ordinateur » (nouvelle image, `image_sha`), une mise à jour à chaud du host
(`sand-host` version) ou une mise à jour d'`orbitd`.

- `snapshot.sh` : instantané horodaté (manifest chemin/taille/sha256/mtime + copies des fichiers texte + versions et paquets)
- `diff.sh` : rapport Markdown entre deux instantanés (fichiers ajoutés / supprimés / modifiés + diff ligne à ligne)
- `check-update.sh` : détection automatique (appelée par une routine) qui crée l'instantané et le rapport en cas de changement
- `scan-personal-data.sh` : scan données perso / secrets obligatoire avant chaque publication

Ne contient que des fichiers app/système génériques : aucune donnée utilisateur (voir le README du dossier).
