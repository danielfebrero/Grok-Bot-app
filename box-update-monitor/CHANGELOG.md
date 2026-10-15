# Changelog de l'ordinateur cloud (box Grok Bot)

Snapshot par snapshot, ce qui a changé dans le runtime de la box (sand-host, image, agent-store-fuse, skills gérées…),
classé en **Bugfix / Change / Features / Perf / Autre-infra**. Entrées les plus récentes en haut.
Seuls les snapshots avec un vrai diff sont listés. Heures de Paris (CEST).

Base de référence : snapshot `2026-10-07_17-05` (« avant update (Dani) ») — image `a6e2020`, sand-host `f95dbfb`,
agent-store-fuse `e44f687`, Chrome 154.0.8037.57, Node 22.14.0 / 20.19.2, Python 3.13.5, Debian 13.7.

---

## 2026-10-15 18:45 (Paris) — snapshot `2026-10-15_17-46` → `2026-10-15_18-45`

Versions touchées : agent-store-fuse `f624c0a` → `a5f7c61` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version`) ; image `4098fb0`, sand-host `701b86e`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump agent-store-fuse (rebuild opaque / oscillation du pin, comme les snapshots précédents). Seul le fichier de version change.

Rapport : [reports/2026-10-15_18-45.md](reports/2026-10-15_18-45.md).


## 2026-10-15 17:46 (Paris) — snapshot `2026-10-15_12-51` → `2026-10-15_17-46`

Versions touchées : **sand-host** `4098fb0` → `701b86e` (hot-update, image inchangée `4098fb0`) ; agent-store-fuse `0c3b083` → `f624c0a` ; exec-daemon, orbitd, Chrome 154.0.8037.97, Node, Python inchangés. Fichiers : +1 / −0 / ~13.

- **Bugfix** : —
- **Change** : `box-chrome` v3 exporte `SAND_DESKTOP_UNIT`, ajoute `--disable-blink-features=AutomationControlled` et un flag launcher versionné ; `stop-window` gagne un mode `--park` (garde le profil Chrome) et un arrêt propre du navigateur du profil avant le sweep ; `start-window` source `sand-desktop-unit.sh` et tue l’unité avant de relancer X.
- **Features** : nouveau script `sand-desktop-unit.sh` (kill des process tagués `SAND_DESKTOP_UNIT` + Chrome du profil) ; `sand-window-router` marque les displays utilisés via `/tmp/sand-window-used.d` (touch périodique + marqueurs `.inflight`).
- **Perf** : `sand-ua-governor` attache récursivement iframes/workers out-of-process (`Target.setAutoAttach` + filtre) pour propager le traitement UA au-delà des enfants directs.
- **Autre/infra** : rebuilds `host-main.cjs` / `sand-eval-runner.cjs` ; `agent-store-worker`, `box-store-vacuum-worker` et `search-index-worker` reçoivent la lib d’écriture atomique no-follow + exclusions sand-data host-owned ; bump agent-store-fuse (pin S3).

Rapport : [reports/2026-10-15_17-46.md](reports/2026-10-15_17-46.md).

## 2026-10-15 12:51 (Paris) — snapshot `2026-10-15_10-54-2` → `2026-10-15_12-51`

Versions touchées : **image** `726deab` → `4098fb0` (box scripts idem) ; exec-daemon build `2026-10-12T14:49` → `2026-10-14T14:48` ; agent-store-fuse `7516de1` → `0c3b083` ; sand-host déjà `4098fb0` (alignement image↔host après Update Computer). Chrome 154.0.8037.97, Node, Python, orbitd inchangés. Fichiers : +4 / −0 / ~42.

- **Bugfix** : restauration des cookies Chrome au cold launch (`box-chrome` + `sand-cookie-persist.mjs --restore-port`) pour éviter qu’une page charge avant le seed et écrase la session.
- **Change** : extension WebAuthn proxy re-pinnée (`oifaalpenddpcpgklkncbmdaocjdnlgc` → `alhdnckjfkmaljmljdhjfpecddnfjide`, crx + policy + native-messaging + update.xml). Paquets Debian : curl/libcurl → u7 ; `librsvg2` security bump ; retrait dpkg de `grok-bot` 0.70.2, tesseract/leptonica/libsecret, et des libs FIDO2/`openssh-client` ajoutées à 10:54. Skills gérées : slides/slides-executor (édition live via Slides MCP), purchases (vérifier commande après échec), agent-email (forme `name+tag@`), code-changes (watch = conversation, pas session).
- **Features** : nouveau module natif `sand-card-text-recognition` (OCR on-device pour fill photo de carte) dans `/opt/sand/deps` + manifeste runtime ; hook hibernate/box-store-flush dans `start-sand-box` (`SAND_BOX_STORE_HIBERNATE_HOOK`, fifo `/run/sand-box-store-flush`) ; `cursor-proclist` expose `cursor_set_compact_window_toolbar` (macOS) ; chunk exec-daemon `81.index.js`.
- **Perf** : —
- **Autre/infra** : rebuilds binaires (`cursorsandbox`, `tools/origin`, `sand-egress-tunnel`, `uv`/`uvx`, `sand-daemon-supervisor`, canvas-runtime, polished-renderer) ; bump agent-store-fuse (pin S3).

Rapport : [reports/2026-10-15_12-51.md](reports/2026-10-15_12-51.md).


## 2026-10-15 10:54 (Paris) — snapshot `2026-10-15_07-46` → `2026-10-15_10-54`

Versions touchées : agent-store-fuse `0c3b083` → `7516de1` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version`) ; **nouveau support WebAuthn / FIDO2 côté box**. image `726deab`, sand-host `4098fb0`, exec-daemon, orbitd, Chrome 154.0.8037.97, Node, Python inchangés. Fichiers : +0 / −0 / ~2.

- **Bugfix** : —
- **Change** : —
- **Features** : support WebAuthn ajouté au Chrome de la box. La policy managée `/etc/opt/chrome/policies/managed/sand-webauthn.json` passe de `{}` à un `ExtensionSettings` qui **force-install** l'extension `oifaalpenddpcpgklkncbmdaocjdnlgc` (proxy WebAuthn), via `update_url` `file:///usr/local/share/sand-webauthn-proxy-update.xml`. Nouveaux paquets Debian : `libfido2-1` (1.15.0-1+b1), `libcbor0.10` (0.10.2-2), `openssh-client` (1:10.0p1-7+deb13u4) — brique FIDO2 / clé de sécurité.
- **Perf** : —
- **Autre/infra** : bump agent-store-fuse (rebuild opaque / oscillation du pin, comme les snapshots précédents).

Rapport : [reports/2026-10-15_10-54.md](reports/2026-10-15_10-54.md).


## 2026-10-15 07:46 (Paris) — snapshot `2026-10-15_06-52` → `2026-10-15_07-46`

Versions touchées : agent-store-fuse `51fd48b` → `0c3b083` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image `726deab`, sand-host `4098fb0`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque / oscillation du pin, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` garde le **même hash** (`47779169…`) que dans les snapshots `19:48`, `20:51`, `21:45`, `22:54`, `00:50`, `01:49`, `02:52` et `06:52`.

Rapport : [reports/2026-10-15_07-46.md](reports/2026-10-15_07-46.md).


## 2026-10-15 06:52 (Paris) — snapshot `2026-10-15_02-52` → `2026-10-15_06-52`

Versions touchées : agent-store-fuse `b811a50` → `51fd48b` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image `726deab`, sand-host `4098fb0`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque / oscillation du pin, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` garde le **même hash** (`47779169…`) que dans les snapshots `19:48`, `20:51`, `21:45`, `22:54`, `00:50`, `01:49` et `02:52`.

Rapport : [reports/2026-10-15_06-52.md](reports/2026-10-15_06-52.md).


## 2026-10-15 02:52 (Paris) — snapshot `2026-10-15_01-49` → `2026-10-15_02-52`

Versions touchées : agent-store-fuse `791571b` → `b811a50` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image `726deab`, sand-host `4098fb0`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque / oscillation du pin, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` garde le **même hash** (`47779169…`) que dans les snapshots `19:48`, `20:51`, `21:45`, `22:54`, `00:50` et `01:49`.

Rapport : [reports/2026-10-15_02-52.md](reports/2026-10-15_02-52.md).


## 2026-10-15 01:49 (Paris) — snapshot `2026-10-15_00-50` → `2026-10-15_01-49`

Versions touchées : agent-store-fuse `0c0f4a8` → `791571b` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image `726deab`, sand-host `4098fb0`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque / oscillation du pin, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` garde le **même hash** (`47779169…`) que dans les snapshots `19:48`, `20:51`, `21:45`, `22:54` et `00:50`.

Rapport : [reports/2026-10-15_01-49.md](reports/2026-10-15_01-49.md).


## 2026-10-15 00:50 (Paris) — snapshot `2026-10-14_22-54` → `2026-10-15_00-50`

Versions touchées : agent-store-fuse `85429d0` → `0c0f4a8` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image `726deab`, sand-host `4098fb0`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque / oscillation du pin, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` garde le **même hash** (`47779169…`) que dans les snapshots `19:48`, `20:51`, `21:45` et `22:54`.

Rapport : [reports/2026-10-15_00-50.md](reports/2026-10-15_00-50.md).

## 2026-10-14 22:54 (Paris) — snapshot `2026-10-14_21-45` → `2026-10-14_22-54`

Versions touchées : agent-store-fuse `f5cd108` → `85429d0` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque / oscillation du pin, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** (`47779169…`) que dans les snapshots `19:48`, `20:51` et `21:45`.

Rapport : [reports/2026-10-14_22-54.md](reports/2026-10-14_22-54.md).

## 2026-10-14 21:45 (Paris) — snapshot `2026-10-14_20-51` → `2026-10-14_21-45`

Versions touchées : agent-store-fuse `403533d` → `f5cd108` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers l’archive précédente) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : rollback du pointeur agent-store-fuse vers `f5cd108` (rebuild opaque / oscillation du pin, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** (`47779169…`) que dans les snapshots `19:48` et `20:51`.

Rapport : [reports/2026-10-14_21-45.md](reports/2026-10-14_21-45.md).

## 2026-10-14 20:51 (Paris) — snapshot `2026-10-14_19-48` → `2026-10-14_20-51`

Versions touchées : agent-store-fuse `f5cd108` → `403533d` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** dans ce snapshot.

Rapport : [reports/2026-10-14_20-51.md](reports/2026-10-14_20-51.md).

## 2026-10-14 19:48 (Paris) — snapshot `2026-10-14_17-55` → `2026-10-14_19-48`

**Hot-update du host bundle** (sans Update de l'ordinateur) : sand-host `726deab` → `4098fb0` (`supervisor_status.lastCommandId = upgrade-4098fb0`). Image `726deab`, box-scripts, exec-daemon, orbitd, agent-store-fuse `f5cd108`, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~10 (9 sous `/opt/sand`, 1 `managed-skills`).

- **Bugfix** : **restauration des cookies de session Chrome à froid** — `box-chrome` pousse désormais le « cookie seed » sauvegardé (`/home/box/sand-data/chrome-cookie-seed.json`) via `sand-cookie-persist.mjs --restore-port` **avant** de donner une URL à Chrome. Sans ça, une page chargée en premier recevait un cookie de session neuf qui écrasait le login sauvegardé. Important pour la persistance des sessions navigateur des agents (X, etc.).
- **Change** : protobuf host — `ComputerUseArgs` gagne `mark_monitor_busy` (+ `ComputerUseSuccess.monitor_busy_marked`, `ComputerUseError.monitor_busy_marked`) ; `PrManagementArgs` gagne deux actions `list_prs` et `get_pr` ; `SendToAgentArgs` gagne `tracked_task_keys` (liste). Signale de nouvelles capacités côté outils (computer-use, gestion PR, suivi de tâches inter-agents).
- **Features** : — (à confirmer dans les tools exposés : `list_prs`/`get_pr` et `tracked_task_keys` sont des champs de schéma, pas forcément déjà actifs).
- **Perf** : `search-index-worker.cjs` allégé (76 Ko → 42 Ko, −723 lignes). `host-main.cjs` 28,6 → 29,5 Mo, `sand-eval-runner.cjs` 21,8 → 23,1 Mo, `officekit office.mjs` 700 → 827 Ko.
- **Autre/infra** : `diff-worker.js` +121/−11, `agent-store-worker.cjs` +1157/−11. dpkg : `grok-bot 0.70.2` de nouveau énuméré (était absent de l'énumération du snapshot précédent, pas une réinstallation), `libsecret-1-0` / `libsecret-common` 0.21.7 présents. `managed-skills/cache.json` rafraîchi (`fetchedAt`).

Rapport : [reports/2026-10-14_19-48.md](reports/2026-10-14_19-48.md).

## 2026-10-14 17:55 (Paris) — snapshot `2026-10-14_15-51` → `2026-10-14_17-55`

Versions touchées : aucune (image `726deab`, sand-host, exec-daemon, orbitd, agent-store-fuse `f5cd108`, Chrome, Node, Python inchangés). Fichiers : +0 / −0 / ~0. Seul changement : la méta `dpkg-packages.txt`.

- **Bugfix** : —
- **Change** : —
- **Features** : **nouveau composant** — chaîne OCR **Tesseract 5.5.0** installée sur la box : `tesseract-ocr`, `libtesseract5`, `libleptonica6`, `libarchive13t64`, packs langues `tesseract-ocr-eng` / `tesseract-ocr-fra` / `tesseract-ocr-osd` (eng, fra, osd disponibles). Binaire `/usr/bin/tesseract` opérationnel.
- **Perf** : —
- **Autre/infra** : installé par un **agent** via `apt-get install -y -qq tesseract-ocr tesseract-ocr-fra` le 2026-10-14 à 17:02 (Paris) — ce n'est **pas** une mise à jour de l'ordinateur (image inchangée), mais une installation de paquet sur la box en cours d'exécution. Capacité OCR désormais dispo pour toute l'équipe.

Rapport : [reports/2026-10-14_17-55.md](reports/2026-10-14_17-55.md).

## 2026-10-14 15:51 (Paris) — snapshot `2026-10-14_14-47` → `2026-10-14_15-51`

Versions touchées : aucune (image, sand-host, exec-daemon, orbitd, agent-store-fuse, Chrome, Node, Python inchangés). Fichiers : +0 / −0 / ~5 (tous sous `managed-skills`).

- **Bugfix** : —
- **Change** : skills gérées — `agent-email` : envoi possible depuis une adresse +tag, routine `name+tag@` limitée à ce tag ; `DraftExternalMessage` ne mentionne plus l’envoi auto via Allow rules. `routines` : listener email documente le filtre par tag. `send-on-behalf` : les « saved send rules » des instructions autorisent aussi l’envoi direct hors carte. `export-bot-template` : wording Public-only (pas d’offre Team) + conservation du scope existant à la mise à jour.
- **Features** : —
- **Perf** : —
- **Autre/infra** : `managed-skills/cache.json` rafraîchi (`fetchedAt` + miroir des corps de skills). Nouvelle variable d’env nommée `SAND_AGENT_STORE_RPC_SERVICE`. Meta `dpkg-packages.txt` : ~23 paquets absents de l’énumération (dont `openssh-client`, `nmap`, `bind9-*`, `grok-bot` 0.70.2) — à traiter comme signal méta, pas comme preuve d’uninstall silencieux sans vérif locale.

Rapport : [reports/2026-10-14_15-51.md](reports/2026-10-14_15-51.md) — [commit](../../../commit/6602c63a933985ad67640251bde11c37aab46253).

## 2026-10-14 14:47 (Paris) — snapshot `2026-10-14_11-47` → `2026-10-14_14-47`

Versions touchées : agent-store-fuse `9172edf` → `f5cd108` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** dans ce snapshot.

Rapport : [reports/2026-10-14_14-47.md](reports/2026-10-14_14-47.md) — [commit](../../../commit/1e5b6a032a96d6d608ad27f0d0c81adf0b1186f8).

## 2026-10-14 11:47 (Paris) — snapshot `2026-10-14_10-52` → `2026-10-14_11-47`

Versions touchées : agent-store-fuse `f9ff468` → `9172edf` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** dans ce snapshot.

Rapport : [reports/2026-10-14_11-47.md](reports/2026-10-14_11-47.md) — [commit](../../../commit/345c197e151178106ca8949e0db8256ba34dce38).

## 2026-10-14 10:52 (Paris) — snapshot `2026-10-14_07-48` → `2026-10-14_10-52`

Versions touchées : agent-store-fuse `12180d7` → `f9ff468` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** dans ce snapshot.

Rapport : [reports/2026-10-14_10-52.md](reports/2026-10-14_10-52.md) — [commit](../../../commit/4430c3d89c2c98b5440d8bd92a13567a2ac2ea8e).

## 2026-10-14 07:48 (Paris) — snapshot `2026-10-14_06-47` → `2026-10-14_07-48`

Versions touchées : agent-store-fuse `3247f7b` → `12180d7` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** dans ce snapshot.

Rapport : [reports/2026-10-14_07-48.md](reports/2026-10-14_07-48.md) — [commit](../../../commit/3c097af0c958b0f2c8f8f62e593846b4abd733a9).

## 2026-10-14 06:47 (Paris) — snapshot `2026-10-14_04-53` → `2026-10-14_06-47`

Versions touchées : agent-store-fuse `45ce8ea` → `3247f7b` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** dans ce snapshot.

Rapport : [reports/2026-10-14_06-47.md](reports/2026-10-14_06-47.md) — [commit](../../../commit/1712c89d47c3d7dec52d377813fac13ad8a864b7).

## 2026-10-14 04:53 (Paris) — snapshot `2026-10-13_21-48` → `2026-10-14_04-53`

Versions touchées : agent-store-fuse `50c9bf7` → `45ce8ea` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** dans ce snapshot.

Rapport : [reports/2026-10-14_04-53.md](reports/2026-10-14_04-53.md) — [commit](../../../commit/4608dce653469e237017fb1d613d011a9d8ab228).

## 2026-10-13 21:48 (Paris) — snapshot `2026-10-13_20-56` → `2026-10-13_21-48`

Versions touchées : agent-store-fuse `21827f8` → `50c9bf7` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** dans ce snapshot.

Rapport : [reports/2026-10-13_21-48.md](reports/2026-10-13_21-48.md) — [commit](../../../commit/851040c57a6467a8ba4b926a2fd21ce6a4a2f9a7).

## 2026-10-13 20:56 (Paris) — snapshot `2026-10-13_19-59` → `2026-10-13_20-56`

Versions touchées : agent-store-fuse `89678ad` → `21827f8` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** dans ce snapshot.

Rapport : [reports/2026-10-13_20-56.md](reports/2026-10-13_20-56.md) — [commit](../../../commit/b95d8d05fd442d15f34dc9182be3cfe6acffb61e).

## 2026-10-13 19:59 (Paris) — snapshot `2026-10-13_19-06` → `2026-10-13_19-59`

Versions touchées : agent-store-fuse `688705f` → `89678ad` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** dans ce snapshot.

Rapport : [reports/2026-10-13_19-59.md](reports/2026-10-13_19-59.md) — [commit](../../../commit/8ee12bbdd6226d52c711b08246ff4ea0c1cbbdce).

## 2026-10-13 19:06 (Paris) — snapshot `2026-10-13_17-59` → `2026-10-13_19-06`

Versions touchées : agent-store-fuse `6efa6c7` → `688705f` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique). Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** dans ce snapshot.

Rapport : [reports/2026-10-13_19-06.md](reports/2026-10-13_19-06.md) — [commit](../../../commit/542bc07a05d1081ebfeab507e9af6dd1c0aa338c).

## 2026-10-13 17:59 (Paris) — snapshot `2026-10-13_17-00` → `2026-10-13_17-59`

Versions touchées : agent-store-fuse `5cfc9a4` → `6efa6c7` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive ; binaire `cursor-agent-store-fuse` reconstruit) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~2.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique). Côté paquets dpkg : apparition de `grok-bot 0.70.2` + dépendances desktop (`libayatana-appindicator3-1`, `libayatana-ido3-0.4-0`, `libayatana-indicator3-7`, `libsecret-1-0`, `libsecret-common`). Ce n'est **pas** une Update de l'ordinateur poussée par Cursor : un agent de l'équipe a installé en local le `.deb` de l'app desktop (téléchargé dans `/workspace/grok-bot-install`, `dpkg` à 17:33), bruit propre à cette box.

Rapport : [reports/2026-10-13_17-59.md](reports/2026-10-13_17-59.md).

## 2026-10-13 08:49 (Paris) — snapshot `2026-10-13_07-49` → `2026-10-13_08-49`

Versions touchées : agent-store-fuse `5cadd0d` → `5cfc9a4` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-13_08-49.md](reports/2026-10-13_08-49.md) — [commit](../../../commit/896b2e4d81ff4e1adb74c8db7cac8428d9613617).

## 2026-10-13 07:49 (Paris) — snapshot `2026-10-13_05-45` → `2026-10-13_07-49`

Versions touchées : agent-store-fuse `511f501` → `5cadd0d` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive ; binaire `cursor-agent-store-fuse` reconstruit) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~2.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-13_07-49.md](reports/2026-10-13_07-49.md) — [commit](../../../commit/7a1b4c97fa3ff73f4c7b4bd911fabd1a6956d5c1).

## 2026-10-13 05:45 (Paris) — snapshot `2026-10-13_03-54` → `2026-10-13_05-45`

Versions touchées : image / box-scripts `7345aa4` → `726deab` (aligne l’image sur le sand-host déjà en `726deab` ; `lastCommandId` `upgrade-726deab` → `None`) ; exec-daemon rebuild (build 2026-10-11T15:02Z → 2026-10-12T14:49Z) ; canvas-sdk nouvelle version ; orbitd release `d2ed475` retirée (reste `28b0833`) ; agent-store-fuse, sand-host, Chrome, Node, Python inchangés. Fichiers : +24 / −8 / ~39.

- **Bugfix** : —
- **Change** : skill gérée `routines` — une routine ne peut pas remplir un login/carte/secret sauvegardé (secure form filling exige l’utilisateur dans le chat) ; le résumé post-save doit prévenir qu’elle pingera l’utilisateur à cette étape.
- **Features** : canvas SDK Grok — nouveaux composants typés (`Avatar`/`AvatarStack`, `WeekView`/`MonthView`, `Scatter`, `Heatmap`/`CalendarHeatmap`, `Diagram`, `Matrix2x2`/`Venn`, listes, quote, TOC, equation, status, divider, link-cards, pros-cons, properties…) ; `controls.d.ts` et `reference.md` nettement élargis.
- **Perf** : —
- **Autre/infra** : Update Computer image `726deab` ; soupsieve 3.0.1→3.0.3 ; extension sand-webauthn-proxy mise à jour (id/crx/update.xml + native messaging) ; `sand-daemon-supervisor` reconstruit ; bundles exec-daemon / canvas-runtime / pnpm / `cursorsandbox` (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-13_05-45.md](reports/2026-10-13_05-45.md) — [commit](../../../commit/12e25388c9dc125dbf23e35f0fd7dc1247075605).

## 2026-10-13 03:54 (Paris) — snapshot `2026-10-13_01-51` → `2026-10-13_03-54`

Versions touchées : agent-store-fuse `c8ef2b3` → `511f501` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` revenu sur l'archive précédente, retour arrière du build de 01:51) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : retour arrière agent-store-fuse vers `511f501` (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-13_03-54.md](reports/2026-10-13_03-54.md) — [commit](../../../commit/81397883b37c0bb8d401501cc88853357a6fb653).

## 2026-10-13 01:51 (Paris) — snapshot `2026-10-13_00-46` → `2026-10-13_01-51`

Versions touchées : agent-store-fuse `511f501` → `c8ef2b3` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-13_01-51.md](reports/2026-10-13_01-51.md) — [commit](../../../commit/647c213f33b33a2f58442daed8feed450d8e192c).

## 2026-10-13 00:46 (Paris) — snapshot `2026-10-12_22-52` → `2026-10-13_00-46`

Versions touchées : agent-store-fuse `50e3127` → `511f501` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-13_00-46.md](reports/2026-10-13_00-46.md) — [commit](../../../commit/60331a15df58700fde09ee3ff58ff7b220829324).

## 2026-10-12 22:52 (Paris) — snapshot `2026-10-12_20-53` → `2026-10-12_22-52`

Versions touchées : agent-store-fuse `f351382` → `50e3127` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-12_22-52.md](reports/2026-10-12_22-52.md) — [commit](../../../commit/23f33f0ed20dade8b33b11a6f13f6132aec58eb6).

## 2026-10-12 20:53 (Paris) — snapshot `2026-10-12_19-05` → `2026-10-12_20-53`

Versions touchées : agent-store-fuse `c44c38f` → `f351382` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-12_20-53.md](reports/2026-10-12_20-53.md) — [commit](../../../commit/8399dfd68ed01bfc884e8a1e5df025df45caf13c).

## 2026-10-12 19:05 (Paris) — snapshot `2026-10-12_18-11` → `2026-10-12_19-05`

Versions touchées : agent-store-fuse `ae03d5d` → `c44c38f` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-12_19-05.md](reports/2026-10-12_19-05.md) — [commit](../../../commit/a524e37510ec48c702b0b21bcffae915decaace5).

## 2026-10-12 18:11 (Paris) — snapshot `2026-10-12_17-06` → `2026-10-12_18-11`

Versions touchées : sand-host `7345aa4` → `726deab` (superviseur `lastCommandId` = `upgrade-726deab`) ; image, box-scripts, exec-daemon, agent-store-fuse, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~8 (`host-main.cjs`, `sand-eval-runner.cjs`, `agent-store-worker.cjs`, `box-store-vacuum-worker.cjs`, `csnaps`, `version`, skill gérée `agent-email` + `cache.json`).

- **Bugfix** : —
- **Change** : skill gérée `agent-email` — nouvelle section « Sending from your inbox » : le mail de la boîte native part via `DraftExternalMessage` (carte, ou envoi direct si une règle Allow le couvre) ou via `SendEmail` quand le propriétaire a coupé le brouillon (avec bcc) ; ajout : la carte n'a pas encore de désinscription en un clic, donc pas de prospection / newsletter / envoi de masse depuis la boîte avec la carte. `box-store-vacuum-worker.cjs` : raisons `cold_create` et `first_connect` retirées de la liste. URLs du portail Cursor désormais paramétrables (`cursor-web-base-urls`, `portal-base-url.ts`) au lieu d'hôtes codés en dur dans `parseCloudCanvasWebLink`. Modules retirés du bundle host : `frozen-prompt-tools.ts`, `sand-subagent-management-tools.ts`, `messages-mac/phone.js`.
- **Features** : outils « PR code tour » (`get-pr-code-tour`, `update-pr-code-tour`, protos `GetPrCodeTour*` / `UpdatePrCodeTour*`, `pr-code-tour-view-context`) dans host-main, agent-store-worker et eval-runner ; brouillons de la boîte agent (`agent-inbox-draft-route.ts`, `agent-inbox-drafts.ts`, `draft-route-verification.ts`, `grok-bot-email-client.ts`) ; nouveaux modules `cloud-agents/` (tool, confirm-policy, environment, images, publish-app, canvas-prompt) ; `tool-permissions.ts`, `background-result-note.ts` ; `ShellStreamStart` gagne un champ `command`. Rôle exact déduit des noms de modules dans un bundle minifié.
- **Perf** : module `partial-tool-call-throttle.js` ajouté (nom seul, classement hypothétique).
- **Autre/infra** : patch-hash `@connectrpc/connect-node` changé ; renommages de variables dans `sand-eval-runner.cjs` ; `cache.json` des skills gérées rafraîchi ; binaire `codebase-telemetry/csnaps` reconstruit (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-12_18-11.md](reports/2026-10-12_18-11.md) — [commit](../../../commit/9287c1c0e2d7640aa7418d276a1ef436e8401b4b).

## 2026-10-12 17:06 (Paris) — snapshot `2026-10-12_15-01` → `2026-10-12_17-06`

Versions touchées : agent-store-fuse `db2b1ee` → `ae03d5d` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-12_17-06.md](reports/2026-10-12_17-06.md) — [commit](../../../commit/dbcbbed43844033203839e87d6725aa883b8c7e7).

## 2026-10-12 15:01 (Paris) — snapshot `2026-10-12_03-44` → `2026-10-12_15-01`

Versions touchées : orbitd — nouvelle release `28b0833` ajoutée sous `/var/lib/sand-daemon-supervisor/orbitd/releases/` à côté de `d2ed475` (révision embarquée `8828749` inchangée) ; image, sand-host, exec-daemon, agent-store-fuse, Chrome, Node, Python inchangés. Fichiers : +2 / −0 / ~0.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : release orbitd `28b0833` déposée par le superviseur (manifest.json schema 1, cible x86_64, payload zstd ~6,9 Mo, exécutable ~21 Mo) ; binaire sans diff lisible (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-12_15-01.md](reports/2026-10-12_15-01.md) — [commit](../../../commit/19cc802784f31df9fd06606e8d03b440b363381b).

## 2026-10-12 03:44 (Paris) — snapshot `2026-10-12_02-51` → `2026-10-12_03-44`

Versions touchées : agent-store-fuse `55863b4` → `db2b1ee` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-12_03-44.md](reports/2026-10-12_03-44.md) — [commit](../../../commit/3f3d21f145f93d69099795cef3189269e9098539).

## 2026-10-12 02:51 (Paris) — snapshot `2026-10-11_23-51` → `2026-10-12_02-51`

Versions touchées : agent-store-fuse `0d66b5d` → `55863b4` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` vers la nouvelle archive) ; image, sand-host, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : nouveau build agent-store-fuse (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-12_02-51.md](reports/2026-10-12_02-51.md) — commit [`2092c7a`](../../../commit/2092c7acd0126892d4cb3abe7b7855d710432282)

---

## 2026-10-11 23:51 (Paris) — snapshot `2026-10-11_20-47` → `2026-10-11_23-51`

Versions touchées : image / box-scripts `91a9d91` → `7345aa4` ; exec-daemon rebuild (build 2026-10-08 → 2026-10-11T15:02Z) ; canvas-sdk nouvelle version ; agent-store-fuse `e1fd3f3` → `0d66b5d` (sand-host, Chrome, Node, Python inchangés). Fichiers : +41 / −38 / ~144.

- **Bugfix** : —
- **Change** : `sand-web-bot-auth.mjs` — suppression des marqueurs/portée optionnelle `xhrFetch` / `iframes` (WEB_BOT_AUTH_DEFAULT_SCOPE retiré) ; doc canvas `reference.md` : notes de règle d’usage après `—` par composant, props `LineChart` étendues.
- **Features** : —
- **Perf** : —
- **Autre/infra** : paquets Python numpy 2.5.3→2.5.4, onnxruntime 1.30.0→1.31.0, soupsieve 2.10→3.0.1 ; Debian libgs10 (Ghostscript) deb13u2→deb13u3 ; extension sand-webauthn-proxy mise à jour ; bundles exec-daemon / canvas-runtime / pnpm (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-11_23-51.md](reports/2026-10-11_23-51.md) — commit [`d66a337`](../../../commit/d66a33768aef1de367be4a4f3793a860f5825927)

## 2026-10-11 20:47 (Paris) — snapshot `2026-10-11_17-53` → `2026-10-11_20-47`

Versions touchées : agent-store-fuse `f4c8738` → `e1fd3f3` (image, sand-host, exec-daemon, Chrome inchangés). Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : `cursor_agent_store_fuse_version` pointe vers le nouveau tarball agent-store-fuse `e1fd3f3` (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-11_20-47.md](reports/2026-10-11_20-47.md) — commit [`8187157`](../../../commit/81871572a12f8a884e544f1c0e0bedd4f45de09b)

## 2026-10-11 17:53 (Paris) — snapshot `2026-10-11_03-46` → `2026-10-11_17-53`

Versions touchées : sand-host `f1727e3` → `7345aa4` (image, box-scripts, agent-store-fuse, exec-daemon, Chrome inchangés). Fichiers : +0 / −0 / ~7.

- **Bugfix** : index content-search : `parseAgentId` avant purge des agents hors disque (au lieu d’un `typeof string` brut) ; shell bash : détection « pipe commande illisible » (fd 5 fermé par `$BASH_ENV` / `.zshenv`, exit 126) avec bascule vers passage en arguments.
- **Change** : skill gérée `purchases` — section Credit cards (accepter / rappeler une carte avant checkout, ne jamais prétendre qu’elle est déjà sauvée hors Secure Form) ; `GetTeamMembersResponse.former_personal_owner_user_id` ; proto SCM `SCMMergeQueueStatus` ; check GTM MCP via `gtmApiProductionUrl` plutôt qu’hostname en dur.
- **Features** : génération vidéo — champ `last_frame_path` / `last_frame` sur `GenerateVideoArgs` et `RunGenerateVideoRequest` ; dossier artefacts cloud `publish-preview/` (`CLOUD_AGENT_PUBLISH_PREVIEW_ARTIFACTS_DIR`).
- **Perf** : —
- **Autre/infra** : `managed-skills/cache.json` (fetchedAt + miroir description purchases). Reste des bundles `host-main.cjs` / `sand-eval-runner.cjs` : rebuild opaque, classement hypothétique.

Rapport : [reports/2026-10-11_17-53.md](reports/2026-10-11_17-53.md) · [commit 35f55ea](../../../commit/35f55ea41a20f7876e50581aca87ca26d7dc1e8b)

## 2026-10-11 03:46 (Paris) — snapshot `2026-10-11_01-52` → `2026-10-11_03-46`

Versions touchées : agent_store_fuse.commit `df6db54` → `f4c8738` (image, sand-host, exec-daemon inchangés).

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` mis à jour vers l'archive agent-store-fuse `f4c8738` (rebuild opaque, classement hypothétique).

Rapport : [reports/2026-10-11_03-46.md](reports/2026-10-11_03-46.md) · [commit d255132](../../../commit/d255132acc7f87082b855a7345af5406e65295cc)

## 2026-10-11 01:52 (Paris) — snapshot `2026-10-11_00-50` → `2026-10-11_01-52`

Versions touchées : agent-store-fuse `f37f6d4` → `df6db54`. Image, box-scripts, sand-host, exec-daemon, Chrome inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** — —
- **Change** — —
- **Features** — —
- **Perf** — —
- **Autre/infra** — Bump agent-store-fuse (marqueur URL uniquement) : `f37f6d46833a…` → `df6db54b754f…`. Rebuild opaque, classement hypothétique.

Rapport : [reports/2026-10-11_01-52.md](reports/2026-10-11_01-52.md) — [commit 911f5a6](../../../commit/911f5a6aefb0d9869bc522ae16e606b1d6e7db0c)

## 2026-10-11 00:50 (Paris) — snapshot `2026-10-10_22-51` → `2026-10-11_00-50`

Versions touchées : agent-store-fuse `a009ba9` → `f37f6d4`. Image, box-scripts, sand-host, exec-daemon, Chrome inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** — —
- **Change** — —
- **Features** — —
- **Perf** — —
- **Autre/infra** — Bump agent-store-fuse (marqueur URL uniquement) : `a009ba9f779c…` → `f37f6d46833a…`. Rebuild opaque, classement hypothétique.

Rapport : [reports/2026-10-11_00-50.md](reports/2026-10-11_00-50.md) — [commit 08ee082](../../../commit/08ee0828ea738241b99bc1afea1c224e66f2a9a5)

## 2026-10-10 22:51 (Paris) — snapshot `2026-10-10_17-49` → `2026-10-10_22-51`

Versions touchées : sand-host `fd09280` → `f1727e3` ; agent-store-fuse `7356478` → `a009ba9` ; app 0.70.0-pre.18 → 0.71.0-pre.1. Image, box-scripts, exec-daemon, Chrome inchangés. Fichiers : +0 / −0 / ~6.

- **Bugfix** — —
- **Change** — Billing / merge équipe xAI : nouveaux outcomes `ADD_PAYMENT_METHOD` ; champ `processingBlockCode` sur le statut de merge ; validateurs bornés `brandedFailureCode` / `brandedVersion` ; `agent-store-worker` expose `LIVE_BYTES_BY_TYPE_KEYS` (types live + `pending`).
- **Features** — Mode CLI local (`CURSOR_AGENT_CLI_LOCAL_MODE` + header `local-cli-mode`) ; couche `shell-exec` runtime-env (filtre env Electron, scrub sockets SSH/DBus/Wayland, ignore-mapping, pression event-loop, suppression de sortie, politiques sandbox/réseau, dumps d'état shell bash/zsh/powershell).
- **Perf** — Tracker de pression event-loop et suppression de sortie côté shell-exec (lisibles dans `sand-eval-runner.cjs`).
- **Autre/infra** — `managed-skills/cache.json` : seul `fetchedAt` change. Bump agent-store-fuse (marqueur URL uniquement). Reste des bundles `host-main.cjs` / `sand-eval-runner.cjs` : rebuild opaque, classement hypothétique.

Rapport : [reports/2026-10-10_22-51.md](reports/2026-10-10_22-51.md) — [commit b93da89](../../../commit/b93da895fb76abf70bfd8a2e7e0972d576c7e110)

## 2026-10-10 17:49 (Paris) — snapshot `2026-10-10_09-52` → `2026-10-10_17-49`

Versions touchées : sand-host `964107d` → `fd09280` ; app 0.70.0-pre.16 → 0.70.0-pre.18. Image, box-scripts, exec-daemon, Chrome inchangés. Fichiers : +0 / −0 / ~4.

- **Bugfix** — —
- **Change** — Les serveurs MCP en statut `error` remontent désormais leur message d'erreur (`errorMessage`) dans l'état MCP ; flag client `email` retiré de la liste des fonctionnalités.
- **Features** — Nouveaux outils `ReadMeetingLines` et `RespondInMeeting` ; canal `meeting-request:` (agents de requêtes en réunion, consignes de wake associées) ; message dédié pour les MCP sur réseau Tailscale privé (suggère Tailscale Funnel).
- **Perf** — —
- **Autre/infra** — `managed-skills/cache.json` : seul `fetchedAt` change. Reste des bundles `host-main.cjs` / `sand-eval-runner.cjs` : rebuild opaque, classement hypothétique.

Rapport : [reports/2026-10-10_17-49.md](reports/2026-10-10_17-49.md) — [commit d12d18a](../../../commit/d12d18a45b64f2ff9a94657cd593b51c6f8e096c)

## 2026-10-10 09:52 (Paris) — snapshot `2026-10-10_07-51` → `2026-10-10_09-52`
- Commit GitHub : [2a7bb15](../../../commit/2a7bb15cf20e01c907ce715390edd25f6bc03b99)
- Rapport brut : [reports/2026-10-10_09-52.md](reports/2026-10-10_09-52.md) (diffs complets : `reports/2026-10-10_09-52-diffs/`)
- Versions touchées : **sand-host** `91efc0d` → `964107d`. Inchangés : image/box-scripts `91a9d91`, agent-store-fuse `7356478`, exec-daemon, Chrome 154.0.8037.97, Node, Python, Debian.
- Fichiers : +1 / −0 / ~9.

**Bugfix** — —
**Change** — Skills gérées : `in-chat-forms` précise que le Secure Form sert aux codes à usage unique *qu'on ne peut pas lire depuis une boîte mail* ; retouches de texte dans `sign-in` et `no-connector-fallback`.
**Features** — Nouvelle skill gérée `agent-email` (inbox native Grok Bot : claim, envoi, lecture, adresses à tag +).
**Perf** — —
**Autre/infra** — `host-main.cjs`, `sand-eval-runner.cjs`, `agent-isolation/agent-store-worker.cjs` reconstruits (rebuild opaque, classement hypothétique) ; `cache.json` des skills rafraîchi.

---

## 2026-10-10 07:51 (Paris) — snapshot `2026-10-10_05-45` → `2026-10-10_07-51`
- Commit GitHub : [a5c963e](../../../commit/a5c963e45360ac83e6c89f7870a3fe0acf184acb)
- Rapport brut : [reports/2026-10-10_07-51.md](reports/2026-10-10_07-51.md) (diffs complets : `reports/2026-10-10_07-51-diffs/`)
- Versions touchées : **agent-store-fuse** `36088c9` → `7356478`. Inchangés : image/box-scripts `91a9d91`, sand-host `91efc0d`, exec-daemon, Chrome 154.0.8037.97, Node, Python, Debian.
- Diff : 0 ajouté, 0 supprimé, 1 modifié (`/usr/local/bin/cursor_agent_store_fuse_version`).

### Bugfix
- —
### Change
- Nouveau commit agent-store-fuse (`7356478`) ; seul le fichier de version change, contenu du changement non visible.
### Features
- —
### Perf
- —
### Autre/infra
- Bump du marqueur de version agent-store-fuse (classement hypothétique, pas de diff de code lisible).

---

## 2026-10-10 05:45 (Paris) — snapshot `2026-10-10_04-52` → `2026-10-10_05-45`
- Commit GitHub : [486fbe6](../../../commit/486fbe611c9a607ed4dd8646ada49d181e1e24b5)
- Rapport brut : [reports/2026-10-10_05-45.md](reports/2026-10-10_05-45.md) (diffs complets : `reports/2026-10-10_05-45-diffs/`)
- Versions touchées : **agent-store-fuse** `a5520b0` → `36088c9`. Inchangés : image/box-scripts `91a9d91`, sand-host `91efc0d`, exec-daemon, Chrome 154.0.8037.97, Node, Python, Debian.
- Diff : 0 ajouté, 0 supprimé, 1 modifié (`/usr/local/bin/cursor_agent_store_fuse_version`).

### Bugfix
- —
### Change
- —
### Features
- —
### Perf
- —
### Autre/infra
- Nouveau commit agent-store-fuse (`36088c9`), seul le fichier de version change ; pas de diff lisible du contenu (rebuild opaque, classement hypothétique).

---

## 2026-10-10 04:52 (Paris) — snapshot `2026-10-10_03-51` → `2026-10-10_04-52`
- Commit GitHub : [47cdea6](../../../commit/47cdea652b4dd5bea36c5a933e7fe3260925897c)
- Rapport brut : [reports/2026-10-10_04-52.md](reports/2026-10-10_04-52.md) (diffs complets : `reports/2026-10-10_04-52-diffs/`)
- Versions touchées : **agent-store-fuse** `8f7e930` → `a5520b0`. Inchangés : image/box-scripts `91a9d91`, sand-host `91efc0d`, exec-daemon, Chrome 154.0.8037.97, Node, Python, Debian.
- Diff : 0 ajouté, 0 supprimé, 1 modifié (`/usr/local/bin/cursor_agent_store_fuse_version`).

### Bugfix
- —
### Change
- —
### Features
- —
### Perf
- —
### Autre/infra
- Bump du commit agent-store-fuse (seul le fichier de version change ; rebuild opaque, classement hypothétique).

---

## 2026-10-10 03:51 (Paris) — snapshot `2026-10-10_02-46` → `2026-10-10_03-51`
- Commit GitHub : [b038926](../../../commit/b0389260577079ebd0615dd304346d75a0676613)
- Rapport brut : [reports/2026-10-10_03-51.md](reports/2026-10-10_03-51.md) (diffs complets : `reports/2026-10-10_03-51-diffs/`)
- Versions touchées : **agent-store-fuse** `b41b395` → `8f7e930`. Inchangés : image/box-scripts `91a9d91`, sand-host `91efc0d`, exec-daemon, Chrome 154.0.8037.97, Node, Python, Debian.
- Diff : 0 ajouté, 0 supprimé, 1 modifié (`/usr/local/bin/cursor_agent_store_fuse_version`).

### Bugfix
- —
### Change
- —
### Features
- —
### Perf
- —
### Autre/infra
- Nouveau commit agent-store-fuse (seul le fichier de version change ; pas de diff lisible du binaire, rebuild opaque, classement hypothétique).

---

## 2026-10-10 02:46 (Paris) — snapshot `2026-10-10_01-47` → `2026-10-10_02-46`
- Commit GitHub : [fff585c](../../../commit/fff585cff6c8521b51518b1a4fa7710dfbf31762)
- Rapport brut : [reports/2026-10-10_02-46.md](reports/2026-10-10_02-46.md) (diffs complets : `reports/2026-10-10_02-46-diffs/`)
- Versions touchées : **agent-store-fuse** `5cf8ba8` → `b41b395`. Inchangés : image/box-scripts `91a9d91`, sand-host `91efc0d`, exec-daemon, Chrome 154.0.8037.97, Node, Python, Debian.
- Diff : 0 ajouté, 0 supprimé, 1 modifié (`/usr/local/bin/cursor_agent_store_fuse_version`).

### Bugfix
- —
### Change
- —
### Features
- —
### Perf
- —
### Autre/infra
- Nouveau commit agent-store-fuse (seul le fichier de version change ; pas de diff lisible du binaire, rebuild opaque, classement hypothétique).

---

## 2026-10-10 01:47 (Paris) — snapshot `2026-10-09_18-06` → `2026-10-10_01-47`
- Commit GitHub : [2a4e6b1](../../../commit/2a4e6b14a58267f3018f0a2e279fc4bee51623c1)
- Rapport brut : [reports/2026-10-10_01-47.md](reports/2026-10-10_01-47.md) (diffs complets : `reports/2026-10-10_01-47-diffs/`)
- Versions touchées : **agent-store-fuse** `3216860` → `5cf8ba8`. Inchangés : image/box-scripts `91a9d91`, sand-host `91efc0d`, exec-daemon, Chrome 154.0.8037.97, Node, Python, Debian.
- Diff : 0 ajouté, 0 supprimé, 1 modifié (`/usr/local/bin/cursor_agent_store_fuse_version`).

### Bugfix
- —

### Change
- —

### Features
- —

### Perf
- —

### Autre/infra
- Bump du commit agent-store-fuse (fichier de version seul, binaire sans diff lisible) : rebuild opaque, classement hypothétique.

---

## 2026-10-09 18:06 (Paris) — snapshot `2026-10-09_14-24` → `2026-10-09_18-06`
- Commit GitHub : [6b133a0](../../../commit/6b133a0c5f9cb662459f4c41bc17cd297ba9a329)
- Rapport brut : [reports/2026-10-09_18-06.md](reports/2026-10-09_18-06.md) (diffs complets : `reports/2026-10-09_18-06-diffs/`)
- Versions touchées : **sand-host** `91a9d91` → `91efc0d` (upgrade `upgrade-91efc0d`). Inchangés : image/box-scripts `91a9d91`, exec-daemon, Chrome 154.0.8037.97, Node, Python, Debian.
- Diff : 0 ajouté, 0 supprimé, 8 modifiés.

### Bugfix
- —

### Change
- `sand-web-bot-auth.mjs` : suppression des réglages de portée optionnels (marqueurs `/tmp/sand-web-bot-auth-xhr-fetch` et `-iframes`, scope par défaut, attache auto des iframes) ; seuls restent les motifs Document/XHR/Fetch. Simplification de la signature web-bot-auth.

### Features
- —

### Perf
- —

### Autre/infra
- `host-main.cjs`, `sand-eval-runner.cjs`, `agent-store-worker.cjs`, `search-index-worker.cjs`, `box-store-vacuum-worker.cjs` : bundles minifiés reconstruits (rebuild opaque, classement hypothétique).
- `managed-skills/cache.json` : simple rafraîchissement (`fetchedAt`).

---

## 2026-10-09 14:24 (Paris) — snapshot `2026-10-09_14-11` → `2026-10-09_14-24` (Update de l'ordinateur par Dani)
- Commit GitHub : [3eade09](../../../commit/3eade09da946a5ef7cff3861d2234b3cec669d91)
- Rapport brut : [reports/2026-10-09_14-24.md](reports/2026-10-09_14-24.md) (diffs complets : `reports/2026-10-09_14-24-diffs/`)
- Versions touchées : **image** `a6e2020` → `91a9d91`, **box-scripts** `a6e2020` → `91a9d91`, **sand-host** `ddf72f7` → `91a9d91`, **exec-daemon** build 2026-10-06 → 2026-10-08 (canvas SDK `4c3356a5…` → `604ba4f7…`), **Chrome** 154.0.8037.57 → 154.0.8037.97, **npm** (exec-daemon) 10.9.2 → 11.19.1, **pnpm** (corepack) 10.33.4 → 10.34.6. Inchangés : agent-store-fuse `3216860`, orbitd, Node 22.14.0 / 20.19.2, Python 3.13.5, Debian 13.7, uv 0.12.24, pptxgenjs 4.0.1.
- Diff : 2444 ajoutés, 3527 supprimés, 897 modifiés (dont 54 binaires sans diff texte). L'essentiel du volume vient de npm, corepack/pnpm et du déplacement d'officekit.

### Bugfix
- `start-sand-box` : remet aussi `/home/box/chrome-profile` (racine, non récursif) au propriétaire `box`, en plus de `Default/`. Corrige probablement des droits root résiduels sur le profil Chrome.
- `ensure-machine-id` : le dossier parent est créé en tant que `box` (via `runuser`) quand le script tourne en root, et on abandonne proprement si on ne peut pas le créer, au lieu d'un `mkdir` root ignoré.
- Paquets Debian de sécurité/maintenance : `liblzma5` 5.8.1-1+deb13u1 → deb13u2, `libpcre2-8-0` 10.46-1~deb13u2 → deb13u3.
### Change
- **officekit** : la commande `office` passe définitivement par `officekit-ts` (Node). Le wrapper `/usr/local/bin/office` n'a plus de repli. L'ancien binaire natif (≈17 Mo) sous `/opt/sand/sand-host/officekit/bin/` devient un shim de 68 octets, `/usr/local/libexec/sand/officekit/` est supprimé (793 fichiers), et les 792 fichiers de données (templates pptx, taxonomie) passent sous `officekit-ts/share/officekit`.
- Skill gérée **send-on-behalf** : nouvelle section « Choosing and reporting the sending account ». Avant d'envoyer un e-mail, il faut considérer chaque compte possible (Gmail/Outlook connectés, inbox propre), demander lequel utiliser si ce n'est pas clair, puis nommer l'expéditeur exactement tel que le rapporte le résultat de l'envoi.
- Docs de référence (`/home/box/reference/app-ui.md`, `debugging-the-box.md`) : les libellés « Update/Reset Grok Bot's Computer » deviennent « Update/Reset Bot's Computer ».
- Extension WebAuthn proxy de la box : nouvel ID d'extension (`agbllaojcpnneljifnhgkibkiaefijgf` → `mpfkaohnmdhkmmoekeammpleipnbgjjd`, toujours en version 0.1.14), avec le manifeste native-messaging et le XML de mise à jour alignés.
- `pptxgenjs` global : dépendances embarquées (`image-size`, `queue`) retirées de son `node_modules` local (62 fichiers), version inchangée 4.0.1.
### Features
- **Vault** : un nouvel onglet Settings « Vault » est documenté (affiché seulement si activé pour le compte). Le host embarque des messages `List/Put/DeleteGrokBotVaultCredential`, ce qui pointe vers un coffre de credentials côté app. C'est déduit des noms, sans test fonctionnel.
- **SDK canvas `grok/canvas`** : nouveaux fichiers `compile-env.d.ts` (devient le point d'entrée des types), `charts-series.d.ts`, `controls.d.ts`, `dag.d.ts`, `width.d.ts` et un `reference.md` généré (catalogue des composants `Page`, `Table`, `Metrics`, `Segmented`, `Button`, `Input`, `Bars`, `Columns`…). `style.d.ts` et `responsive.d.ts` sont retirés, et les runtimes canvas sont rebuildés.
- **npm 11.19.1** (exec-daemon) : nouvelles commandes `approve-scripts` / `deny-scripts` / `install-scripts` (allow-list des scripts d'install), `stage`, `trust` (OIDC GitHub/GitLab/CircleCI) et `undeprecate`.
- `pack-onepassword-extension` : nouveau mode `--crx`, qui accepte directement le CRX du Chrome Web Store (signature vérifiée côté backend, digest sha256 contrôlé) en plus du tarball, et écrit un fichier `.version`.
### Perf
- `sand-ua-governor.mjs` : l'injection du script UA (Runtime.evaluate, Page.enable, addScriptToEvaluateOnNewDocument) et l'application de l'UA desktop sont lancées en parallèle (`Promise.all`). La reprise `runIfWaitingForDebugger` n'attend plus la fin du traitement UA, donc les nouvelles cibles Chrome démarrent plus vite.
### Autre / infra
- Rebuild du host : `host-main.cjs` (diff ≈ 555k lignes, bundle), `sand-eval-runner.cjs`, `exec-daemon/index.js`. Le contenu détaillé n'est pas analysé au-delà des points ci-dessus.
- `box-store-vacuum-worker` : ajout de constantes de validation (regex de hostname DNS, plafonds de texte 200/3276). `search-index-worker` : simple réordonnancement du bundle voice-call (aucun changement fonctionnel visible).
- `box-chrome` : journalisation optionnelle des URLs ouvertes (`CURSOR_MCP_BROWSER_OPEN_LOG`), déjà vue côté host le 9 oct. 02:53, maintenant aussi dans l'image.
- `cursor-proclist` : `package.json` déclare une liste `files`. Binaires opaques rebuildés : `cursorsandbox`, `gh`, `polished-renderer.node`, `tools/origin`, `sand-daemon-supervisor`, `uv`/`uvx` (même version) et `table-reservation-goat-pp-cli`.

---

## 2026-10-09 12:49 (Paris) — snapshot `2026-10-09_04-50` → `2026-10-09_12-49`
- Commit GitHub : [75ed628](../../../commit/75ed628da4680eae6a106662d1b4d8a156b9867d)
- Rapport brut : [reports/2026-10-09_12-49.md](reports/2026-10-09_12-49.md)
- Versions touchées : **agent-store-fuse** `2394551` → `3216860`. Le reste est inchangé (image `a6e2020`, sand-host `ddf72f7`, orbitd, Chrome, Node).
- Diff : 0 ajouté, 0 supprimé, 2 modifiés.

### Bugfix
—
### Change
- Nouvelle build du module de stores persistants des agents (agent-store-fuse, qui monte `/cursor/stores`) : le binaire `/usr/local/bin/cursor-agent-store-fuse` est remplacé (9 387 440 → 9 403 720 octets, +16 280).
- Le pin `/usr/local/bin/cursor_agent_store_fuse_version` pointe vers l'archive `agent-store-fuse-x64-3216860…-bk.tar.gz`.
### Features
—
### Perf
—
### Autre / infra
- Binaire sans diff texte : on ne peut pas dire ce que contient le commit `3216860` (bugfix, feature ou perf). Rebuild opaque, classement hypothétique.

---

## 2026-10-09 04:50 (Paris) — snapshot `2026-10-09_02-53` → `2026-10-09_04-50`
- Commit GitHub : [c4c8a59](../../../commit/c4c8a59d82adcaceba585c2616822325e0186887)
- Rapport brut : [reports/2026-10-09_04-50.md](reports/2026-10-09_04-50.md)
- Versions touchées : **agent-store-fuse** `3653d35` → `2394551`. Le reste est inchangé (sand-host `ddf72f7`).
- Diff : 0 ajouté, 0 supprimé, 2 modifiés.

### Bugfix
—
### Change
- Nouvelle build d'agent-store-fuse : le binaire `/usr/local/bin/cursor-agent-store-fuse` est remplacé (9 379 480 → 9 387 440 octets, +7 960).
- Le pin `cursor_agent_store_fuse_version` pointe vers `agent-store-fuse-x64-2394551…-bk.tar.gz`.
### Features
—
### Perf
—
### Autre / infra
- Binaire sans diff texte. Rebuild opaque, classement hypothétique.

---

## 2026-10-09 02:53 (Paris) — snapshot `2026-10-08_20-36` → `2026-10-09_02-53`
- Commit GitHub : [db01287](../../../commit/db01287ee448f1f6d886a1b5da66c84e85aedca9)
- Rapport brut : [reports/2026-10-09_02-53.md](reports/2026-10-09_02-53.md)
- Versions touchées : **sand-host** `3f90dc1` → `ddf72f7` (hot-update du host, image inchangée `a6e2020`). Skills gérées `office-pptx`, `slides-executor` et `sign-in`. Dépendance **piscina** 4.9.3 → 4.9.4. agent-store-fuse inchangé (`3653d35`).
- Diff : 0 ajouté, 0 supprimé, 15 modifiés.

### Bugfix
- **piscina 4.9.4** (pool de worker threads du host) : les options du pool, de `run()` et de `destroy/close` sont maintenant recopiées dans des objets **sans prototype** (`withNullPrototype`), pour qu'elles ne soient plus jamais lues à travers une chaîne de prototypes polluée. C'est un durcissement contre la pollution de prototype (`node_modules/piscina/dist/{common,index}.js`).
- **Recherche dans les conversations** (`content-search/search-index-worker.cjs`) : l'identifiant d'agent est validé partout par `parseAgentId` (refuse les ids vides, `/`, `\`, NUL, `.` et `..`). C'est appliqué aux résultats messages/médias, à `applyServerRange` (qui utilise l'id validé pour toutes les requêtes SQL) et à `reconcile()` (les dossiers d'agents au nom invalide sont ignorés). La fonction existait déjà, elle est seulement déplacée et appliquée plus largement : c'est un durcissement contre les ids malformés ou de type « chemin ».
### Change
- **Skill `office-pptx`** réécrite pour le nouvel outillage **officekit-ts** (`/opt/sand/sand-host/officekit-ts`) :
  - `office pptx render-slides` est remplacé par `office render` ;
  - `office pptx rearrange` et `office pptx swap-layout` disparaissent de la skill : couper et réordonner se fait maintenant en python-pptx (exemple de code fourni) ;
  - `check-overlap` ne mesure plus le texte. Il compare seulement les boîtes du XML (texte sur texte, formes hors slide) et ne voit plus le débordement de texte, qu'il faut repérer sur le rendu. `--autofit` pose `normAutofit` sans plancher de 70 % ;
  - `office validate --repair` cible la liste précise des défauts PptxGenJS/python-pptx, et `check-fonts` signale aussi les polices de thème manquantes avec la police de substitution ;
  - `office pptx clean` retire maintenant les slides non listées.
  - C'est un **retrait de capacités** côté CLI (rearrange, swap-layout, mesure de débordement), compensé par des scripts python.
- **Skill `slides-executor`** : la branche « Without the office CLI » est supprimée (−110 lignes). Le setup teste `officekit-ts/bin/office` (au lieu de `officekit/bin`) avec la consigne « Do not install the office CLI ». Les phases storyboard, build, fix et finish passent en sections de premier niveau. L'ordre des slides passe par le script de storyboard, et non plus par `office pptx rearrange`.
- **MCP** (indice lisible dans `host-main.cjs`, champ ajouté aussi dans `agent-store-worker.cjs`) : nouveau champ `McpArgs.arguments_repaired`. Quand il est vrai, le host **ne recoerce plus** les arguments selon le schéma de l'outil (`args.argumentsRepaired ? args : coerceMcpArgsToToolSchema(...)`). Cela évite probablement une double « réparation » des arguments, mais c'est hypothétique faute de source.
### Features
- **Skill `sign-in`** : le Secure Form (`RequestUserForm`) d'une connexion porte maintenant un **titre qui nomme le site** (ex. « Sign in to GitHub ») en plus du **domaine** passé comme `domain`.
- **Connexion source / SCM au-delà de GitHub** (`agent-store-worker.cjs`, retrouvé dans `host-main.cjs`) : `ConnectScmArgs` accepte maintenant les cibles **GitLab** (`instance_id`), **Bitbucket** (`instance_id`) et **Azure DevOps**, plus `reconnect`, `repo_url` et `instance_host` (instances auto-hébergées). Les résultats et approbations renvoient un `scm_target`, et un refus garde les `recorded_args`.
- **`box-chrome`** : nouvelle journalisation optionnelle des URLs ouvertes. Si `CURSOR_MCP_BROWSER_OPEN_LOG` est défini, chaque argument `http(s)://` passé au lanceur Chrome est ajouté au fichier sous la forme `<epoch_ms>\topen\t<url>`.
- **Statut des serveurs MCP** (indice lisible dans `host-main.cjs`) : `statusFromBoxListStatus` reçoit maintenant `pendingSignIn` en plus du détail technique. Un état « connexion en attente » est probablement remonté, mais c'est hypothétique.
### Perf
—
### Autre / infra
- `host-main.cjs` (+21 395 / −16 357 lignes après découpage du minifié) et `sand-eval-runner.cjs` (+19 276 / −11 043) sont reconstruits. Rebuild opaque, classement hypothétique. Les seuls indices exploitables sont reportés ci-dessus (SCM, `argumentsRepaired`, `pendingSignIn`).
- `extensions/codebase-telemetry/csnaps` est un binaire modifié (13 218 728 → 13 212 408 octets), sans diff texte.
- `managed-skills/cache.json` est resynchronisé (reflet des 3 skills ci-dessus).

---

## 2026-10-08 20:36 (Paris) — snapshot `2026-10-08_04-50` → `2026-10-08_20-36`
- Commit GitHub : [85f87f0](../../../commit/85f87f0b35f552659b61837bdaa60e4f0ed818ae)
- Rapport brut : [reports/2026-10-08_20-36.md](reports/2026-10-08_20-36.md)
- Versions touchées : **agent-store-fuse** `82d6332` → `3653d35`. Le reste est inchangé (sand-host `3f90dc1`).
- Diff : 0 ajouté, 0 supprimé, 1 modifié.

### Bugfix
—
### Change
- Le pin `/usr/local/bin/cursor_agent_store_fuse_version` pointe vers `agent-store-fuse-x64-3653d35…-bk.tar.gz`.
### Features
—
### Perf
—
### Autre / infra
- Seul le fichier de version change : le binaire `cursor-agent-store-fuse` a le **même hash** dans ce snapshot. Soit la nouvelle archive est prise en compte plus tard, soit le binaire est identique. Contenu fonctionnel inconnu.

---

## 2026-10-08 04:50 (Paris) — snapshot `2026-10-08_02-46` → `2026-10-08_04-50`
- Commit GitHub : [e7d2329](../../../commit/e7d232949a1b44ce36fe83ddc81de3802455377e)
- Rapport brut : [reports/2026-10-08_04-50.md](reports/2026-10-08_04-50.md)
- Versions touchées : **agent-store-fuse** `e44f687` → `82d6332`. Le reste est inchangé (sand-host `3f90dc1`).
- Diff : 0 ajouté, 0 supprimé, 1 modifié.

### Bugfix
—
### Change
- Le pin `cursor_agent_store_fuse_version` pointe vers `agent-store-fuse-x64-82d6332…-bk.tar.gz`.
### Features
—
### Perf
—
### Autre / infra
- Seul le fichier de version change, le binaire garde le même hash. Contenu fonctionnel inconnu.

---

## 2026-10-08 02:46 (Paris) — snapshot `2026-10-07_17-05` → `2026-10-08_02-46`
- Commit GitHub : [1fae1ef](../../../commit/1fae1efc17132fb9c0ca8cb207be3888384a3123) (publié avec le snapshot de base `2026-10-07_17-05`)
- Rapport brut : [reports/2026-10-08_02-46.md](reports/2026-10-08_02-46.md)
- Versions touchées : **sand-host** `f95dbfb` → `3f90dc1` (hot-update, image inchangée `a6e2020`). Skills gérées : `office-pptx` (nouvelle), `slides-executor`, `slides`, `routines`, `site-playbooks-luma`.
- Diff : 1 ajouté, 0 supprimé, 11 modifiés.

### Bugfix
—
### Change
- **Skill `routines`** : le prompt sauvegardé d'une routine doit maintenant **se terminer par une ligne en clair** qui dit qui reçoit le résultat, où, et ce qui se passe quand il n'y a rien de nouveau. Après une création ou une modification, l'agent attend le résultat de la sauvegarde puis résume la routine.
- **Skill `site-playbooks-luma`** : précise que les données réservées à l'hôte (lien de connexion, liste d'invités) ne sont pas dans le JSON public. `virtual_info.has_access=false` veut dire « lien masqué », pas « absent » : ne pas réessayer.
- **`agent-store-worker.cjs`** : le message `CursorRule` gagne un champ `warnings` (liste de chaînes).
- **Recherche dans les conversations** : nouveau drapeau `fault_recovered` à côté de `fault_news_owed`.
### Features
- **Nouvelle skill gérée `office-pptx`** (+526 lignes) : guide complet pour lire, vérifier, éditer, créer (PptxGenJS) et rendre des `.pptx` avec la CLI `office` (inventory, check-fonts, check-overlap, validate/repair, rearrange, swap-layout, render-slides, thumbnail-grid…).
- **Skills `slides` / `slides-executor`** : nouvelle phase **`clean`**. Quand l'utilisateur est satisfait du deck livré, les fichiers de travail sont supprimés, seul le `.pptx` final reste, et un nouveau statut `cleaned` apparaît. `slides-executor` gagne aussi une branche « With the office CLI » (officekit), avec repli « Without the office CLI ».
- **Recherche dans les conversations** (`search-index-worker.cjs`, +154 / −23) : synchronisation de l'index **par plages venues du serveur**. On trouve une table `server_cursors` (génération, `topped_through_seq`, `backfilled_from_seq`), des colonnes `seq` et des index `(agent_id, seq)` sur messages et médias, et un job `apply-server-range` qui réinitialise, avance ou vérifie. Les agents « possédés par le serveur » ne sont plus modifiés par les écritures locales.
### Perf
- **`sand-ua-governor.mjs`** (réglage user-agent du navigateur de la box via CDP) : les appels `Runtime.evaluate`, `Page.enable` et `addScriptToEvaluateOnNewDocument` sont lancés **en parallèle** (`Promise.all`), comme l'application de l'UA desktop et l'installation du script. Le script n'est retiré que s'il existe, et la reprise `runIfWaitingForDebugger` n'attend plus la fin du traitement UA. Moins d'allers-retours à l'ouverture d'un onglet.
### Autre / infra
- `host-main.cjs` (+144 675 / −119 834 lignes après découpage) et `sand-eval-runner.cjs` (+54 437 / −31 483) sont reconstruits. Rebuild opaque, classement hypothétique. Indice : de nombreuses lignes touchant `RequestUserForm` (+22 / −13).
- `managed-skills/cache.json` est resynchronisé (+167 / −10).
