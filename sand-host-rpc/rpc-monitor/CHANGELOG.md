# Changelog — surface des méthodes RPC du host Sand/Cursor

Suivi snapshot par snapshot des **déclarations RPC** du host : méthodes
**ajoutées / supprimées / modifiées** (changement de `kind` ou des types I/O).
Extraction **exhaustive** : scan de tous les `.cjs/.js/.mjs` de `/opt/sand`,
relevé de chaque `typeName: "<pkg>.<Service>"` + son bloc `methods`, tous
packages confondus. Entrées les plus récentes en haut. Heures de Paris
(CEST/CET). Généré/alimenté par `check.sh`.

<!-- NOUVELLES-ENTREES-ICI -->

## 2026-10-24 (Paris) — baseline

Première capture de référence (`snapshot.sh`).

- **21 services**, **2317 déclarations RPC** distinctes, tous packages :
  `aiserver.v1` (16 services), `agent.v1` (3), `origin.v1` (1), `orbit.v1` (1).
- Fichiers porteurs de définitions : `sand-host/host-main.cjs`,
  `sand-host/sand-eval-runner.cjs`.
- Signature suivie par méthode : `I|O|kind` (type requête, type réponse, kind Connect).
- Les snapshots suivants ne sont listés ici que s'ils présentent un vrai diff.

Voir le snapshot baseline dans `snapshots/` et le manifeste `manifest.json`.
