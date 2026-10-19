# Grok-Bot-app 🍍

Ce dépôt est une **observation en continu de l'app Grok Bot telle qu'elle tourne sur la box** (l'ordinateur cloud de l'agent) : ce qui change d'une mise à jour à l'autre, les schémas des outils MCP exposés à l'agent, et un instantané caviardé de son arborescence et de son prompt. Aucune donnée personnelle : médias, secrets, mémoire et identifiants sont exclus par conception.

## 📓 Changelogs — commence ici

| Changelog | Ce qu'il suit |
|---|---|
| **[box-update-monitor/CHANGELOG.md](box-update-monitor/CHANGELOG.md)** | Runtime de la box, snapshot par snapshot (image, `sand-host`, `agent-store-fuse`, exec-daemon, Chrome/Node/Python…), classé Bugfix / Change / Features / Perf / Autre-infra. |
| **[mcp-schema-monitor/CHANGELOG.md](mcp-schema-monitor/CHANGELOG.md)** | Schémas des outils MCP (noms, descriptions, `inputSchema`/`outputSchema`) et statut des serveurs, jour par jour. |

## 🧭 État courant de la box

Capturé au snapshot `2026-10-19_20-53` (heure de Paris) :

| Composant | Version |
|---|---|
| Image (`image_sha`) | `fcbae14` |
| Host (`sand-host`) | `9d2ac1b` (hot-update `upgrade-9d2ac1b`) |
| agent-store-fuse | `163bab1` |
| Chrome | 155.0.8059.84 |
| Node | 22.23.3 (exec-daemon) / 20.19.2 (système) |
| Python | 3.13.5 |
| Debian | 13.7 · noyau 6.12.94+ |

Rappel : une **« Mise à jour de l'ordinateur »** change l'`image_sha` ; une **mise à jour à chaud** change seulement la version du host (`sand-host`) sans recréer l'image.

## 🗂️ Contenu du dépôt

- **[box-update-monitor/](box-update-monitor/)** — snapshots horodatés + diffs ligne à ligne du runtime. `snapshot.sh`, `diff.sh`, `check-update.sh` (routine horaire), `scan-personal-data.sh` (scan PII bloquant avant publication). → [changelog](box-update-monitor/CHANGELOG.md)
- **[mcp-schema-monitor/](mcp-schema-monitor/)** — suivi quotidien des schémas MCP, calqué sur box-update-monitor (même verrou, même scan PII). → [changelog](mcp-schema-monitor/CHANGELOG.md)
- **[non-stock/](non-stock/)** — arborescence miroir des fichiers non-dpkg des racines box observées (`exec-daemon/`, `home/`, `usr/`, `workspace/`), symlinks non suivis, paquets Debian filtrés, médias/secrets/caches/fichiers > 8 Mio exclus.
- **[context-prompt/](context-prompt/)** — snapshot **verbatim caviardé** du prompt système de l'agent (bloc mémoire → `[REDACTED]`, UUID → `<AGENT_ID>`, utilisateur → `<USER>`).
- **`tool-inventory-exhaustive.json` / `tool-inventory-names-only.json`** — inventaire des outils disponibles à l'agent (built-in, namespace `cursor`, serveurs MCP).

## 🔒 Confidentialité

Chaque publication passe un **scan données perso / secrets bloquant** (`scan-personal-data.sh`) sur exactement ce qui serait poussé : en cas de détection, rien n'est commité. On ne publie que des fichiers app/système génériques et des schémas — jamais de résultats d'appels d'outils, de tokens, de mémoire agent ni d'identité utilisateur. L'identité du modèle/fournisseur sous-jacent n'apparaît nulle part.
