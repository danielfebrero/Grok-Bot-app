# Changelog de l'ordinateur cloud (box Grok Bot)

Snapshot par snapshot, ce qui a changé dans le runtime de la box (sand-host, image, agent-store-fuse, skills gérées…),
classé en **Bugfix / Change / Features / Perf / Autre-infra**. Entrées les plus récentes en haut.
Seuls les snapshots avec un vrai diff sont listés. Heures de Paris (CEST).

Base de référence : snapshot `2026-10-07_17-05` (« avant update (Dani) ») — image `a6e2020`, sand-host `f95dbfb`,
agent-store-fuse `e44f687`, Chrome 154.0.8037.57, Node 22.14.0 / 20.19.2, Python 3.13.5, Debian 13.7.

---

## 2026-10-27 09:45 (Paris) — snapshot `2026-10-27_08-52` → `2026-10-27_09-45`

**Hot-update du host bundle (sand-host) `7cb9acc` → `ab49b45`** (supervisor `upgrade-7cb9acc` → `upgrade-ab49b45`, `pendingUpgradeVersion` = None) : remplacement à chaud du `sand-host` par le sand-supervisor, sans Update de l'ordinateur. Bundle host (`host-main.cjs` 32.3→32.8 Mo, `sand-eval-runner.cjs` 25.4→25.9 Mo), `box-chrome` (Accept-Language via `--accept-lang` + record profil), `sand-ua-governor` / `sand-fingerprint-profiles` (override Accept-Language par session / flag de lancement), `search-index-worker` (normalisation caractères cachés + plafond noms voyageurs vol 59), `agent-store-worker` (`TriggeringUserInfo.followup_id`), et bump concomitant **agent-store-fuse** `62d65171` → `37f15904`. Image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +1 / −0 / ~25.

### Features
- **Skill `subscription-management` (nouveau)** : annuler / lire prix & date de renouvellement d'un abonnement nommé (connectors, stops rétention + confirm final, preuve page/mail, option `place_phone_call`).
- **Suivi post-commande (skill purchases)** : section **After the order** — routine 3×/jour (9h/13h/18h, 7j) mail → Shop → browser ; shopping crée cette routine après confirmation.
- **E-mail agent** : `OfferEmailAddress` remplace `ClaimEmailInbox` ; adresse off jusqu'à activation user (`agent-email` + `no-connector-fallback`).
- **Apps review overrides (Team)** : RPCs `RequestReviewOverride` / `GrantAppReviewOverride` / `RevokeAppReviewOverride` / `CancelReviewOverride` / `ListAppReviewOverrides` + `PublishReviewOverride*` ; règles `Get/List/ActivateTeamAppsReviewRulesVersion`.
- **Sand Team admin** : egress HTTP rules, trusted CA certs, setup-file upload/list/delete/current-version, `UpdateTeamTailnetBoxAccess`.
- **Slack org** : `Get/UpdateOrganizationSlackInboundInstallPolicy`, `ListOrganizationSlackConnectedWorkspaces` / `ListSlackConnectedWorkspaces`.
- **Origin / CI** : `GetOriginCiJobTrace`, `GetMergeQueueTimeline`.
- **Synced client settings** : `GetSyncedClientSettings` / `UpdateSyncedClientSettings`.
- **Plugins / bots** : `ListInstalledPlugins`, `SearchTeamBots`, `SendBonfireInvite`, `PresentCallQuestion`, `ReportSignInAttempt`, `ListPlacementRepoCandidates`, `RevokeCollapseSourceProof`, `WatchSessions`, `ShareProject` / `JoinProjectSource`.
- **Shopping cards** : badges shortlist (Top rated / Lowest price / Best value / …, max 2).
- **Chrome Accept-Language** : `box-chrome` lit `/tmp/sand-browser-accept-language-*` et pose `--accept-lang` ; governor n'écrase plus le header du contexte déjà flaggé.

### Change
- **Spendlog** : montant `total CURRENCY` (plus `$total` fixe).
- **Orders & returns** : ne pas re-proposer une routine de suivi si `<automation_status>` en a déjà une pour la commande.
- **Sign-in / in-chat-forms** : Secure Form 2FA/OTP sans explication vault.
- **Vols** : plafond noms voyageurs 60→59 (`search-index-worker`).
- **recent-changes** : note Action Button (2026-10-19) retirée du feed.

### Autre-infra
- agent-store-fuse pointeur S3 `…-62d65171…-bk.tar.gz` → `…-37f15904…-bk.tar.gz` (rebuild concomitant au hot-update host).
- `csnaps` (codebase-telemetry) 14.38→14.43 Mo.

Rapport : [reports/2026-10-27_09-45.md](reports/2026-10-27_09-45.md).

## 2026-10-27 08:52 (Paris) — snapshot `2026-10-27_05-51` → `2026-10-27_08-52`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `d7000925de3e23cbfe62f00c2bfd7ad2899b5aea` → `62d651713b82c6bea01bac2e91a533e9a6936d9b` (tarball S3). Toutes les autres versions plateforme inchangées (image `fcbae14`, box-scripts `fcbae14`, sand-host `7cb9acc`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7). Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-27_08-52.md](reports/2026-10-27_08-52.md).

## 2026-10-27 05:51 (Paris) — snapshot `2026-10-27_04-53` → `2026-10-27_05-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `12403a2cfce6e24c6bca0266680c0d60d145d87f` → `d7000925de3e23cbfe62f00c2bfd7ad2899b5aea` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `7cb9acc` (supervisor `upgrade-7cb9acc`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-27_05-51.md](reports/2026-10-27_05-51.md).

## 2026-10-27 04:53 (Paris) — snapshot `2026-10-27_03-51` → `2026-10-27_04-53`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `309eaeafa6fa6496f9fc47d8511ef9d3de3e6c2b` → `12403a2cfce6e24c6bca0266680c0d60d145d87f` (tarball S3). Toutes les autres versions plateforme inchangées (image `fcbae14`, sand-host `7cb9acc`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7). Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-27_04-53.md](reports/2026-10-27_04-53.md).

## 2026-10-27 03:51 (Paris) — snapshot `2026-10-27_01-50` → `2026-10-27_03-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `eaf43fac8ff50845c39fdeba5be4a2ec61753078` → `309eaeafa6fa6496f9fc47d8511ef9d3de3e6c2b` (tarball S3). Toutes les autres versions plateforme inchangées (image `fcbae14`, sand-host `7cb9acc`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7). Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-27_03-51.md](reports/2026-10-27_03-51.md).

## 2026-10-27 01:50 (Paris) — snapshot `2026-10-26_23-47` → `2026-10-27_01-50`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `cc66aefe34c497ccab52d7f8714c374e6b1a3272` → `eaf43fac8ff50845c39fdeba5be4a2ec61753078` (tarball S3). Toutes les autres versions plateforme inchangées (image `fcbae14`, sand-host `7cb9acc`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7). Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-27_01-50.md](reports/2026-10-27_01-50.md).

## 2026-10-26 23:47 (Paris) — snapshot `2026-10-26_22-48` → `2026-10-26_23-47`

**Rebuild agent-store-fuse uniquement (binaire inclus)** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `24f5bc6cca764850cc2bbc2ad54783d722355a82` → `cc66aefe34c497ccab52d7f8714c374e6b1a3272` (tarball S3), et cette fois le binaire `cursor-agent-store-fuse` est reconstruit : 9501280 → 9505248 o, sha256 `9a7669a7d70e…` → `94f62f30259e…`. Toutes les autres versions plateforme inchangées (image `fcbae14`, sand-host `7cb9acc`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7). Fichiers : +0 / −0 / ~2.

Rapport : [reports/2026-10-26_23-47.md](reports/2026-10-26_23-47.md).

## 2026-10-26 22:48 (Paris) — snapshot `2026-10-26_20-54` → `2026-10-26_22-48`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `a872edf440984297b6b78cfb3c2b6bc60830be50` → `24f5bc6cca764850cc2bbc2ad54783d722355a82` (tarball S3). Toutes les autres versions plateforme inchangées (image `fcbae14`, sand-host `7cb9acc`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7). Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-26_22-48.md](reports/2026-10-26_22-48.md).

## 2026-10-26 20:54 (Paris) — snapshot `2026-10-26_19-48` → `2026-10-26_20-54`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `c76f6fef260691a89fcf718bcbf74b4078b3495f` → `a872edf440984297b6b78cfb3c2b6bc60830be50` (tarball S3). Toutes les autres versions plateforme inchangées (image `fcbae14`, sand-host `7cb9acc`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7). Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-26_20-54.md](reports/2026-10-26_20-54.md).

## 2026-10-26 19:48 (Paris) — snapshot `2026-10-26_17-51` → `2026-10-26_19-48`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `5d804d3509e95db6a5edefae929ffb630ee7bb92` → `c76f6fef260691a89fcf718bcbf74b4078b3495f` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9501280 o, même sha256 `9a7669a7…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `7cb9acc` (supervisor `upgrade-7cb9acc`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-26_19-48.md](reports/2026-10-26_19-48.md).

## 2026-10-26 17:51 (Paris) — snapshot `2026-10-26_11-52` → `2026-10-26_17-51`

**Hot-update du host bundle (sand-host) `85c241a` → `7cb9acc`** (supervisor `upgrade-85c241a` → `upgrade-7cb9acc`, `pendingUpgradeVersion` = None) : remplacement à chaud du `sand-host` par le sand-supervisor, sans Update de l'ordinateur. Bundle host (`host-main.cjs` 32.1→32.3 Mo, `sand-eval-runner.cjs` 25.3→25.4 Mo), `sand-egress-tunnel` 2.93→2.94 Mo, scripts desktop/fenêtre (attente Xvfb 100→150, supervision VNC + codes de sortie 72/74), `search-index-worker.cjs` (parsers flight), `office.mjs`, et refetch skills gérées `purchases` + `scheduling`. Toutes les autres versions plateforme sont inchangées : image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, **agent-store-fuse `5d804d35` inchangé**, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7. Fichiers : +0 / −0 / ~12.

### Features
- **Outils vol Grok Bot** : nouveaux tools `FlightSearch` / `FlightDetails` + carte `flight_travelers` / balises `<flight_travelers_reply>` (host-main + search-index-worker : validation IATA, voyageurs, offres).
- **Publish state projets** : RPC `ListProjectPublishState` (`project_bc_id` → runs `ProjectRunPublishState` : proposals/apps/secrets_requests/running/error) + message `ProjectRunSecretsRequest`.
- **Inbox buckets** : `GetInboxBuckets` (buckets + badge + sections), `GetFocusedGroupsOfChanges`, `MarkPullRequestsAsViewed` (par change_ids ou focused_groups).
- **Connecteur Grok** : `SetGrokConnectorEnabled` (`server_identifier`, `enabled`).
- **Apps security** : `AppsSettings.security` → `AppsSecuritySettings` (`review_enabled`, `acknowledged_disable_review`, `allow_review_overrides`).
- **Attribution PR → agent** : `GetPrAgentAttribution` (`pr_url`, `expected_scope` → thread/project/grok_bot_agent_id).
- **Retry enregistrement appel sortant** : `RetryGrokBotOutboundCallRecording` (`outbound_call_id` → `accepted`).
- **Repo never-draft** : `UpdateRepoNeverDraftSetting` (`pull_requests_never_draft`).
- **Watch page** : `GrokBotTranscriptWatchPageTitleChanged` gagne `6 content_changed`.
- **Paiements (skill purchases)** : section **Bank verification** (3-D Secure / app banque comme étape humaine après submit) ; ordre « How to pay » réécrit (nommé → X Money si outil OK → carte perso) ; X Money n'est plus systématiquement exclu des offres wallet.
- **Scheduling (skill)** : connecteur calendrier manquant → connect card immédiate (Outlook si mail M365, sinon Google), sans question widget préalable.

### Bugfix / Change
- **Fenêtre box / VNC** : `start-window` exige display **et** VNC vivant (`vnc_alive` = token + port en écoute) ; nouveaux exit codes `WINDOW_VNC_NOT_STARTED=72` / `WINDOW_VNC_UNAVAILABLE=74` mappés côté host (`box_fork_window_vnc_*`) ; cleanup via `sand_desktop_unregister_group` + `end_desktop`.
- **Démarrage desktop** : boucle d'attente Xvfb `seq 1 100` → `seq 1 150` dans `start-desktop.sh`.
- **Local exec** : `classifyLocalExecFailure` (spawn/errno) pour classer les échecs d'exec local.

### Autre / infra
- `managed-skills/cache.json` : `fetchedAt` + corps `purchases` / `scheduling` (pas seulement timestamp).
- `officekit-ts/lib/office.mjs` : gros repack (noms minifiés, ~+878 o) — pas de surface fonctionnelle isolée hors host tools.
- agent-store-fuse inchangé ce cycle.

Rapport : [reports/2026-10-26_17-51.md](reports/2026-10-26_17-51.md).

## 2026-10-26 11:52 (Paris) — snapshot `2026-10-26_09-53` → `2026-10-26_11-52`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `9baa34c0fd4c8a076b95c76b858a532a2a7156c0` → `5d804d3509e95db6a5edefae929ffb630ee7bb92` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9501280 o, même sha256 `9a7669a7…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `85c241a` (supervisor `upgrade-85c241a`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-26_11-52.md](reports/2026-10-26_11-52.md).

## 2026-10-26 09:53 (Paris) — snapshot `2026-10-26_05-53` → `2026-10-26_09-53`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `1b2084e22d6f2459dad93a520d18d3c8c2d9f823` → `9baa34c0fd4c8a076b95c76b858a532a2a7156c0` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9501280 o, même sha256 `9a7669a7…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `85c241a` (supervisor `upgrade-85c241a`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-26_09-53.md](reports/2026-10-26_09-53.md).

## 2026-10-26 05:53 (Paris) — snapshot `2026-10-26_04-51` → `2026-10-26_05-53`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `1645ea012ff21062f4d3cf709d600dd53a569fbb` → `1b2084e22d6f2459dad93a520d18d3c8c2d9f823` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9501280 o, même sha256 `9a7669a7…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `85c241a` (supervisor `upgrade-85c241a`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-26_05-53.md](reports/2026-10-26_05-53.md).

## 2026-10-26 04:51 (Paris) — snapshot `2026-10-26_02-52` → `2026-10-26_04-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `e1ba77ac15d9b30e559d74eb74f2f5360ec3250a` → `1645ea012ff21062f4d3cf709d600dd53a569fbb` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9501280 o, même sha256 `9a7669a7…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `85c241a` (supervisor `upgrade-85c241a`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-26_04-51.md](reports/2026-10-26_04-51.md).

## 2026-10-26 02:52 (Paris) — snapshot `2026-10-26_00-51` → `2026-10-26_02-52`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `e4c1d27ca28d66e432733238906ebf64340465df` → `e1ba77ac15d9b30e559d74eb74f2f5360ec3250a` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9501280 o, même sha256 `9a7669a7…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `85c241a` (supervisor `upgrade-85c241a`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-26_02-52.md](reports/2026-10-26_02-52.md).

## 2026-10-26 00:51 (Paris) — snapshot `2026-10-25_23-53` → `2026-10-26_00-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `db7c8404adfd5f142b8c8d7112ce9ea4e6305632` → `e4c1d27ca28d66e432733238906ebf64340465df` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9501280 o, même sha256 `9a7669a7…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `85c241a` (supervisor `upgrade-85c241a`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-26_00-51.md](reports/2026-10-26_00-51.md).

## 2026-10-25 23:53 (Paris) — snapshot `2026-10-25_22-52` → `2026-10-25_23-53`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `f457d3385ac3d811f2def9fece9f41ab0fa47db2` → `db7c8404adfd5f142b8c8d7112ce9ea4e6305632` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9501280 o, même sha256 `9a7669a7…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `85c241a` (supervisor `upgrade-85c241a`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-25_23-53.md](reports/2026-10-25_23-53.md).

## 2026-10-25 22:52 (Paris) — snapshot `2026-10-25_16-50` → `2026-10-25_22-52`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `27a19f855e6a132ae1a9db63d4eb61a54fb04e73` → `f457d3385ac3d811f2def9fece9f41ab0fa47db2` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `85c241a` (supervisor `upgrade-85c241a`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-25_22-52.md](reports/2026-10-25_22-52.md).

## 2026-10-24 22:49 (Paris) — snapshot `2026-10-24_20-52` → `2026-10-24_22-49`

**Bump agent-store-fuse + install `apt` côté agent (`zstd`)** — aucun push plateforme. Deux changements ce cycle :

1. **Rebuild/bump agent-store-fuse** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `204ac1218bbc24e75c7654a3efda1b51edbe0dd1` → `831fcbe7af021645b440d7e4d728630e01d9cd8a` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`).
2. **Nouveau paquet `zstd 1.5.7+dfsg-1`** apparu dans `dpkg` — installé par un agent via `apt-get install -y zstd` (`Requested-By: box (1000)`, `2026-10-24 21:45`), pas un push plateforme.

Aucune version plateforme ne bouge : image `fcbae14`, box-scripts `fcbae14`, sand-host `ef8271b` (supervisor `upgrade-ef8271b`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1 (reste du diff capturé dans `meta/dpkg-packages.txt`).

### Autre / infra
- `zstd` installé côté agent (apt), pas une livraison plateforme.

Rapport : [reports/2026-10-24_22-49.md](reports/2026-10-24_22-49.md).


## 2026-10-24 18:44 (Paris) — snapshot `2026-10-24_17-47` → `2026-10-24_18-44`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `35aaf402a3611d873755a963d906dcca0d685d19` → `24854d212a9e215aba99c6d69c7c80d3114ac660` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `ef8271b` (supervisor `upgrade-ef8271b`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-24_18-44.md](reports/2026-10-24_18-44.md).

## 2026-10-24 17:47 (Paris) — snapshot `2026-10-24_13-53` → `2026-10-24_17-47`

**Hot-update du host bundle (sand-host) `ab3b435` → `ef8271b`** (supervisor `upgrade-ab3b435` → `upgrade-ef8271b`, `pendingUpgradeVersion` = None) : remplacement à chaud du `sand-host` par le sand-supervisor, sans Update de l'ordinateur. Toutes les autres versions plateforme sont inchangées : image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, **agent-store-fuse `35aaf402` inchangé**, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7. Fichiers : +0 / −0 / ~8.

### Features
- **Télémétrie de coût estimé** : nouveau message proto `EstimatedCost` + enum `EstimatedCostSource` (`UNSPECIFIED`/`CATALOG`/`USAGE`), et `TurnEndedUpdate` gagne le champ `7 estimated_cost` (host-main.cjs). La fin de tour peut désormais rapporter un coût estimé, sourcé du catalogue ou de l'usage.
- **Appels téléphoniques (phone-call)** : `search-index-worker.cjs` reconnaît un préfixe d'entrée `phone-call-question:` et masque dans la timeline les `send-message` de type `user-form` dont l'id commence par ce préfixe (question d'appel téléphonique cachée de l'historique visible).
- **WriteCanvas** : `WriteCanvasSuccess` gagne un champ `4 page_url` (en plus de `canvas_id`, `title`, `url`) — lien de page retourné à la création d'un canvas.
- **ComputerUse** : `ComputerUseArgs` gagne un champ `12 capture_name_prefix` (préfixe de nom pour les captures d'écran).

### Change
- **SandboxPolicy** : nouveau champ `16 allow_user_override_read_paths` (bool) dans la policy de sandbox (host-main.cjs + agent-store-worker.cjs) — permet d'autoriser l'utilisateur à surcharger les chemins de lecture.

### Bugfix
- **Timers longs** : le `realClock.setTimeout` est réécrit pour découper un délai en segments de `MAX_TIMER_DELAY_MS` (2147483647 ms, ~24,8 j) et se ré-armer, au lieu d'un seul `setTimeout` (qui déborde l'entier 32 bits de Node au-delà de ~24,8 j et tire immédiatement). `MAX_TIMER_DELAY_MS` déplacé dans `init_clock`.

### Autre / infra
- `agent-store-worker.cjs` : micro-changement (`let copied = false;` → `let copied;` dans le chemin de quarantaine de DB).
- `pdf-worker.js` : helper de bundling `__toESM` ajouté + imports `node:module` / `node:path` (repackaging esbuild, pas de changement fonctionnel visible).
- `officekit-ts/.../sports_strategy_playbook.js` : ajout d'un commentaire `// oxlint-disable-next-line no-useless-assignment`.
- `managed-skills/cache.json` : seul `fetchedAt` change (resync du cache).

Rapport : [reports/2026-10-24_17-47.md](reports/2026-10-24_17-47.md).

## 2026-10-24 13:53 (Paris) — snapshot `2026-10-24_12-48` → `2026-10-24_13-53`

**Nouveaux paquets réseau/capture installés (image inchangée → install `apt` côté agent, pas un push plateforme)** : la suite **Wireshark/TShark/tcpdump** apparaît dans `dpkg`. Ajouts : `tcpdump 4.99.5-2`, `tshark 4.4.19-0+deb13u1`, `wireshark-common 4.4.19`, `libwireshark18`, `libwiretap15`, `libwsutil16`, `libwireshark-data`, `libpcap0.8t64 1.10.5-2`, `libmaxminddb0 1.12.2-1`, `libnl-genl-3-200`, `libbcg729-0`, `liblua5.4-0`, `libsbc1`, `libsmi2t64`, `libspandsp2t64`, `libspeexdsp1`. Aucune version plateforme ne bouge : image `fcbae14`, box-scripts `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, agent-store-fuse `35aaf402`, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés.

### Autre / infra
- `managed-skills/cache.json` : seul le champ `fetchedAt` change (resync du cache, +1/−1), pas de contenu de skill modifié.
- Fichiers : +0 / −0 / ~1 (le reste du diff est capturé dans `meta/dpkg-packages.txt`).

Rapport : [reports/2026-10-24_13-53.md](reports/2026-10-24_13-53.md).

## 2026-10-24 12:48 (Paris) — snapshot `2026-10-24_11-48` → `2026-10-24_12-48`

**Refresh plateforme des skills gérées (managed-skills) Cursor** : 14 `SKILL.md` mis à jour + `cache.json` (contenu de skill, pas seulement `fetchedAt`). Skills touchées : `accommodation-booking`, `box-desktop`, `flight-booking`, `food-ordering`, `grocery-shopping`, `in-chat-forms`, `orders-and-returns`, `purchases`, `recent-changes`, `restaurant-booking`, `scheduling`, `shopping`, `sign-in`, `site-playbooks-target`. Les diffs ajoutent surtout des chemins de paiement (vault `purchases`, cartes enregistrées, Secure Form) et des précisions de flux. Aucune version plateforme ne bouge : image `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, **agent-store-fuse `35aaf402` inchangé**, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés.

### Change
- Contenu de 14 skills gérées rafraîchi par la plateforme (push Cursor, hors agent-store-fuse). Fichiers : +0 / −0 / ~15.

Rapport : [reports/2026-10-24_12-48.md](reports/2026-10-24_12-48.md).

## 2026-10-24 11:48 (Paris) — snapshot `2026-10-24_09-53` → `2026-10-24_11-48`

**Suite netfilter/iptables/nftables installée (install `apt` côté agent, pas un push plateforme)** — cohérent avec le travail `.sand-live-intercept` (capture paquets, ARM ~11:59). Ajouts `dpkg` : `iptables 1.8.11-2`, `nftables 1.1.3-1`, `libip4tc2`, `libip6tc2`, `libnetfilter-conntrack3 1.1.0-1`, `libnfnetlink0 1.0.2-3`, `libnftables1 1.1.3-1`, `libnftnl11 1.2.9-1`. **En plus**, rebuild/bump agent-store-fuse : pointeur `cursor_agent_store_fuse_version` `141f384c25dc5b4923687c36bb14a89cfce436f6` → `35aaf402a3611d873755a963d906dcca0d685d19`. Aucune version plateforme ne bouge : image `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés.

### Autre / infra
- Nouveaux paquets réseau (iptables/nftables) installés côté agent + bump agent-store-fuse. Fichiers : +0 / −0 / ~1 (pointeur fuse ; reste capturé dans `meta/dpkg-packages.txt`).

Rapport : [reports/2026-10-24_11-48.md](reports/2026-10-24_11-48.md).

## 2026-10-24 09:53 (Paris) — snapshot `2026-10-24_06-44` → `2026-10-24_09-53`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `c23654e8c855cf98fa7ee28ab8cd16e1e8cd7410` → `141f384c25dc5b4923687c36bb14a89cfce436f6` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-24_09-53.md](reports/2026-10-24_09-53.md).

## 2026-10-24 06:44 (Paris) — snapshot `2026-10-24_05-50` → `2026-10-24_06-44`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `4301d59ac3d89a43f46b5cadec3bc1676a27722c` → `c23654e8c855cf98fa7ee28ab8cd16e1e8cd7410` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-24_06-44.md](reports/2026-10-24_06-44.md).

## 2026-10-24 05:50 (Paris) — snapshot `2026-10-24_03-51` → `2026-10-24_05-50`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `d3e8350a691e9937b86f982fab5ace3764e5003e` → `4301d59ac3d89a43f46b5cadec3bc1676a27722c` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-24_05-50.md](reports/2026-10-24_05-50.md).

## 2026-10-24 03:51 (Paris) — snapshot `2026-10-24_02-51` → `2026-10-24_03-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `8e0791472b33a41ed311823384eb3f1f6616f238` → `d3e8350a691e9937b86f982fab5ace3764e5003e` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-24_03-51.md](reports/2026-10-24_03-51.md).


## 2026-10-24 02:51 (Paris) — snapshot `2026-10-24_01-51` → `2026-10-24_02-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `1f853a942db3294376f46deaa9f588f15b17cfa6` → `8e0791472b33a41ed311823384eb3f1f6616f238` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-24_02-51.md](reports/2026-10-24_02-51.md).

## 2026-10-24 01:51 (Paris) — snapshot `2026-10-24_00-54` → `2026-10-24_01-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `3bf638238eff06a7b49e5dc955af9952fad2fc08` → `1f853a942db3294376f46deaa9f588f15b17cfa6` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-24_01-51.md](reports/2026-10-24_01-51.md).


## 2026-10-24 00:54 (Paris) — snapshot `2026-10-23_23-45` → `2026-10-24_00-54`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `afde87fd9eb33c2ee5bde8f00326ded6000249e7` → `3bf638238eff06a7b49e5dc955af9952fad2fc08` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-24_00-54.md](reports/2026-10-24_00-54.md).


---

## 2026-10-23 23:45 (Paris) — snapshot `2026-10-23_22-54` → `2026-10-23_23-45`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `ba25ad859e41f02131402c96e45a37a1cff64fd6` → `afde87fd9eb33c2ee5bde8f00326ded6000249e7` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-23_23-45.md](reports/2026-10-23_23-45.md).

## 2026-10-23 22:54 (Paris) — snapshot `2026-10-23_21-52` → `2026-10-23_22-54`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `066f67b61f88e4e643edc101c02f150af28d4192` → `ba25ad859e41f02131402c96e45a37a1cff64fd6` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-23_22-54.md](reports/2026-10-23_22-54.md).


---

## 2026-10-23 21:52 (Paris) — snapshot `2026-10-23_20-51` → `2026-10-23_21-52`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `641bc36cdedce93113dbb71675fc413377941f18` → `066f67b61f88e4e643edc101c02f150af28d4192` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-23_21-52.md](reports/2026-10-23_21-52.md).

## 2026-10-23 20:51 (Paris) — snapshot `2026-10-23_18-47` → `2026-10-23_20-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `48149515a9b2802709207fc5e6ef233b8696ac4e` → `641bc36cdedce93113dbb71675fc413377941f18` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `ab3b435` (supervisor `upgrade-ab3b435`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-23_20-51.md](reports/2026-10-23_20-51.md).


---

## 2026-10-23 18:47 (Paris) — snapshot `2026-10-23_14-52` → `2026-10-23_18-47`

**Hot-update sand-host `a92b6f6` → `ab3b435`** (+ bump agent-store-fuse + refetch skills gérées). Bundle host (`host-main.cjs` 31.6→31.8 Mo, `sand-eval-runner.cjs` 24.9→25.0 Mo), `sand-egress-tunnel` 2.92→2.93 Mo, session Chrome (`RestoreOnStartup:1`, ledger sign-out cookies, `box-dbus` flock), 14 skills commerce/booking + `channels`/`sign-in`. Image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +1 / −0 / ~34.

- **Bugfix** : cookies de session Chrome — `box-chrome-policy` force `RestoreOnStartup:1` (écriture atomique rename) pour que Linux ne purge pas les session cookies à la fermeture de la dernière fenêtre ; `box-chrome` v5 crée `Default/Sessions` et `Sessions_Encrypted` en mode 555 vides (bloque l'écriture Chrome qui sinon restaure des onglets morts). Retrait du mode GPU `spoof_gpu` / `--use-angle=gl` (reste `SAND_CHROME_LEGACY_GPU`).
- **Change** : `sand-session-sync.mjs` + `cdp-cookies.mjs` + `sand-cookie-persist.mjs` — ledger sign-out (`/home/box/sand-data/chrome-cookie-signouts.json`, digests sha256, max 512 clés) : les cookies signés-out ne sont plus réinjectés depuis le seed ; `deleteCookies` les retire du jar live. Nouveau `box-scripts/box-dbus` (flock sur le socket dbus, kill du holder stale). `start-desktop.sh` / `start-window` : bus session sous `unix:path=$XDG_RUNTIME_DIR/bus` via `box-dbus` (guid stable). Skills commerce : paiement = d'abord méthode **déjà sauvée sur le compte du site** (pas une carte du vault utilisateur) ; règle « montrer dans la langue de l'utilisateur » ; cartes `shipment` aussi après paiement order-card (`food-ordering`, `shopping`) ; `channels` adresse un thread DM `slack:<D>:<ts>` ; `shopping` Pay-now depuis la shopping card sans re-search.
- **Features** : RPCs host nouveaux — `CreateGrokBotAgentAccessInvite` / `ListGrokBotAgentAccessInvites`, `CreateBonfireInvite` / `GetBonfireStatus`, `RequestOnePasswordLogins`, `SetProactiveMessaging`, `GetCiFailureInsights`, `GetSecurityAgentSuggestedRepos`, `GetCursorInboxPrAction`, `Get/SetPullRequestDraftPreference`, `ListProjectPublishState`, `UpdateTeamTailnetCompanyCa`, `DeleteOriginUserSlackReviewReminder`, `CreateChangeDraftOverride`.
- **Perf** : —
- **Autre/infra** : hot-update sand-host `ab3b435` (supervisor `upgrade-ab3b435`) ; agent-store-fuse pointeur tarball S3 `011306c9…` → `48149515…` (binaire identique 9504880 o) ; `managed-skills/cache.json` refetch (`fetchedAt` + 14 SKILL.md : accommodation/flight/food/grocery/job/orders/restaurant×2/rideshare/scheduling/shopping + channels/sign-in/site-playbooks-instacart).

Rapport : [reports/2026-10-23_18-47.md](reports/2026-10-23_18-47.md).


---

## 2026-10-23 14:52 (Paris) — snapshot `2026-10-23_09-50` → `2026-10-23_14-52`

**Toolchain C/C++ (clang/LLVM 19 + cppcheck) installé + bump agent-store-fuse.** Côté composants plateforme surveillés, **seul** `agent-store-fuse` a bougé : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `71c0b034eb0f12f3317f29a89191167a4a050f3b` → `011306c935558a4c91141f612d29a8e9bcb6d635` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`), binaire `cursor-agent-store-fuse` inchangé.

En revanche **33 paquets dpkg ajoutés** (0 retiré) — très vraisemblablement un `apt-get install` d'un agent de l'équipe, pas une mise à jour de l'ordinateur : chaîne clang/LLVM 19 (`clang`, `clang-19`, `clang-tidy`, `clang-tools`, `llvm-19` + `-dev`/`-runtime`/`-tools`/`-linker-tools`, `libclang*-19`, `libclang-rt-19-dev`), analyseur statique `cppcheck`, support 32-bit (`lib32gcc-s1`, `lib32stdc++6`, `libc6-i386`), headers de dev (`libffi-dev`, `libncurses-dev`, `libxml2-dev`, `libz3-dev`), `libobjc-14-dev`/`libobjc4`, `binfmt-support`, `libpfm4`, `libpipeline1`, `libtinyxml2-11`, `libyaml-0-2`, `libgc1`, `python3-yaml`, `python3-pygments`. Python dists apparus en conséquence : `PyYAML-6.0.2`, `pygments-2.18.0`.

Composants plateforme **inchangés** : image `fcbae14`, box-scripts `fcbae14`, sand-host `a92b6f6` (supervisor `upgrade-a92b6f6`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7. Fichiers texte surveillés : +0 / −0 / ~1 (pointeur fuse).

Rapport : [reports/2026-10-23_14-52.md](reports/2026-10-23_14-52.md).


---

## 2026-10-23 09:50 (Paris) — snapshot `2026-10-23_08-53` → `2026-10-23_09-50`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `d15d9c1024bc977a1c7a1d6ff143847ff52ad361` → `71c0b034eb0f12f3317f29a89191167a4a050f3b` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `a92b6f6` (supervisor `upgrade-a92b6f6`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-23_09-50.md](reports/2026-10-23_09-50.md).


---


## 2026-10-23 08:53 (Paris) — snapshot `2026-10-23_06-51` → `2026-10-23_08-53`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `db1f6e02d456127165e2736ba8e1335f59a6a7f8` → `d15d9c1024bc977a1c7a1d6ff143847ff52ad361` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `a92b6f6` (supervisor `upgrade-a92b6f6`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-23_08-53.md](reports/2026-10-23_08-53.md).


---


## 2026-10-23 06:51 (Paris) — snapshot `2026-10-23_03-52` → `2026-10-23_06-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `e9b89684f72d5ffdf4a1b9c23c99c2290a7e3181` → `db1f6e02d456127165e2736ba8e1335f59a6a7f8` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `a92b6f6` (supervisor `upgrade-a92b6f6`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-23_06-51.md](reports/2026-10-23_06-51.md).


---


## 2026-10-23 03:52 (Paris) — snapshot `2026-10-22_23-47` → `2026-10-23_03-52`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `172ca22b3e2f37d6f94fa51e928c89edda10e464` → `e9b89684f72d5ffdf4a1b9c23c99c2290a7e3181` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256 `5c3513e4…`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `a92b6f6` (supervisor `upgrade-a92b6f6`), exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

Rapport : [reports/2026-10-23_03-52.md](reports/2026-10-23_03-52.md).


---

## 2026-10-22 23:47 (Paris) — snapshot `2026-10-22_18-53` → `2026-10-22_23-47`

**Hot-update sand-host `4006392` → `a92b6f6`** (+ bump agent-store-fuse). Bundle host (`host-main.cjs` 31.5→31.6 Mo, `sand-eval-runner.cjs` 24.8→24.9 Mo), `box-chrome` (launch markers), `sand-cookie-persist.mjs` (`--save-port`), proto agent-store-worker, skills gérées (`purchases`, `routines`, `recent-changes` + note 1Password). Image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~13.

- **Bugfix** : —
- **Change** : `box-chrome` — enregistre les marqueurs stealth actifs au cold launch (`/tmp/sand-box-browser-launch-markers-<port>` : hardware_profile, coherent_fingerprint, spoof_gpu, webrtc_proxy_only, no_local_dns, disable_quic, accept_language, egress_timezone, fail_closed) via `marker_gate_on`. `sand-cookie-persist.mjs` — mode `--save-port` pour rafraîchir le seed depuis le jar live (save-no-seed / save-saved / save-unchanged / save-paused / save-failed). Skill `purchases` : guide X Money embarqué sous `## X Money guide` (sinon Read du catalog). Skill `routines` : une suggestion d'économie d'usage attend le yes de l'utilisateur. Proto `ConversationStateStructure` : champ `project_coordinator_system_section`.
- **Features** : note managed-skills `recent-changes` — intégration 1Password brokered (approbation dans l'app 1Password, bot ne voit pas les mots de passe ; rollout progressif depuis le 22/10). Nouveaux RPC host : `GetGrokBotAgentCustomMcp` / `UpdateGrokBotAgentCustomMcp`, `GetGrokBotMessagingLinkStatus`, `GetGrokBotVoiceCall` / `RecordGrokBotVoiceCall`, `GetConflictInsights`, `ListOriginCiJobs`, `RecordCliUsage`, `RequestFactoryVulnerabilityPatch` / `RequestFactoryVulnerabilityVerification`, `SetBackgroundComposerVisibility`. Flag host `grok_bot_slack_sharing`.
- **Perf** : —
- **Autre/infra** : hot-update sand-host `a92b6f6` (supervisor `upgrade-a92b6f6`) ; agent-store-fuse pointeur tarball S3 `ef71388e…` → `172ca22b…` (binaire identique 9504880 o) ; retrait apparent des flags `sand_enable_pressure_cpu_profiler` et `sand_web_bot_auth_signing` du bundle host.

Rapport : [reports/2026-10-22_23-47.md](reports/2026-10-22_23-47.md).


---

## 2026-10-22 18:53 (Paris) — snapshot `2026-10-22_17-54` → `2026-10-22_18-53`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `023c11162a46b7dfbe5f7e800bf0c2f4b1063f82` → `ef71388e495e596274f4c3901dc845b80c7c53f0` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Binaire `cursor-agent-store-fuse` inchangé (9504880 o, même sha256). Image `fcbae14`, box-scripts `fcbae14`, sand-host `4006392`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump pointeur tarball S3 agent-store-fuse uniquement (binaire identique).

Rapport : [reports/2026-10-22_18-53.md](reports/2026-10-22_18-53.md).


---

## 2026-10-22 17:54 (Paris) — snapshot `2026-10-22_14-51` → `2026-10-22_17-54`

**Hot-update sand-host `fd395ed` → `4006392`** (+ bump agent-store-fuse). Bundle host (`host-main.cjs` 31.3→31.5 Mo, `sand-eval-runner.cjs` 24.7→24.8 Mo), proto agent-store-worker (`CursorRule.body_omitted`), skill `sign-in` (1Password : `ListCredentials` → `GetCredentialProviderStatus`), `managed-skills/cache.json` refetch. Image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Meta dpkg : apparition de `rsync` + `libpopt0`. Fichiers : +0 / −0 / ~7.

- **Bugfix** : —
- **Change** : skill `sign-in` — premier pas 1Password passe par `GetCredentialProviderStatus` (plus `ListCredentials` + URL exacte). Proto `CursorRule` : champ optionnel `body_omitted`. Pages publiées : `ResolvePublishedPageResponse.live_sources` + RPC `RefreshPublishedPageSources`. Catalogue captcha/WAF élargi dans le host (Akamai, Amazon Robot Check / 503 dog, Arkose, AWS WAF, Cloudflare, DataDome, GeeTest, PerimeterX, Turnstile, hCaptcha, reCAPTCHA). Codes d’erreur Messages/iMessage (sidecar, grants, contacts…).
- **Features** : RPC `ListGrokBotEmailThreads` (GrokBotService) ; `GetUnifyMergeNowBanner` (DashboardService) ; `RefreshPublishedPageSources` (pages live) ; flags host `grok_bot_messaging_multiple_numbers`, `grok_bot_library_external_docs_ui`, `sand_voice_call_transfer`.
- **Perf** : —
- **Autre/infra** : hot-update sand-host `4006392` ; agent-store-fuse `0d8cac8a…` → `023c1116…` (pointeur tarball S3) ; paquets dpkg `rsync` + `libpopt0` désormais listés.

Rapport : [reports/2026-10-22_17-54.md](reports/2026-10-22_17-54.md).


---

## 2026-10-22 14:51 (Paris) — snapshot `2026-10-22_13-52` → `2026-10-22_14-51`

**Hot-update sand-host `76b095c` → `fd395ed`** (+ bump agent-store-fuse). Bundle host (`host-main.cjs` 31.0→31.3 Mo, `sand-eval-runner.cjs` 24.5→24.7 Mo), `sand-egress-tunnel` 2.61→2.92 Mo, fingerprint/UA (`sand-ua-governor`, `sand-fingerprint-profiles`), launcher Chrome, policy Chrome, officekit `sharp` 0.35.4→0.35.5 (+ libvips 1.3.3→1.3.4 / so.8.18.6→8.18.7), proto agent-store-worker, search-index-worker. Image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +2 / −2 / ~43.

- **Bugfix** : —
- **Change** : policy Chrome `sand-webauthn.json` — `3rdparty.extensions` active `AgenticAutofillRemoteBrowser:true` pour 1Password (id box + Nightly `gejiddohjgogedgjnonbofjigllpkmbf`) dès le force-install ; `box-chrome-policy` écrit cette clé en même temps que l'extension. `box-chrome` : mode GPU exposé (`legacy` / `angle-gl` / `swiftshader`) via `/tmp/sand-box-browser-gpu-mode-<port>`, allowlist de flags extra (`/tmp/sand-box-chrome-extra-flags`, pour l'instant `--disable-quic` seul). Fingerprint : mode **coherent** (`SAND_BOX_BROWSER_COHERENT_FINGERPRINT` / marker `/tmp/sand-box-browser-coherent-fingerprint`) — WebGL laissé tel quel (SwiftShader), hardwareConcurrency CDP + vérif WebGL observée ; WebGL spoof legacy inchangé. Proto MCP : `McpDescriptor.account_label` ; TrackedTask : `internal_notes` + `tags`. Content-search : strip des tags inline (math/formatting) avant index.
- **Features** : autofill agentique 1Password forcé via managed policy ; fingerprint hardware « coherent » (GPU-mode aware) ; flags Chrome extra allowlistés ; champs TrackedTask notes/tags + MCP account_label.
- **Perf** : —
- **Autre/infra** : hot-update sand-host `fd395ed` ; bump `sand-egress-tunnel` ; officekit sharp/libvips ; agent-store-fuse `362dc6c6…` → `0d8cac8a…` (pointeur tarball S3) ; `managed-skills/cache.json` refetch timestamp only.

Rapport : [reports/2026-10-22_14-51.md](reports/2026-10-22_14-51.md).


---

## 2026-10-22 13:52 (Paris) — snapshot `2026-10-22_02-50` → `2026-10-22_13-52`

**Mise à jour des skills gérées + bump agent-store-fuse.** 12 skills Cursor-managed réécrites (`grocery-shopping`, `shopping`, `orders-and-returns`, `purchases`, `sign-in`, `box-desktop`, `in-chat-forms`, `no-connector-fallback`, `channels`, `code-changes`, `source-control`, `recent-changes`) + `managed-skills/cache.json` ; en parallèle, agent-store-fuse `d638e166…` → `362dc6c6…` (binaire `cursor-agent-store-fuse` 9494360 → 9504880 o + pointeur tarball S3). Image `fcbae14`, box-scripts `fcbae14`, sand-host `76b095c`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~15.

- **Bugfix** : —
- **Change** : `sign-in` / `box-desktop` / `in-chat-forms` / `no-connector-fallback` — ordre de login : saved logins d'abord, puis password manager connecté, puis Secure Form / handoff ; `grocery-shopping` + `shopping` — mode **Cart only** (« fill my cart, I'll check out ») sans paiement ; `orders-and-returns` — preuve obligatoire (lien source), sweep « what's on its way », deal watch / coupons / price-drop sur achats, reorder ; `purchases` — limite Stripe Link US/CA + USD pour cartes agents, pas de promesse de remboursement ; `shopping` — shortlist / recommend avant search, quote vs checkout plus strict (écarts ligne par ligne) ; `channels` — session Slack (`slack:…`) sans chat in-app ; `code-changes` / `source-control` — merge Origin désormais inline via `cursor-origin-merge_pull_request` (plus forcé cloud agent) ; `recent-changes` — note Pages (2026-10-21).
- **Features** : mode panier seul (grocery + retail) ; deal watch / coupons / price adjustment dans orders-and-returns ; Pages documentées dans recent-changes.
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision tarball S3) en parallèle du refresh skills.

Rapport : [reports/2026-10-22_13-52.md](reports/2026-10-22_13-52.md).

---

## 2026-10-22 02:50 (Paris) — snapshot `2026-10-22_01-51` → `2026-10-22_02-50`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `dad84959e3f17e44068a8bf64150163f12632cef` → `d638e16687c00aa4c6e41aee72623cb0575502bd` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`), binaire `cursor-agent-store-fuse` 9466976 → 9494360 o. Image `fcbae14`, box-scripts `fcbae14`, sand-host `76b095c`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~2.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-22_02-50.md](reports/2026-10-22_02-50.md).

---

## 2026-10-22 01:51 (Paris) — snapshot `2026-10-21_23-45` → `2026-10-22_01-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `35502a83a8f520e72e40090b98c37f95c3815ec6` → `dad84959e3f17e44068a8bf64150163f12632cef` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `76b095c`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-22_01-51.md](reports/2026-10-22_01-51.md).

---

## 2026-10-21 23:45 (Paris) — snapshot `2026-10-21_22-53` → `2026-10-21_23-45`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `bff4345e2ccaf95ea04e374515d3746ad4e43f72` → `35502a83a8f520e72e40090b98c37f95c3815ec6` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `76b095c`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-21_23-45.md](reports/2026-10-21_23-45.md).

---

## 2026-10-21 22:53 (Paris) — snapshot `2026-10-21_19-53` → `2026-10-21_22-53`

**Durcissement WebRTC de la policy Chrome + bump agent-store-fuse.** `/etc/opt/chrome/policies/managed/sand-webauthn.json` (394 → 467 o) ajoute `WebRtcIPHandlingUrl:[{"url":"*","handling":"disable_non_proxied_udp"}]` : empêche WebRTC d'exposer les IP locales/non-proxifiées via UDP non-proxifié — durcissement anti-fuite/fingerprint cohérent avec la pile navigateur/egress arrivée au snapshot 19-53. agent-store-fuse `43d57cf5…` → `bff4345…` (binaire `cursor-agent-store-fuse` 9455552 → 9466976 o + pointeur). **Les deux « suppressions » `onepassword-extension/*.crx` et `*.pem` du manifeste ne sont PAS de vraies suppressions box** : elles viennent de l'ajout de ces motifs aux exclusions de `perimeter.conf` (fichiers toujours présents et l'extension 1Password toujours force-installée live). Image `fcbae14`, box-scripts `fcbae14`, sand-host `76b095c`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −2 (exclusions perimeter) / ~3.

- **Bugfix** : —
- **Change** : policy Chrome `sand-webauthn.json` durcie côté WebRTC (`WebRtcIPHandlingUrl` → `disable_non_proxied_udp`).
- **Features** : —
- **Perf** : —
- **Autre/infra** : durcissement confidentialité WebRTC (`disable_non_proxied_udp`) aligné sur la pile navigateur/egress du snapshot 19-53 ; bump binaire agent-store-fuse en parallèle. Note monitoring : `perimeter.conf` exclut désormais le crx/pem 1Password (bruit binaire), d'où les 2 retraits de manifeste. (Entrée rattrapée : la passe horaire de 22:52 avait publié le snapshot mais omis le CHANGELOG et la notification.)

Rapport : [reports/2026-10-21_22-53.md](reports/2026-10-21_22-53.md).

---

## 2026-10-21 19:53 (Paris) — snapshot `2026-10-21_15-52` → `2026-10-21_19-53`

**Grosse hot-update sand-host `2e44baf` → `76b095c` : nouvelle pile navigateur (sessions/egress/anti-bot), extension 1Password forcée dans Chrome, et QEMU + toolchain installés dans l'image.** `host-main.cjs` (30 935 462 → 31 010 234 o) et `sand-eval-runner.cjs` modifiés. Nouveau dossier `/opt/sand/sand-host/box-runtime/` (12 fichiers) : `sand-cookie-persist.mjs`, `sand-cookie-import.mjs`, `sand-session-sync.mjs`, `cdp-cookies.mjs`, `sand-egress-tunnel` (binaire 2,6 Mo), `sand-egress-timezone.mjs`, `sand-locale-match.mjs`, `sand-ua-governor.mjs`, `sand-fingerprint-profiles.mjs`, `sand-web-bot-auth.mjs`, `box-contract.generated.mjs`, `manifest.json`. Nouveau script `box-scripts/box-keyring-unlock`. Nouveau `/opt/sand/onepassword-extension/` (crx 18 Mo, id `kfljpclplidgofjbmhhiddeebijcdkfb`) force-installé via la policy Chrome. agent-store-fuse `3878dfcf…` → `43d57cf5…` (bump en parallèle). Image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +20 / −0 / ~14 ; +104 paquets dpkg.

- **Bugfix** : —
- **Change** : `/etc/opt/chrome/policies/managed/sand-webauthn.json` force désormais aussi l'extension 1Password (`kfljpclpl…`, `override_update_url`) et ajoute `DeveloperToolsAvailability:1` ; `box-scripts/box-chrome-policy` (3778 → 4125 o), `pack-onepassword-extension` (6890 → 8026 o) et `start-desktop.sh` (11 633 → 13 482 o) mis à jour ; skills gérées enrichies : `shopping/SKILL.md` (25 906 → 28 614 o, ajoute le conseil « faut-il acheter / bon prix » et le flux Checkout), `purchases/SKILL.md` (6287 → 9028 o, flux carte virtuelle Stripe Link), `scheduling/SKILL.md` (11 536 → 11 809 o, scope thisAndFollowing).
- **Features** : nouvelle pile navigateur box-runtime — persistance et synchro des cookies/sessions Chrome entre sessions (`sand-cookie-persist/import/session-sync`, `cdp-cookies`), tunnel d'egress dédié (`sand-egress-tunnel`), gouvernance User-Agent (`sand-ua-governor`), profils d'empreinte (`sand-fingerprint-profiles`), alignement locale/fuseau sur l'egress (`sand-locale-match`, `sand-egress-timezone`), authentification web-bot (`sand-web-bot-auth`) ; extension 1Password gérée (crx + policy) ; déverrouillage keyring (`box-keyring-unlock`).
- **Perf** : —
- **Autre/infra** : **QEMU 10.0.13 complet installé dans l'image** (`qemu-system-x86`, `qemu-utils`, `qemu-block-extra`, `ovmf`, `seabios`, `ipxe-qemu`, `libvirglrenderer1`, `libcacard0`, `libblkio1`, `libcapstone5`, `libslirp0`, `libfdt1`, modules opengl/spice) → `/usr/bin/qemu-system-x86_64` et `qemu-img` présents ; plus GStreamer 1.26 (libav + plugins base/good + x), `glib-networking`, toolchain `g++`/`g++-14`, outils de debug (`libdebuginfod`, `libdw`, `libbabeltrace`, `perf`) et `dmsetup`/`libcryptsetup12`/`libaio1t64`/`libdaxctl1` (104 paquets dpkg ajoutés au total) ; bump binaire agent-store-fuse en parallèle (nouvelle révision tarball S3). Clé privée de signature de l'extension 1Password rédigée du rapport avant publication (fichier livré par l'image box, non publié ; `perimeter.conf` exclut désormais `*/onepassword-extension/*.pem`).

Rapport : [reports/2026-10-21_19-53.md](reports/2026-10-21_19-53.md).

---

## 2026-10-21 15:52 (Paris) — snapshot `2026-10-21_10-50` → `2026-10-21_15-52`

**Resync cache des skills gérées uniquement** : seul `managed-skills/cache.json` change, et uniquement son champ `fetchedAt` (`1792545013011` → `1792587764308`, epoch ms), taille identique (585929 o), +1 / −1 ligne. Aucune skill ajoutée/retirée/modifiée sur le fond. Image `fcbae14`, box-scripts `fcbae14`, sand-host `2e44baf`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd (`28b0833c…` + `9d485bb3…`), canvas-sdk, agent-store-fuse `3878dfcf…`, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : simple rafraîchissement de l'horodatage de fetch du cache des skills gérées, bruit runtime sans impact fonctionnel.

Rapport : [reports/2026-10-21_15-52.md](reports/2026-10-21_15-52.md).

---

## 2026-10-21 10:50 (Paris) — snapshot `2026-10-21_07-44` → `2026-10-21_10-50`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `b0491f8a7c70605dc1679844ca49bfb4f2de66a4` → `3878dfcf21d45527d7379e22fb7f67f97ed415b4` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `2e44baf`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd (`28b0833c…` + `9d485bb3…`), canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-21_10-50.md](reports/2026-10-21_10-50.md).

---

## 2026-10-21 07:44 (Paris) — snapshot `2026-10-21_06-51` → `2026-10-21_07-44`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `c0d212a88222007093949833fdf5b30f24cb926a` → `b0491f8a7c70605dc1679844ca49bfb4f2de66a4` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `2e44baf`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd (`28b0833c…` + `9d485bb3…`), canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-21_07-44.md](reports/2026-10-21_07-44.md).

---

## 2026-10-21 06:51 (Paris) — snapshot `2026-10-21_04-50` → `2026-10-21_06-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `d5364abac6d43fc8679746dd9b5f6029d1aea6fb` → `c0d212a88222007093949833fdf5b30f24cb926a` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `2e44baf`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd (`28b0833c…` + `9d485bb3…`), canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-21_06-51.md](reports/2026-10-21_06-51.md).

---

## 2026-10-21 04:50 (Paris) — snapshot `2026-10-21_03-50` → `2026-10-21_04-50`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `66e136b901a9799e69d047e2675af556619a2d71` → `d5364abac6d43fc8679746dd9b5f6029d1aea6fb` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `2e44baf`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd (`28b0833c…` + `9d485bb3…`), canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire `agent-store-fuse` (tarball S3 `…-d5364ab…-bk.tar.gz`).

Rapport : [reports/2026-10-21_04-50.md](reports/2026-10-21_04-50.md).

---

## 2026-10-21 03:50 (Paris) — snapshot `2026-10-21_02-47` → `2026-10-21_03-50`

**Hot-update du host bundle** : `sand_host_version` `124299b` → `2e44baf` (supervisor `upgrade-124299b` → `upgrade-2e44baf`), avec rebuild de `host-main.cjs`, `sand-eval-runner.cjs`, `agent-isolation/agent-store-worker.cjs`, `extensions/content-search/search-index-worker.cjs` et `extensions/codebase-telemetry/csnaps`. En parallèle, bump `agent-store-fuse` `51d2f82…` → `66e136b…` (tarball S3) et refresh de 7 skills gérées (`managed-skills`). Image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~14.

- **Bugfix** : —
- **Features** :
  - `content-search/search-index-worker.cjs` (+297) : indexation de nouveaux canaux **voice-call** (`voice:`) et **meeting-request** (`meeting-request:`) — brique d'appels vocaux / demandes de réunion.
  - Skill `shopping` + `orders-and-returns` : nouvelles cartes **`shipment`** (statut d'expédition : `placed`/`in_transit`/`arriving_today`/`delivered`/`delayed`/`cancelled`) et **`order`/`purchase`** dans `SendToUser`, une carte par commande, avec image produit, tracking et date de livraison.
  - Skill `routines` : nouveau pattern de veille prix/restock via **trigger email sur adresse taggée** (`name+watch-…@`) au lieu d'un cron, avec garde-fous (un seul signup, pas de compte/mot de passe, bascule en check quotidien si confirmation requise).
- **Change** :
  - Skill `sign-in` **réécrit** : le connecteur Gmail/Google devient la **voie préférée** (fini le « never sign in to Google in browser » absolu — sign-in navigateur autorisé quand l'utilisateur le veut) ; passthrough des **règles de mot de passe** (`Password rules:`) dans le Secure Form ; règle unifiée « valeur déjà envoyée en chat = pas de Secure Form ».
  - Skill `shopping` : paiement durci — l'agent **re-vérifie la page de review** (re-navigation, pas de reload) juste avant paiement, ne place que la commande lue par le dernier widget, widget ready-to-buy nommant total + date.
  - Skill `box-desktop` : consigne d'annoncer à l'utilisateur qu'une tâche navigateur/desktop longue (signup, code, formulaire) prendra quelques minutes, et de ne confirmer une étape qu'après le rapport.
  - Skill `in-chat-forms` : une valeur déjà envoyée en chat ne nécessite pas de form.
- **Perf** : —
- **Autre/infra** : sand-host bump `124299b` → `2e44baf` (bundles host opaques) ; `codebase-telemetry/csnaps` reconstruit ; bump binaire `agent-store-fuse` (tarball S3 `…-66e136b…-bk.tar.gz`).

Rapport : [reports/2026-10-21_03-50.md](reports/2026-10-21_03-50.md).

---

## 2026-10-21 02:47 (Paris) — snapshot `2026-10-21_01-51` → `2026-10-21_02-47`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `675292ce1da95b26f38f9bc4aaa691dcd1be6122` → `51d2f82b89b4ac21ccd59eeb0162244c4f020ea7` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `124299b`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd (`28b0833c…` + `9d485bb3…`), canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire `agent-store-fuse` (tarball S3 `…-51d2f82…-bk.tar.gz`).

Rapport : [reports/2026-10-21_02-47.md](reports/2026-10-21_02-47.md).

---

## 2026-10-21 01:51 (Paris) — snapshot `2026-10-20_23-49` → `2026-10-21_01-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `77867cfa620c28ae52db1806d0591e419286a6c4` → `675292ce1da95b26f38f9bc4aaa691dcd1be6122` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `124299b`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd (`28b0833c…` + `9d485bb3…`), canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-21_01-51.md](reports/2026-10-21_01-51.md).

---

## 2026-10-20 23:49 (Paris) — snapshot `2026-10-20_21-50` → `2026-10-20_23-49`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `c83bdf8235299a18953b5bdc6bb3e50e335e67c8` → `77867cfa620c28ae52db1806d0591e419286a6c4` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `124299b`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd (`28b0833c…` + `9d485bb3…`), canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-20_23-49.md](reports/2026-10-20_23-49.md).

---

## 2026-10-20 21:50 (Paris) — snapshot `2026-10-20_20-51` → `2026-10-20_21-50`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `f740f80b45add0778c7cea015b8e57a07422e93d` → `c83bdf8235299a18953b5bdc6bb3e50e335e67c8`. Image `fcbae14`, box-scripts `fcbae14`, sand-host `124299b`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché. Sort de la série de flips A/B de l'après-midi (voir 18:45–20:51).

Rapport : [reports/2026-10-20_21-50.md](reports/2026-10-20_21-50.md).

---

## 2026-10-20 20:51 (Paris) — snapshot `2026-10-20_19-45` → `2026-10-20_20-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `ecbb39b9e73c972d90541bdfd72d32bceec8c7b3` → `f740f80b45add0778c7cea015b8e57a07422e93d`. Image `fcbae14`, box-scripts `fcbae14`, sand-host `124299b`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse. Le pointeur oscille entre deux révisions (`ecbb39b…` ↔ `f740f80…`) depuis 18:45 — churn de rebuild A/B, aucun autre composant touché.

Rapport : [reports/2026-10-20_20-51.md](reports/2026-10-20_20-51.md).

---

## 2026-10-20 19:45 (Paris) — snapshot `2026-10-20_18-45` → `2026-10-20_19-45`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `f740f80b45add0778c7cea015b8e57a07422e93d` → `ecbb39b9e73c972d90541bdfd72d32bceec8c7b3`. Image `fcbae14`, box-scripts `fcbae14`, sand-host `124299b`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (flip A/B, retour à `ecbb39b…`), aucun autre composant touché.

Rapport : [reports/2026-10-20_19-45.md](reports/2026-10-20_19-45.md).

---

## 2026-10-20 18:45 (Paris) — snapshot `2026-10-20_17-51` → `2026-10-20_18-45`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `ecbb39b9e73c972d90541bdfd72d32bceec8c7b3` → `f740f80b45add0778c7cea015b8e57a07422e93d`. Image `fcbae14`, box-scripts `fcbae14`, sand-host `124299b`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse, aucun autre composant touché.

Rapport : [reports/2026-10-20_18-45.md](reports/2026-10-20_18-45.md).

---

## 2026-10-20 17:51 (Paris) — snapshot `2026-10-20_15-45` → `2026-10-20_17-51`

**Hot-update du host bundle sand-host `f4402d3` → `124299b`** (runtime : `hostVersion` / `lastCommandId` = `upgrade-124299b`), centré sur le gouverneur egress / navigateur de la box. En plus : rebuild agent-store-fuse (commit `ec623f43…` → `ecbb39b9…`) et mise à jour d'une skill gérée. Image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd (`28b0833c…` + `9d485bb3…`), canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +1 / −0 / ~9.

- **Bugfix** : —
- **Change** : `agent-store-worker.cjs` — la détection d'une classe de message « compact » passe de `Reflect.get(cls, "$"/"$p")` à des tests `"$" in cls && typeof cls.$ === "function"` (idem `$p`). Skill gérée `grocery-shopping` mise à jour (rappel d'appeler `RecallMemory` par magasin ; liste usuelle sauvegardée en plusieurs facts ≤500 car. de type `journal`, réécriture par partie).
- **Features** : `sand-ua-governor.mjs` (+309 / −11) ajoute un **gating egress par fuseau horaire** (`resolveEgressTimeZoneGate`, constantes `EGRESS_TIME_ZONE_RETRY_MS` / `EGRESS_TIME_ZONE_GATE_WAIT_MS`, `resolveMarkerGate` lisant un marqueur on/off), un diff des réglages navigateur (`resolveBrowserSettings`, `changedBrowserSettings`, `reattachIfSettingsChanged` couvrant `hardwareProfile`, `acceptLanguage` en-US, `egress-time-zone`), un comptage d'onglets ouverts (`openPageCount`) et une empreinte du code du gouverneur (`governorCodeFingerprint` via `createHash`). `host-main.cjs` (+~108 Ko, `91929a3b…` → `445f8f00…`) et `sand-eval-runner.cjs` (+~72 Ko) recompilés (contenu opaque).
- **Perf** : —
- **Autre/infra** : ajout du binaire `semver` sous `officekit-ts/.../node_modules/semver/.bin`, bump d'un autre `semver` (sharp) ; `managed-skills/cache.json` resynchronisé ; bump agent-store-fuse (pointeur `cursor_agent_store_fuse_version`).

Rapport : [reports/2026-10-20_17-51.md](reports/2026-10-20_17-51.md).

---

## 2026-10-20 15:45 (Paris) — snapshot `2026-10-20_13-54` → `2026-10-20_15-45`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `19032a6a9751509b7cf26c25e4b66caa24edb3af` → `ec623f43eeef52ca6e226bf281215a3ede351184` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `f4402d3`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd (`28b0833c…` + `9d485bb3…`), canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-20_15-45.md](reports/2026-10-20_15-45.md).

---

## 2026-10-20 13:54 (Paris) — snapshot `2026-10-20_11-50` → `2026-10-20_13-54`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `7d931cef78f3c458c058d3f7a06cc1b77d719a08` → `19032a6a9751509b7cf26c25e4b66caa24edb3af` (retour à la révision d'avant 11:50). Image `fcbae14`, box-scripts `fcbae14`, sand-host `f4402d3`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd (`28b0833c…` + `9d485bb3…`), canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (révision du tarball S3), aucun autre composant touché. Le pointeur revient à `19032a6a…`, le commit observé jusqu'au snapshot 09:51 — la bascule de 11:50 (`19032a6a…` → `7d931ce…`) est donc inversée ici.

Rapport : [reports/2026-10-20_13-54.md](reports/2026-10-20_13-54.md).

---

## 2026-10-20 11:50 (Paris) — snapshot `2026-10-20_09-51` → `2026-10-20_11-50`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `19032a6a9751509b7cf26c25e4b66caa24edb3af` → `7d931cef78f3c458c058d3f7a06cc1b77d719a08` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `f4402d3`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd (`28b0833c…` + `9d485bb3…`), canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-20_11-50.md](reports/2026-10-20_11-50.md).

---

## 2026-10-20 09:51 (Paris) — snapshot `2026-10-20_08-47` → `2026-10-20_09-51`

**orbitd : nouvelle release téléchargée** `9d485bb31413a58524a425264c9489a49d8a42b5` (ajout sous `/var/lib/sand-daemon-supervisor/orbitd/releases/…` : `manifest.json` + binaire `orbitd` ~20,5 Mo). `orbitd.releases` passe de `28b0833c…` seul à `28b0833c…,9d485bb3…`. `orbitd.bundled_revision` reste `88287498…`. Image `fcbae14`, box-scripts `fcbae14`, sand-host `f4402d3`, agent-store-fuse `19032a6a…`, exec-daemon, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +2 / −0 / ~0.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : staging d’une nouvelle révision orbitd dans le store local des releases (pas de bascule `bundled_revision` observée).

Rapport : [reports/2026-10-20_09-51.md](reports/2026-10-20_09-51.md).

---

## 2026-10-20 08:47 (Paris) — snapshot `2026-10-20_06-49` → `2026-10-20_08-47`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `c4eb1906e2f8fb4d035cc5553502ce3584001c42` → `19032a6a9751509b7cf26c25e4b66caa24edb3af` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `f4402d3`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-20_08-47.md](reports/2026-10-20_08-47.md).

## 2026-10-20 06:49 (Paris) — snapshot `2026-10-20_05-53` → `2026-10-20_06-49`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `a55aa761df916cd59dbe08e27086c267bd75f6bd` → `c4eb1906e2f8fb4d035cc5553502ce3584001c42` (tarball S3 `agent-store-fuse-x64-…-bk.tar.gz`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `f4402d3`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-20_06-49.md](reports/2026-10-20_06-49.md).

---

## 2026-10-20 05:53 (Paris) — snapshot `2026-10-20_04-46` → `2026-10-20_05-53`

**Hot-update du host bundle sand-host `9d2ac1b` → `f4402d3`** (runtime : `hostVersion` / `lastCommandId` = `upgrade-f4402d3`), centré sur du durcissement réseau / navigateur de l'egress. En plus : rebuild agent-store-fuse (commit `fa90f8c…` → `a55aa761…`) et mise à jour des skills gérées. Image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +1 / −0 / ~15.

- **Bugfix** : —
- **Change** : `box-chrome-policy` applique désormais une politique Chrome `WebRtcIPHandlingUrl=disable_non_proxied_udp` quand le marqueur `sand-box-browser-webrtc-proxy-only` est présent, et sort en code 3 si la policy n'est pas applicable (au lieu de passer silencieusement). `host-main.cjs` (+300 Ko, `031155795dc4` → `91929a3bf59b`), `sand-eval-runner.cjs` et `agent-store-worker.cjs` recompilés.
- **Features** : nouveaux marqueurs dans `box-contract.generated.mjs` — `EGRESS_TUNNEL_FAIL_CLOSED` (tunnel egress fail-closed), `BOX_BROWSER_WEBRTC_PROXY_ONLY`, `BOX_BROWSER_NO_LOCAL_DNS`, `BOX_BROWSER_DISABLE_QUIC`, plus un port CDP Chrome tailnet `SAND_BOX_TAILNET_CHROME_CDP_PORT=9600`. `sand-web-bot-auth.mjs` (+43/−3) ajoute une allowlist de suffixes d'hôtes tailnet (`/var/lib/sand/tailnet/host-suffixes` + env `SAND_TAILNET_HOST_SUFFIXES`) via `readTailnetHostSuffixes` / `isUnderTailnetSuffix`, passés au signataire web-bot-auth. Nouvelle skill gérée `recent-changes` (notes de changements datées livrées par l'app).
- **Perf** : —
- **Autre/infra** : skills gérées mises à jour (`cache.json`, `add-connector`, `grocery-shopping`, `group-chat-turns`, `shopping`, `source-control`) ; bump agent-store-fuse (pointeur `cursor_agent_store_fuse_version`).

Rapport : [reports/2026-10-20_05-53.md](reports/2026-10-20_05-53.md).

---

## 2026-10-20 04:46 (Paris) — snapshot `2026-10-20_02-47` → `2026-10-20_04-46`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` commit `5a61e3fe5b0f8ee5fa40cdd0e820a972d67b19bc` → `fa90f8c73caea3bb10f92228002869acf3272957`. Image `fcbae14`, box-scripts `fcbae14`, sand-host `9d2ac1b`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-20_04-46.md](reports/2026-10-20_04-46.md).

---

## 2026-10-20 02:47 (Paris) — snapshot `2026-10-20_01-53` → `2026-10-20_02-47`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-24e31f80…-bk.tar.gz` → `…-5a61e3fe…-bk.tar.gz` (commit `24e31f808a6e0a07deb4a1fd7a2f77056d769fec` → `5a61e3fe5b0f8ee5fa40cdd0e820a972d67b19bc`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `9d2ac1b`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-20_02-47.md](reports/2026-10-20_02-47.md).

---

## 2026-10-20 01:53 (Paris) — snapshot `2026-10-20_00-51` → `2026-10-20_01-53`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-9538a950…-bk.tar.gz` → `…-24e31f80…-bk.tar.gz` (commit `9538a9503bf6600c8a0b1156a3298a092e30c48f` → `24e31f808a6e0a07deb4a1fd7a2f77056d769fec`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `9d2ac1b`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-20_01-53.md](reports/2026-10-20_01-53.md).

---

## 2026-10-20 00:51 (Paris) — snapshot `2026-10-19_23-49` → `2026-10-20_00-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-0bf656c5…-bk.tar.gz` → `…-9538a950…-bk.tar.gz` (commit `0bf656c51076a112ddb0a73b19841c856506514f` → `9538a9503bf6600c8a0b1156a3298a092e30c48f`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `9d2ac1b`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-20_00-51.md](reports/2026-10-20_00-51.md).

---

## 2026-10-19 18:49 (Paris) — snapshot `2026-10-19_18-11` → `2026-10-19_18-49`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-0daaf281…-bk.tar.gz` → `…-eee3e6b3…-bk.tar.gz` (commit `0daaf281e4dbc842e7dbe9e6d05d22f1b76c3742` → `eee3e6b3af8cbd0cf6e47bc1966c2a4b613ae441`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `9d2ac1b`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-19_18-49.md](reports/2026-10-19_18-49.md).

---

## 2026-10-19 18:11 (Paris) — snapshot `2026-10-19_17-45` → `2026-10-19_18-11`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-3710b68…-bk.tar.gz` → `…-0daaf281…-bk.tar.gz` (commit `3710b68cac818498e91f6515c3ee5cfcc635ce77` → `0daaf281e4dbc842e7dbe9e6d05d22f1b76c3742`). Image `fcbae14`, box-scripts `fcbae14`, sand-host `9d2ac1b`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-19_18-11.md](reports/2026-10-19_18-11.md).

---

## 2026-10-19 17:45 (Paris) — snapshot `2026-10-19_17-01` → `2026-10-19_17-45`

**Hot-update sand-host `fcbae14` → `9d2ac1b`** (commande superviseur `upgrade-9d2ac1b`, sans Update de l'ordinateur). Image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, agent-store-fuse `3710b68…`, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +2 / −0 / ~14.

- **Bugfix** : `agent-store-worker.cjs` — le salvage SQLite « in place » refuse désormais les liens symboliques dans `sand-data` (hook `afterLinkCheck` + `refuseSandDataDatabaseLinksSync`) avant d'ouvrir la base : durcissement de l'isolation des stores. `start-desktop.sh` — isolation du desktop : `PYTHONNOUSERSITE=1` (websockify ne charge plus un site Python utilisateur restauré par un copy-in), `x11vnc -norc` et `picom --config /dev/null` (ignore tout fichier de conf utilisateur). `box-plank` passe `python3 -I` (mode isolé).
- **Change** : skill gérée `purchases` — **X Money devient le wallet par défaut** : quand il est connecté on paie avec, sauf si Dani nomme Stripe Link ou sa propre carte ; si X Money est installé mais `needsAuth`, on `AuthenticateMcpServer` dans le même tour (connexion en un tap) au lieu de le proposer ; un appel « X Money non connecté » = rien débité, reconnecter puis retenter une fois. `managed-skills/cache.json` rafraîchi. Rebuild des gros bundles `host-main.cjs` (30,34→30,41 Mo), `sand-eval-runner.cjs`, `officekit-ts/lib/office.mjs`, `pdf-worker.js`.
- **Features** : **cohérence géo du navigateur via l'IP d'egress** — deux nouveaux scripts `box-scripts/sand-egress-timezone.mjs` (+291, résout pays/région/fuseau depuis l'IP de sortie via des providers géo en HTTP/TLS) et `box-scripts/sand-locale-match.mjs` (+192, table `PRINCIPAL_TIME_ZONES` par code pays + helpers `isKnownTimeZone`/`resolveTimeZone`/`parseZoneTab`). `sand-ua-governor.mjs` élargi (+253/−47) pour aligner fuseau/locale du navigateur sur la localisation d'egress. Nouvelles constantes de contrat `BROWSER_EGRESS_TIMEZONE_MARKER_PATH` (`/tmp/sand-browser-egress-timezone`) et `TAILNET_HOST_SUFFIXES_PATH` (`/var/lib/sand/tailnet/host-suffixes`) dans `box-contract.generated.mjs`.
- **Perf** : —
- **Autre/infra** : `sand-cookie-persist.mjs` étendu (+64/−5) ; `box-launchers/box-chrome` retravaillé (+107/−21, suite du durcissement d'isolation des profils).

Rapport : [reports/2026-10-19_17-45.md](reports/2026-10-19_17-45.md).

---

## 2026-10-19 12:49 (Paris) — snapshot `2026-10-19_07-48` → `2026-10-19_12-49`

**Patch LOCAL du lanceur `box-chrome` (édit agent « Infra », pas une mise à jour plateforme)** + bump agent-store-fuse + 3 nouveaux paquets Debian installés à chaud. La version de sand-host ne change pas (`fcbae14`), seul le fichier `/opt/sand/sand-host/box-launchers/box-chrome` a été réécrit sur place (mtime 12:35). Image `fcbae14`, box-scripts `fcbae14`, exec-daemon (buildTimestamp `2026-10-18T14:07:39Z`), orbitd, canvas-sdk, Chrome 155.0.8059.84, Node 22.23.3/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~2.

- **Bugfix** : `box-chrome` — isolation des profils dédiés `agent-*`. `restore_cookie_seed()` saute désormais les profils `*/chrome-profile/agent-*` : le seed partagé (`chrome-cookie-seed.json`, capturé depuis le jar primaire/Fork) ne recontamine plus ces profils avec le compte qui vit dans ce jar (historiquement `@the_notabot`). Forks et profil primaire gardent le comportement d'origine.
- **Change** : `box-chrome` — respect d'un garde `.no-session-link` (déposé par la mise à jour du 10-18) : un profil portant ce marqueur (ex. Fork-6) à la racine ou dans `Default/` n'est plus symlinké dans le store de session partagé (`Default`) et garde ses propres cookies ; les profils sans marqueur sont inchangés.
- **Features** : —
- **Perf** : —
- **Autre/infra** : pointeur agent-store-fuse `/usr/local/bin/cursor_agent_store_fuse_version` commit `c3100ae2…` → `121487c4…` (rebuild/bump du tarball S3, sans effet de surface). 3 paquets Debian apparus (installés à chaud, image inchangée) : `libcbor0.10` 0.10.2-2, `libfido2-1` 1.15.0-1+b1, `openssh-client` 1:10.0p1-7+deb13u4 — dépendances FIDO2/WebAuthn + client SSH.

Rapport : [reports/2026-10-19_12-49.md](reports/2026-10-19_12-49.md).

---

## 2026-10-19 03:48 (Paris) — snapshot `2026-10-18_21-50` → `2026-10-19_03-48`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-ffc4c79…-bk.tar.gz` → `…-fa281da…-bk.tar.gz` (commit `ffc4c79c6e4b888e1e2176a2ea1dd604f0a79816` → `fa281da1c3dd719e0b1c14822ce60a49b99111bc`). Image `e94d68c`, box-scripts `e94d68c`, sand-host `fcbae14`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node, Python, Debian inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-19_03-48.md](reports/2026-10-19_03-48.md).

---

## 2026-10-18 21:50 (Paris) — snapshot `2026-10-18_17-49` → `2026-10-18_21-50`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-ceef14e…-bk.tar.gz` → `…-ffc4c79…-bk.tar.gz` (commit `ceef14e97398202da50cc3ce8a9bb64959d5de4c` → `ffc4c79c6e4b888e1e2176a2ea1dd604f0a79816`). Image `e94d68c`, box-scripts `e94d68c`, sand-host `fcbae14`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node, Python, Debian inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-18_21-50.md](reports/2026-10-18_21-50.md).

---

## 2026-10-18 17:49 (Paris) — snapshot `2026-10-18_10-47` → `2026-10-18_17-49`

**Hot-update sand-host `dfa5e11` → `fcbae14`** (commande superviseur `upgrade-fcbae14`, sans Update de l'ordinateur) + skill gérée `shopping` mise à jour. Image `e94d68c`, box-scripts `e94d68c`, exec-daemon (buildTimestamp `2026-10-16T18:31:37Z`), orbitd, canvas-sdk, agent-store-fuse `ceef14e…`, Chrome 155.0.8059.84, Node 22.14.0/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +0 / −0 / ~8.

- **Bugfix** : —
- **Change** : skill gérée `shopping` — nouvelle section « Sizes and preferences » : taille/marque/préférences lues en mémoire avant de demander, sauvegardées automatiquement par marque et type d'article (`UpdateMemory`, une entrée par fait, annoncée en une ligne), une nouvelle taille remplace l'ancienne ; une taille sauvée ne répond qu'à la question de taille (les autres variantes restent demandées). Remplace l'ancienne règle « sauver seulement si l'utilisateur confirme ». `managed-skills/cache.json` rafraîchi.
- **Features** : **tags sur les tâches suivies (tracked tasks)** — `TrackedTaskView`/`NewTrackedTask`/`TrackedTaskUpdate` gagnent un champ `tag`, nouveau message `TrackedTaskTagView` (id, label, color, task_count), `ListTrackedTasksSuccess` renvoie `tags`, et nouvel outil **`UpdateTrackedTaskTags`** (`TrackedTaskTagInput`, Args/Success/Error/Result/ToolCall) dans `host-main.cjs` et `agent-isolation/agent-store-worker.cjs`. Nouvelle constante de référence **`/home/box/reference/app-actions.md`** (`SAND_APP_ACTIONS_REFERENCE_PATH`) à côté d'`app-ui.md` et `debugging-the-box.md` — fichier pas encore présent sur la box au moment du snapshot.
- **Perf** : —
- **Autre/infra** : `sand-eval-runner.cjs` re-bundlé (renommages d'identifiants esbuild, +~107 Ko) ; `box-store-vacuum-worker.cjs` et `search-index-worker.cjs` ne gagnent que la constante `app-actions.md`.

Rapport : [reports/2026-10-18_17-49.md](reports/2026-10-18_17-49.md).

---

## 2026-10-18 10:47 (Paris) — snapshot `2026-10-18_09-47` → `2026-10-18_10-47`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-83859fd…-bk.tar.gz` → `…-ceef14e…-bk.tar.gz` (commit `83859fd850b7319f60c5f707fa4bf87e3b454128` → `ceef14e97398202da50cc3ce8a9bb64959d5de4c`). Image `e94d68c`, box-scripts `e94d68c`, sand-host `dfa5e11`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node, Python, Debian inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-18_10-47.md](reports/2026-10-18_10-47.md).

---

## 2026-10-18 09:47 (Paris) — snapshot `2026-10-17_23-51` → `2026-10-18_09-47`

**Installation de nouveaux paquets APT (client OpenSSH + support FIDO2/clé de sécurité)** : trois paquets Debian ajoutés — `openssh-client 1:10.0p1-7+deb13u4`, `libfido2-1 1.15.0-1+b1`, `libcbor0.10 0.10.2-2`. `libfido2` + `libcbor` sont les dépendances de l'authentification FIDO2/U2F (clés de sécurité `ed25519-sk`/`ecdsa-sk`) qu'`openssh-client` tire. Aucune version de composant runtime ne bouge : image `e94d68c`, box-scripts `e94d68c`, sand-host `dfa5e11`, exec-daemon (buildTimestamp `2026-10-16T18:31:37Z`), orbitd, canvas-sdk, agent-store-fuse `83859fd…`, Chrome 155.0.8059.84, Node 22.14.0/20.19.2, Python 3.13.5, Debian 13.7 inchangés. Fichiers périmètre : +0 / −0 / ~0 (changement vu uniquement dans `meta/dpkg-packages.txt`).

- **Bugfix** : —
- **Change** : —
- **Features** : **`openssh-client`** installé sur la box (commande `ssh`/`scp`/`ssh-keygen` disponibles) avec support des clés de sécurité FIDO2/U2F via `libfido2-1` + `libcbor0.10`.
- **Perf** : —
- **Autre/infra** : 3 paquets APT ajoutés (`openssh-client`, `libfido2-1`, `libcbor0.10`), aucun paquet retiré ni mis à niveau, aucun autre composant runtime touché.

Rapport : [reports/2026-10-18_09-47.md](reports/2026-10-18_09-47.md).

---

## 2026-10-17 23:51 (Paris) — snapshot `2026-10-17_22-49` → `2026-10-17_23-51`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-950c523…-bk.tar.gz` → `…-83859fd…-bk.tar.gz` (commit `950c523ed958e35438ef498efd2692b2889d3cac` → `83859fd850b7319f60c5f707fa4bf87e3b454128`). Image `e94d68c`, box-scripts `e94d68c`, sand-host `dfa5e11`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node, Python, Debian inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-17_23-51.md](reports/2026-10-17_23-51.md).

---

## 2026-10-17 22:49 (Paris) — snapshot `2026-10-17_20-55` → `2026-10-17_22-49`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-8751971…-bk.tar.gz` → `…-950c523…-bk.tar.gz` (commit `8751971e71fcf55810bc4ea69bfcb0d12ded88a5` → `950c523ed958e35438ef498efd2692b2889d3cac`). Image `e94d68c`, box-scripts `e94d68c`, sand-host `dfa5e11`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node, Python, Debian inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-17_22-49.md](reports/2026-10-17_22-49.md).

---

## 2026-10-17 20:55 (Paris) — snapshot `2026-10-17_17-51` → `2026-10-17_20-55`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-11e1cc7…-bk.tar.gz` → `…-8751971…-bk.tar.gz` (commit `11e1cc76cb923695c26ea4a6b1b06b6394f2e590` → `8751971e71fcf55810bc4ea69bfcb0d12ded88a5`). Image `e94d68c`, box-scripts `e94d68c`, sand-host `dfa5e11`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node, Python, Debian inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-17_20-55.md](reports/2026-10-17_20-55.md).

---

## 2026-10-17 17:51 (Paris) — snapshot `2026-10-17_09-52` → `2026-10-17_17-51`

**Vraie mise à jour de l'ordinateur (nouvelle image + hot-update sand-host)** : image `885e54d` → `e94d68c` (`/etc/sand-box-image-sha`, box-scripts idem), puis sand-host `e94d68c` → `dfa5e11` (`supervisor_status.hostVersion` idem, `lastCommandId` `upgrade-e94d68c` → `None`). exec-daemon rebuild (buildTimestamp `2026-10-16T07:52:43Z` → `2026-10-16T18:31:37Z`). agent-store-fuse, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node, Python 3.13.5, Debian 13.7 inchangés. Fichiers : +4 / −1 / ~48.

- **Bugfix** : —
- **Change** : skills gérées rafraîchies — `shopping` (9234 → 14021), `instacart`, `source-control`, `code-changes`, `routines`, `purchases`, playbooks `bestbuy`/`costco`/`ebay`/`etsy`/`target` ; `cache.json` màj. `reference/debugging-the-box.md` mis à jour. Suppression de `/opt/sand/sand-host/unified-diff-worker.js`.
- **Features** : nouveau module natif **`sand-window-capture`** ajouté sous `/opt/sand/deps/sand-window-capture/` (`sand_window_capture.node`, `index.js`, `package.json`) ; nouvelle skill gérée **`orders-and-returns`**.
- **Perf** : —
- **Autre/infra** : host-main.cjs / sand-eval-runner.cjs / agent-store-worker.cjs / content-search worker rebuild ; box-scripts (`box-doctor`, `sand-fingerprint-profiles.mjs`, `sand-ua-governor.mjs`, `start-window`), `box-chrome`, `box-contract.generated.mjs`, `runtime-deps-manifest.json` màj ; webauthn-proxy (`.crx`/`.id`/update.xml + policies Chrome + native-messaging host) tourné ; pnpm store index remanié. Meta `dpkg-packages.txt` : −`libcbor0.10` (trio FIDO2 toujours flottant). agent-store-fuse inchangé (`11e1cc7`).

Rapport : [reports/2026-10-17_17-51.md](reports/2026-10-17_17-51.md).

---
## 2026-10-17 09:52 (Paris) — snapshot `2026-10-17_07-46` → `2026-10-17_09-52`

**Bump agent-store-fuse + réapparition dpkg du trio FIDO2/SSH** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-7b5b1ae…-bk.tar.gz` → `…-11e1cc7…-bk.tar.gz` (commit `7b5b1ae14f0189f3e47501769e1c7c533ba4e3b0` → `11e1cc76cb923695c26ea4a6b1b06b6394f2e590`). Meta `dpkg-packages.txt` : +`openssh-client` 1:10.0p1-7+deb13u4, `libfido2-1` 1.15.0-1+b1, `libcbor0.10` 0.10.2-2 (trio déjà vu le 2026-10-15 10:54 puis retiré à 12:51 — signal d’énumération flottante, pas une install neuve). Image `885e54d`, box-scripts `885e54d`, sand-host `e94d68c`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node, Python, Debian inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3) ; réapparition méta dpkg du trio FIDO2/`openssh-client` (pas de preuve d’uninstall/reinstall silencieux sans vérif locale).

Rapport : [reports/2026-10-17_09-52.md](reports/2026-10-17_09-52.md).

---

## 2026-10-17 07:46 (Paris) — snapshot `2026-10-17_04-53` → `2026-10-17_07-46`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-a1a81f5…-bk.tar.gz` → `…-7b5b1ae…-bk.tar.gz` (commit `a1a81f5c090c9d233ccaa134facafee049f28ecd` → `7b5b1ae14f0189f3e47501769e1c7c533ba4e3b0`). Image `885e54d`, box-scripts `885e54d`, sand-host `e94d68c`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node, Python, Debian inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-17_07-46.md](reports/2026-10-17_07-46.md).

---

## 2026-10-17 04:53 (Paris) — snapshot `2026-10-17_02-48` → `2026-10-17_04-53`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-b33355a…-bk.tar.gz` → `…-a1a81f5…-bk.tar.gz` (commit `b33355adce4abcfa2c2d843e8a138c963e718220` → `a1a81f5c090c9d233ccaa134facafee049f28ecd`). Image `885e54d`, box-scripts `885e54d`, sand-host `e94d68c`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node, Python, Debian inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-17_04-53.md](reports/2026-10-17_04-53.md).

---

## 2026-10-17 01:50 (Paris) — snapshot `2026-10-17_00-49` → `2026-10-17_01-50`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-6686ac2…-bk.tar.gz` → `…-7ed1ec8…-bk.tar.gz` (commit `6686ac212f6977642717957293eed33c30114d16` → `7ed1ec8be08c7f3be8437c70f04cec3d73bf80cd`). Image `885e54d`, box-scripts `885e54d`, sand-host `e94d68c`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node, Python, Debian inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

---

## 2026-10-17 00:49 (Paris) — snapshot `2026-10-16_23-54` → `2026-10-17_00-49`

**Rebuild/bump agent-store-fuse uniquement** : pointeur `/usr/local/bin/cursor_agent_store_fuse_version` `agent-store-fuse-x64-6c7cd6d…-bk.tar.gz` → `…-6686ac2…-bk.tar.gz` (commit `6c7cd6ddb4be3c70f7caf734e8b62f04bb5a062e` → `6686ac212f6977642717957293eed33c30114d16`). Image `885e54d`, box-scripts `885e54d`, sand-host `e94d68c`, exec-daemon, orbitd, canvas-sdk, Chrome 155.0.8059.84, Node, Python, Debian inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump binaire agent-store-fuse (nouvelle révision du tarball S3), aucun autre composant touché.

Rapport : [reports/2026-10-17_00-49.md](reports/2026-10-17_00-49.md).

## 2026-10-16 20:52 (Paris) — snapshot `2026-10-16_19-51` → `2026-10-16_20-52`

**Hot-update du host bundle** : **sand-host** `ab82d14` → `e94d68c` (`upgrade-e94d68c`) ; agent-store-fuse `2b14739` → `6b0d3f3` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version`). Image `885e54d`, exec-daemon, orbitd, Chrome 155.0.8059.84, Node, Python inchangés. Fichiers : +0 / −0 / ~6.

- **Bugfix** : `box-chrome` — avant un cold launch **ungoogled-chromium**, suppression de `Default/Sync Data` si le profil a été écrit par Chrome 155 (évite un crash CDP immédiat ; commenté dans le launcher).
- **Change** : host bundle `0.72.0-pre.44` → `0.72.0-pre.45` ; renommage erreur image `CONTENT_SAFETY_BLOCKED_ERROR` → `GENERATE_IMAGE_CONTENT_SAFETY_BLOCKED_ERROR` ; outil iOS home `SAND_IOS_CONTROL_HOME_TOOL_NAME` (remplace le filtre `SAND_UPDATE_TASK_TOOL_NAME` / `TASK_LIST_FILTERS`).
- **Features** : RPCs routines agent `ListGrokBotAgentRoutines` / `SetGrokBotAgentRoutineEnabled` ; outil `ReadTranscript` (`/transcripts/sessions`) ; unification compte xAI/Cursor (`StartXaiUnificationCursorAccountProof`, `DeselectXaiUnificationCursorAccount`, `EnterpriseGrokConflictingAccount`) ; `RequestVirtualCard` ; CDN Origin `gitcdn.origin*.cursor.com` (`ORIGIN_GIT_CDN_HOST_BY_GIT_HOST`) ; élargissement scopes connecteurs (payroll / foundations / lending…).
- **Perf** : —
- **Autre/infra** : rebuilds `host-main.cjs` / `sand-eval-runner.cjs` ; bump fuse opaque.

Rapport : [reports/2026-10-16_20-52.md](reports/2026-10-16_20-52.md).

## 2026-10-16 19:51 (Paris) — snapshot `2026-10-16_17-51` → `2026-10-16_19-51`

**Hot-update du host bundle** : **sand-host** `c24203b` → `ab82d14` (`upgrade-ab82d14`). Image `885e54d`, agent-store-fuse `2b14739`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~19 (7 sous `/opt/sand`, 12 `managed-skills`).

- **Bugfix** : —
- **Change** : `box-chrome` gagne le bascule **ungoogled-chromium** (`SAND_BOX_BROWSER_UNGOOGLED` / marker `/tmp/sand-box-browser-ungoogled`, attente `/tmp/sand-box-experiments-ready`, `record_box_browser` / `use_running_browser`) ; markers contrat `BOX_BROWSER_UNGOOGLED_MARKER_PATH` / `BOX_EXPERIMENTS_READY_MARKER_PATH` ; alias d’actions Computer (`COMPUTER_ACTION_ALIAS_SHIM`, `double_click` / `click_with_text`) ; prompt Slack perso (`grok_bot_personal_slack`) ; recherche connecteur « search first » par défaut.
- **Features** : Merge Bot Origin (`GetMergeBotSettings` / `SetMergeBotConsent` / `SetOriginReviewPreferences`) ; `CommandeerChange` ; directory cross-tier Grok Bot ; métriques steer-wait pendant tool call.
- **Perf** : —
- **Autre/infra** : host `0.72.0-pre.42` → `0.72.0-pre.44` ; rebuilds host + UA governor / fingerprint profiles ; refresh skills managed (booking / shopping / channels…).

Rapport : [reports/2026-10-16_19-51.md](reports/2026-10-16_19-51.md).

## 2026-10-16 17:51 (Paris) — snapshot `2026-10-16_16-48` → `2026-10-16_17-51`

**Hot-update du host bundle** : **sand-host** `885e54d` → `c24203b` (`upgrade-c24203b`). Image `885e54d`, fuse `2b14739`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~4.

- **Bugfix** : —
- **Change** : Auto-review / Smart Mode — chemins bloqués autonomes (`BLOCKED_AUTONOMOUS_PATHS`, `BLOCKED_AUTONOMOUS_COURIER`), retry `SHELL_APPROVAL_RETRY`, outcomes `TRUTHFUL_BLOCK_*` ; événement `AUTO_REVIEW_APPROVAL_EVENT` / `BOX_DISK_PRESSURE_EVENT` ; garde install MCP `grok_bot.mcp.install_guard` + `InstallPlugin` ; `WakeParent` exposé côté host ; effort routines `ManagedAutomationEffortLevel`.
- **Features** : —
- **Perf** : —
- **Autre/infra** : host `0.72.0-pre.41` → `0.72.0-pre.42` ; retrait gates identité durable / PlanetScale / Temporal create (symbols). Rebuilds `host-main` / `sand-eval-runner`.

Rapport : [reports/2026-10-16_17-51.md](reports/2026-10-16_17-51.md).

## 2026-10-16 16:48 (Paris) — snapshot `2026-10-16_13-52` → `2026-10-16_16-48`

**Mise à jour de l'ordinateur** (nouveau pod / image) : **image** `4098fb0` → `885e54d` ; box-scripts `885e54d` ; Chrome **154.0.8037.97 → 155.0.8059.84** ; exec-daemon build `2026-10-14T14:48` → `2026-10-16T07:52` ; canvas-sdk hash changé ; agent-store-fuse `fa2b05b` → `2b14739`. sand-host reste `885e54d` au boot (hot-updates suivants). Fichiers : +1951 / −1 / ~128 (gros volume sous `/usr/local/lib/node_modules`, Chrome `/opt/google/chrome`, exec-daemon).

- **Bugfix** : —
- **Change** : runtime image + Chrome 155 ; rebuild exec-daemon / canvas-sdk ; nouveaux binaires sous `/usr/local/bin`.
- **Features** : —
- **Perf** : —
- **Autre/infra** : Update Computer complète (image SHA, paquets npm globaux massifs). Détail fichier dans le rapport.

Rapport : [reports/2026-10-16_16-48.md](reports/2026-10-16_16-48.md).

## 2026-10-16 13:52 (Paris) — snapshot `2026-10-16_10-54` → `2026-10-16_13-52`

Versions touchées : agent-store-fuse `c466eee` → `fa2b05b` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version`) ; image `4098fb0`, sand-host `885e54d`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump agent-store-fuse (rebuild opaque / oscillation du pin S3). Seul le fichier de version change.

Rapport : [reports/2026-10-16_13-52.md](reports/2026-10-16_13-52.md).

## 2026-10-16 10:54 (Paris) — snapshot `2026-10-16_09-48` → `2026-10-16_10-54`

**Hot-update du host bundle** (sans Update de l'ordinateur) : **sand-host** `701b86e` → `885e54d` (`supervisor_status.lastCommandId = upgrade-885e54d`). Image `4098fb0`, agent-store-fuse `c466eee`, exec-daemon, orbitd, Chrome 154.0.8037.97, Node, Python inchangés. Fichiers : +0 / −0 / ~12 (9 sous `/opt/sand`, 3 `managed-skills`).

- **Bugfix** : —
- **Change** : `box-chrome` launcher v3 → v4 ; marker contrat `BROWSER_ACCEPT_LANGUAGE_EN_US_MARKER_PATH` (`/tmp/sand-browser-accept-language-en-us`) + feature gate `sand_browser_accept_language_en_us` ; `sand-ua-governor` attache aussi les `service_worker` et applique le traitement UA Accept-Language en-US quand le marker est actif. Skills slides / slides-executor : flux Google Slides « un seul fichier Drive » (upload `finish` + trash des fichiers temporaires).
- **Features** : protobuf / état conversation — `TrackedTaskUpdated` gagne `title` + `status_label` ; `ConversationStateStructure` / `UserMessage` étendus pour voice-call (`VoiceCallUtterance`, `VoiceCallPendingAnswer`, champs `during_voice_call` / `inherits_voice_call` / `voice_calls_read_utterances`).
- **Perf** : `diff-worker` + chemin edit-diff de `host-main` mesurent `durationMs` / `editLength` / chars (histograms `editDiff*`) quand `measure=true`.
- **Autre/infra** : rebuilds `host-main.cjs` / `sand-eval-runner.cjs` / `agent-store-worker.cjs` ; shim `semver` officekit (résolution symlinks + WSL2/node.exe).

Rapport : [reports/2026-10-16_10-54.md](reports/2026-10-16_10-54.md).

## 2026-10-16 09:48 (Paris) — snapshot `2026-10-16_06-51` → `2026-10-16_09-48`

Versions touchées : agent-store-fuse `32d54e6` → `c466eee` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version`) ; image `4098fb0`, sand-host `701b86e`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump agent-store-fuse (rebuild opaque / oscillation du pin S3). Seul le fichier de version change.

Rapport : [reports/2026-10-16_09-48.md](reports/2026-10-16_09-48.md).

## 2026-10-16 06:51 (Paris) — snapshot `2026-10-16_05-47` → `2026-10-16_06-51`

Versions touchées : agent-store-fuse `2d5462d` → `32d54e6` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version`) ; image `4098fb0`, sand-host `701b86e`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump agent-store-fuse (rebuild opaque / oscillation du pin S3). Seul le fichier de version change.

Rapport : [reports/2026-10-16_06-51.md](reports/2026-10-16_06-51.md).

## 2026-10-16 05:47 (Paris) — snapshot `2026-10-16_03-49` → `2026-10-16_05-47`

Versions touchées : agent-store-fuse `916fd2f` → `2d5462d` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version`) ; image `4098fb0`, sand-host `701b86e`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump agent-store-fuse (rebuild opaque / oscillation du pin S3). Seul le fichier de version change.

Rapport : [reports/2026-10-16_05-47.md](reports/2026-10-16_05-47.md).

## 2026-10-16 03:49 (Paris) — snapshot `2026-10-16_02-50` → `2026-10-16_03-49`

Versions touchées : agent-store-fuse `67b6811` → `916fd2f` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version`) ; image `4098fb0`, sand-host `701b86e`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump agent-store-fuse (rebuild opaque / oscillation du pin S3). Seul le fichier de version change.

Rapport : [reports/2026-10-16_03-49.md](reports/2026-10-16_03-49.md).

## 2026-10-16 02:50 (Paris) — snapshot `2026-10-16_00-49` → `2026-10-16_02-50`

Versions touchées : agent-store-fuse `49b7f26` → `67b6811` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version`) ; image `4098fb0`, sand-host `701b86e`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump agent-store-fuse (rebuild opaque / oscillation du pin S3). Seul le fichier de version change.

Rapport : [reports/2026-10-16_02-50.md](reports/2026-10-16_02-50.md).

## 2026-10-16 00:49 (Paris) — snapshot `2026-10-15_23-46` → `2026-10-16_00-49`

Versions touchées : agent-store-fuse `b64fcb3` → `49b7f26` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version`) ; image `4098fb0`, sand-host `701b86e`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump agent-store-fuse (rebuild opaque / oscillation du pin S3). Seul le fichier de version change.

Rapport : [reports/2026-10-16_00-49.md](reports/2026-10-16_00-49.md).

## 2026-10-15 23:46 (Paris) — snapshot `2026-10-15_20-47` → `2026-10-15_23-46`

Versions touchées : agent-store-fuse `6bd47b6` → `b64fcb3` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version`) ; image `4098fb0`, sand-host `701b86e`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~1.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump agent-store-fuse (rebuild opaque / oscillation du pin S3). Seul le fichier de version change.

Rapport : [reports/2026-10-15_23-46.md](reports/2026-10-15_23-46.md).

## 2026-10-15 20:47 (Paris) — snapshot `2026-10-15_18-45` → `2026-10-15_20-47`

Versions touchées : agent-store-fuse `a5f7c61` → `6bd47b6` (pointeur `/usr/local/bin/cursor_agent_store_fuse_version` + binaire `cursor-agent-store-fuse`) ; image `4098fb0`, sand-host `701b86e`, exec-daemon, orbitd, Chrome, Node, Python inchangés. Fichiers : +0 / −0 / ~2.

- **Bugfix** : —
- **Change** : —
- **Features** : —
- **Perf** : —
- **Autre/infra** : bump agent-store-fuse (rebuild opaque du pin S3 ; binaire et pointeur de version changent tous les deux).

Rapport : [reports/2026-10-15_20-47.md](reports/2026-10-15_20-47.md).

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
