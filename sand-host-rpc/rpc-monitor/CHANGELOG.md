# Changelog — surface des méthodes RPC du host Sand/Cursor

Suivi snapshot par snapshot des **déclarations RPC** du host : méthodes
**ajoutées / supprimées / modifiées** (changement de `kind` ou des types I/O).
Extraction **exhaustive** : scan de tous les `.cjs/.js/.mjs` de `/opt/sand`,
relevé de chaque `typeName: "<pkg>.<Service>"` + son bloc `methods`, tous
packages confondus. Entrées les plus récentes en haut. Heures de Paris
(CEST/CET). Généré/alimenté par `check.sh`.

<!-- NOUVELLES-ENTREES-ICI -->

## 2026-10-27 20:11 (Paris) — `2026-10-27_18-06` → `2026-10-27_20-11`

Changement de surface RPC détecté : **1 méthode(s) ajoutée(s)**.

Ajoutées : `DashboardService.SyncClientSettings`

Rapport : [reports/2026-10-27_20-11.md](reports/2026-10-27_20-11.md).


## 2026-10-27 18:06 (Paris) — `2026-10-27_10-01` → `2026-10-27_18-06`

Changement de surface RPC détecté : **3 méthode(s) ajoutée(s)**, **330 supprimée(s)**.

Ajoutées : `DashboardService.GetInstalledPluginDetail`, `DashboardService.ListStorefrontPlugins`, `DashboardService.RetrySandMessagesPermission`

Supprimées : `OriginService.AcceptShareInvite`, `OriginService.AddAppSigningKey`, `OriginService.AddAssignee`, `OriginService.AddCommentToThread`, `OriginService.AddLabelToChange`, `OriginService.AddNamespaceInboundIpAllowlistEntry`, `OriginService.AddNamespaceSshCertificateAuthority`, `OriginService.AddSshPublicKey`, `OriginService.ApplyMergeConflictResolution`, `OriginService.ApplyReviewRequestChanges`, `OriginService.BatchGetChanges`, `OriginService.BatchGetCommitCheckRollups`, `OriginService.BatchGetRepoContentAtSha`, `OriginService.CanMerge`, `OriginService.CancelMergeQueueEntry`, `OriginService.CancelMyAppInstallRequest`, `OriginService.CancelOriginCiJob`, `OriginService.CancelOriginCiRun`, `OriginService.CancelPullRequestBypassRequest`, `OriginService.CheckChangeCreationAccess`, `OriginService.CheckMarketplaceListingSlug`, `OriginService.ClaimCursorInboxPrAction`, `OriginService.CloseReviewRequest`, `OriginService.CommandeerChange`, `OriginService.CompareCommits` …

Rapport : [reports/2026-10-27_18-06.md](reports/2026-10-27_18-06.md).


## 2026-10-27 10:01 (Paris) — `2026-10-26_18-12` → `2026-10-27_10-01`

Changement de surface RPC détecté : **34 méthode(s) ajoutée(s)**.

Ajoutées : `AppPublishService.CancelReviewOverride`, `AppPublishService.RequestReviewOverride`, `BackgroundComposerService.ListPlacementRepoCandidates`, `DashboardService.ActivateTeamAppsReviewRulesVersion`, `DashboardService.DecideAppReviewOverride`, `DashboardService.GetOrganizationSlackInboundInstallPolicy`, `DashboardService.GetSyncedClientSettings`, `DashboardService.GetTeamAppsReviewRules`, `DashboardService.GetTeamAppsReviewRulesVersion`, `DashboardService.GetTeamSandEgressHttpRules`, `DashboardService.GetTeamSandTrustedCaCertificates`, `DashboardService.GrantAppReviewOverride`, `DashboardService.ListAppReviewOverrides`, `DashboardService.ListInstalledPlugins`, `DashboardService.ListOrganizationSlackConnectedWorkspaces`, `DashboardService.ListSlackConnectedWorkspaces`, `DashboardService.ListTeamAppsReviewRulesVersions`, `DashboardService.RevokeAppReviewOverride`, `DashboardService.RevokeCollapseSourceProof`, `DashboardService.SaveTeamAppsReviewRules`, `DashboardService.SendBonfireInvite`, `DashboardService.UpdateOrganizationSlackInboundInstallPolicy`, `DashboardService.UpdateSyncedClientSettings`, `DashboardService.UpdateTeamSandEgressHttpRules`, `DashboardService.UpdateTeamSandTrustedCaCertificates` …

Rapport : [reports/2026-10-27_10-01.md](reports/2026-10-27_10-01.md).


## 2026-10-26 18:12 (Paris) — `2026-10-25_18-01` → `2026-10-26_18-12`

Changement de surface RPC détecté : **8 méthode(s) ajoutée(s)**.

Ajoutées : `AppPublishService.ListProjectPublishState`, `BackgroundComposerService.GetPrAgentAttribution`, `DashboardService.SetGrokConnectorEnabled`, `GrokBotService.RetryGrokBotOutboundCallRecording`, `OriginService.GetFocusedGroupsOfChanges`, `OriginService.GetInboxBuckets`, `OriginService.MarkPullRequestsAsViewed`, `OriginService.UpdateRepoNeverDraftSetting`

Rapport : [reports/2026-10-26_18-12.md](reports/2026-10-26_18-12.md).


## 2026-10-25 18:01 (Paris) — `2026-10-24_18-01` → `2026-10-25_18-01`

Changement de surface RPC détecté : **4 méthode(s) ajoutée(s)**.

Ajoutées : `GrokBotService.GetGrokBotIMessageDraftSender`, `GrokBotService.ListGrokBotAgentsForOperatorInternal`, `GrokBotService.StartGrokBotSilentSubagentInternal`, `OriginService.GetChangesetOwnerRules`

Rapport : [reports/2026-10-25_18-01.md](reports/2026-10-25_18-01.md).


## 2026-10-24 18:01 (Paris) — `2026-10-24_09-52-38` → `2026-10-24_18-01`

Changement de surface RPC détecté : **27 méthode(s) ajoutée(s)**, **2 supprimée(s)**.

Ajoutées : `AppPublishService.CheckPageGrantOwner`, `AppPublishService.ResolvePublishedPageReplyLine`, `AppPublishService.StartPageOwnerCode`, `AutomationsService.GetAgentSdkApplicationIcon`, `BackgroundComposerService.GetBackgroundComposerRepoUrl`, `DashboardService.GetOrganizationGoogleDriveBlockedSharedDrives`, `DashboardService.GetTeamDataRetentionPolicy`, `DashboardService.HandOffMcpOAuthCode`, `DashboardService.SetActiveTailnet`, `DashboardService.UpdateOrganizationGoogleDriveBlockedSharedDrives`, `DashboardService.UpdateTeamDataRetentionPolicy`, `GrokBotService.AdminEnqueueGrokBotServiceAccountMessage`, `GrokBotService.AdminListGrokBotServiceAccountSessions`, `GrokBotService.AdminReadGrokBotServiceAccountConversation`, `GrokBotService.AttachGrokBotAgentDelegatedCredential`, `GrokBotService.CallGrokBotMcpAppTool`, `GrokBotService.ClaimGrokBotSlackConnect`, `GrokBotService.GetGrokBotDefaultPaymentMethod`, `GrokBotService.HangUpGrokBotOutboundCall`, `GrokBotService.NudgeOnePasswordAccessRequest`, `GrokBotService.RemoveGrokBotAgentDelegatedCredential`, `GrokBotService.ResolveGrokBotAddress`, `GrokBotService.SearchGrokBotPlacesArea`, `GrokBotService.SetGrokBotDefaultPaymentMethod`, `GrokBotService.SuggestGrokBotAddresses` …

Supprimées : `GrokBotService.EnsureGrokBotCursorMainBot`, `OriginService.DisableOriginUserSlackReviewReminder`

Rapport : [reports/2026-10-24_18-01.md](reports/2026-10-24_18-01.md).


## 2026-10-24 (Paris) — baseline

Première capture de référence (`snapshot.sh`).

- **21 services**, **2317 déclarations RPC** distinctes, tous packages :
  `aiserver.v1` (16 services), `agent.v1` (3), `origin.v1` (1), `orbit.v1` (1).
- Fichiers porteurs de définitions : `sand-host/host-main.cjs`,
  `sand-host/sand-eval-runner.cjs`.
- Signature suivie par méthode : `I|O|kind` (type requête, type réponse, kind Connect).
- Les snapshots suivants ne sont listés ici que s'ils présentent un vrai diff.

Voir le snapshot baseline dans `snapshots/` et le manifeste `manifest.json`.
