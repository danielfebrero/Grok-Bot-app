# Changelog — surface des méthodes RPC du host Sand/Cursor

Suivi snapshot par snapshot des **déclarations RPC** du host : méthodes
**ajoutées / supprimées / modifiées** (changement de `kind` ou des types I/O).
Extraction **exhaustive** : scan de tous les `.cjs/.js/.mjs` de `/opt/sand`,
relevé de chaque `typeName: "<pkg>.<Service>"` + son bloc `methods`, tous
packages confondus. Entrées les plus récentes en haut. Heures de Paris
(CEST/CET). Généré/alimenté par `check.sh`.

<!-- NOUVELLES-ENTREES-ICI -->

## 2026-11-14 17:54 (Paris) — `2026-11-13_18-00` → `2026-11-14_17-54`

Changement de surface RPC détecté : **24 méthode(s) ajoutée(s)**.

Ajoutées : `BackgroundComposerService.SetProjectJoinByLink`, `DashboardService.GetCloudAgentStartupScript`, `DashboardService.GetGiftCodeInfo`, `DashboardService.StartTeamGrokBotOwnershipTransfer`, `DashboardService.UpdateCloudAgentStartupScript`, `GrokBotService.AdminExplainGrokBotSlackAccess`, `GrokBotService.ArchiveGrokBotChangelogEntryDraftInternal`, `GrokBotService.CreateGrokBotChangelogEntryInternal`, `GrokBotService.DisconnectGrokBotThirdPartyConnection`, `GrokBotService.GetGrokBotChangelogEntryInternal`, `GrokBotService.GetGrokBotMessagingContactPhoto`, `GrokBotService.GetGrokBotPaperTermsInternal`, `GrokBotService.GetGrokBotShoppingOrderState`, `GrokBotService.GetGrokBotSlackWorkspaceInstall`, `GrokBotService.ListGrokBotChangelogEntriesInternal`, `GrokBotService.ListGrokBotThirdPartyConnections`, `GrokBotService.ListPublicGrokBotChangelog`, `GrokBotService.ResolveGrokBotAgentFileReads`, `GrokBotService.SetGrokBotChangelogEntryStatusInternal`, `GrokBotService.SetGrokBotPaperTermsInternal`, `GrokBotService.StartGrokBotSlackWorkspaceInstall`, `GrokBotService.UninstallGrokBotSlackWorkspaceInstall`, `GrokBotService.UpdateGrokBotChangelogEntryInternal`, `GrokBotService.UpsertGrokBotChangelogEntryDraftInternal`

Rapport : [reports/2026-11-14_17-54.md](reports/2026-11-14_17-54.md).


## 2026-11-13 18:00 (Paris) — `2026-11-12_09-56` → `2026-11-13_18-00`

Changement de surface RPC détecté : **50 méthode(s) ajoutée(s)**, **2 supprimée(s)**.

Ajoutées : `AppPublishService.GetAppsEligibility`, `AppPublishService.StartFormLinkOwnerCode`, `AppPublishService.SubmitTextedFormLink`, `BackgroundComposerService.ListPrivateWorkerSetupManifests`, `DashboardService.CancelGrokBatch`, `DashboardService.CreateGrokBatch`, `DashboardService.DeleteTeamGrokBotScimIdpIdentity`, `DashboardService.GenerateTeamGrokBotScimIdpSigningKey`, `DashboardService.GetEnterpriseGrokBotTeamTrialInternal`, `DashboardService.GetGrokBatch`, `DashboardService.GetGrokBatchRequestResult`, `DashboardService.GetTeamGrokBotScimIdpIdentity`, `DashboardService.ListGrokBatchRequests`, `DashboardService.ListGrokBatches`, `DashboardService.SetEnterpriseGrokBotTeamTrialExpirationInternal`, `DashboardService.UpdateTeamGrokBotScimIdpIdentity`, `DashboardService.UpdateTeamTailnetExitNode`, `GrokBotService.DeleteGrokBotFactory`, `GrokBotService.GetGrokBotPurchaseAutoApprove`, `GrokBotService.ListGrokBotFactories`, `GrokBotService.RunGrokBotFactory`, `GrokBotService.SaveGrokBotFactory`, `GrokBotService.SetGrokBotPurchaseAutoApprove`, `GrokBotService.SetGrokBotPurchaseSpendLimit`, `UnifiedDashboardService.AddUser` …

Supprimées : `DashboardService.GetPullRequestDraftPreference`, `DashboardService.SetPullRequestDraftPreference`

Rapport : [reports/2026-11-13_18-00.md](reports/2026-11-13_18-00.md).


## 2026-11-12 09:56 (Paris) — `2026-11-11_17-58` → `2026-11-12_09-56`

Changement de surface RPC détecté : **19 méthode(s) ajoutée(s)**.

Ajoutées : `AppPublishService.KeepApp`, `AppPublishService.RecordAppVisit`, `BackgroundComposerService.CreatePrivateWorkerPoolTrust`, `BackgroundComposerService.DeletePrivateWorkerPoolTrust`, `BackgroundComposerService.ListPrivateWorkerPoolTrusts`, `DashboardService.CommitOrganizationNetworkBlocklistUpload`, `DashboardService.CreateOrganizationNetworkBlocklistExport`, `DashboardService.CreateOrganizationNetworkBlocklistUpload`, `DashboardService.GetOrganizationNetworkBlocklist`, `DashboardService.LookupGrokAccountsByXUserInternal`, `DashboardService.RetryTeamXaiMembershipHold`, `DashboardService.SaveTeamPluginServerSettings`, `DashboardService.SetOrganizationNetworkBlocklistEnabled`, `GrokBotService.ConfirmGrokBotAccountPhoneCode`, `GrokBotService.GetGrokBotAccountPhoneStatus`, `GrokBotService.ListGrokBotOutcomes`, `GrokBotService.ListGrokBotUpdates`, `GrokBotService.MintSandBoxPluginSyncGitToken`, `GrokBotService.SendGrokBotAccountPhoneCode`

Rapport : [reports/2026-11-12_09-56.md](reports/2026-11-12_09-56.md).


## 2026-11-11 17:58 (Paris) — `2026-11-10_17-57` → `2026-11-11_17-58`

Changement de surface RPC détecté : **19 méthode(s) ajoutée(s)**, **8 supprimée(s)**.

Ajoutées : `DashboardService.CreateTeamGrokBotScimToken`, `DashboardService.DeleteTeamGrokBotScimBot`, `DashboardService.DisableTeamGrokBotScimDirectory`, `DashboardService.EnableTeamGrokBotScimDirectory`, `DashboardService.GetTeamGrokBotRoiPreview`, `DashboardService.GetTeamGrokBotScimDirectory`, `DashboardService.ListAuthorizedApplications`, `DashboardService.ListOrganizationRemoteControlDevices`, `DashboardService.ListTeamGrokBotScimBots`, `DashboardService.ListTeamRemoteControlDevices`, `DashboardService.RevokeAuthorizedApplication`, `DashboardService.RevokeOrganizationRemoteControlDevice`, `DashboardService.RevokeTeamGrokBotScimToken`, `DashboardService.RevokeTeamRemoteControlDevice`, `DashboardService.SendTestPushNotificationInternal`, `DashboardService.SetTeamGrokBotAgentTeamApproved`, `DashboardService.UpdateTeamGrokBotScimIpAllowlist`, `GrokBotService.PresignXaiConnectorIconUploadInternal`, `GrokBotService.StartGrokBotLocalPitman`

Supprimées : `BackgroundComposerService.AdminDeleteNamedAgent`, `BackgroundComposerService.AdminListTeamNamedAgents`, `BackgroundComposerService.AdminSettleInteraction`, `BackgroundComposerService.AdminUnparkFollowups`, `BackgroundComposerService.ProvisionSyntheticsServiceAccounts`, `BackgroundComposerService.StartCloudAgentLoadTest`, `DashboardService.GetSyncedClientSettings`, `DashboardService.UpdateSyncedClientSettings`

Rapport : [reports/2026-11-11_17-58.md](reports/2026-11-11_17-58.md).


## 2026-11-10 17:57 (Paris) — `2026-11-10_02-03` → `2026-11-10_17-57`

Changement de surface RPC détecté : **23 méthode(s) ajoutée(s)**.

Ajoutées : `AiService.TestTeamInferenceHookDecisionRules`, `AnalyticsService.LogStatsigExposures`, `BackgroundComposerService.ListPluginConsentGrants`, `BackgroundComposerService.RevokePluginConsentGrant`, `BackgroundComposerService.SetEnvironmentVmAccess`, `DashboardService.DeleteTeamGrokBotIdpConnection`, `DashboardService.GetBitbucketForgeInstallationGrants`, `DashboardService.GetTeamGrokBotIdpConnection`, `DashboardService.GrantBitbucketForgeInstallation`, `DashboardService.ListMcpAllowlistRuleMatches`, `DashboardService.ListTeamPluginsTable`, `DashboardService.RecordConnectorPolicyBlocked`, `DashboardService.RevokeBitbucketForgeInstallation`, `DashboardService.UpdateTeamGrokBotIdpConnection`, `DashboardService.UploadOAuthApplicationLogo`, `GrokBotService.ConfirmGrokBotMessagingPhoneCode`, `GrokBotService.GetGrokBotMessagingPrompt`, `GrokBotService.GetTeamSandSetupTest`, `GrokBotService.LeaveGrokBotMessagingGroup`, `GrokBotService.MarkGrokBotMessagingPromptSeen`, `GrokBotService.SendGrokBotMessagingPhoneCode`, `GrokBotService.StartTeamSandSetupTest`, `GrokBotService.StopTeamSandSetupTest`

Rapport : [reports/2026-11-10_17-57.md](reports/2026-11-10_17-57.md).


## 2026-11-10 02:03 (Paris) — `2026-11-09_18-09` → `2026-11-10_02-03`

Changement de surface RPC détecté : **3 méthode(s) ajoutée(s)**.

Ajoutées : `ControlService.AddWorkspaceFolder`, `GrokBotService.ListGrokBotUserSkills`, `GrokBotService.UpdateGrokBotAgentSlackSharing`

Rapport : [reports/2026-11-10_02-03.md](reports/2026-11-10_02-03.md).


## 2026-11-09 18:09 (Paris) — `2026-11-08_18-09` → `2026-11-09_18-09`

Changement de surface RPC détecté : **3 méthode(s) ajoutée(s)**.

Ajoutées : `GrokBotService.CancelGrokBotCartCheckout`, `GrokBotService.SendGrokBotChangelogNowInternal`, `GrokBotService.StartGrokBotCartCheckout`

Rapport : [reports/2026-11-09_18-09.md](reports/2026-11-09_18-09.md).


## 2026-11-08 18:09 (Paris) — `2026-11-07_18-03` → `2026-11-08_18-09`

Changement de surface RPC détecté : **8 méthode(s) ajoutée(s)**.

Ajoutées : `DashboardService.ApproveOAuthAuthorizationRequest`, `DashboardService.CreateOAuthApplication`, `DashboardService.DeleteOAuthApplication`, `DashboardService.GetOAuthAuthorizationRequest`, `DashboardService.ListOAuthApplications`, `DashboardService.RotateOAuthApplicationSecret`, `DashboardService.SetOAuthApplicationEnabled`, `DashboardService.UpdateOAuthApplication`

Rapport : [reports/2026-11-08_18-09.md](reports/2026-11-08_18-09.md).


## 2026-11-07 18:03 (Paris) — `2026-11-06_18-08` → `2026-11-07_18-03`

Changement de surface RPC détecté : **19 méthode(s) ajoutée(s)**, **1 supprimée(s)**.

Ajoutées : `AiService.SetVoiceChatEnabled`, `BackgroundComposerService.CompleteSlackSignIn`, `BackgroundComposerService.GetCursorFilesUrl`, `BackgroundComposerService.RedeemSlackSignInLink`, `ControlService.WarmCursorFilesServer`, `DashboardService.GetGrokCreditPackPurchaseAccount`, `DashboardService.GetOrganizationConnectorAccountOwnership`, `DashboardService.ListGrokModels`, `DashboardService.ReportGrokCreditPackPurchase`, `DashboardService.UpdateOrganizationConnectorAccountOwnership`, `DashboardService.UpdateTeamVoiceChatSetting`, `GrokBotService.AdminGetGrokBotModelComparison`, `GrokBotService.AdminGetGrokBotModelComparisonOptions`, `GrokBotService.AdminListGrokBotModelComparisons`, `GrokBotService.AdminSendGrokBotModelComparisonFollowUp`, `GrokBotService.AdminStartGrokBotModelComparison`, `GrokBotService.AdminStopGrokBotModelComparison`, `GrokBotService.DeleteGrokBotTryOnImage`, `GrokBotService.GetGrokBotMessagingGroupChatLog`

Supprimées : `GrokBotService.ListGrokBotTeamsJoinedTeams`

Rapport : [reports/2026-11-07_18-03.md](reports/2026-11-07_18-03.md).


## 2026-11-06 18:08 (Paris) — `2026-11-05_18-11` → `2026-11-06_18-08`

Changement de surface RPC détecté : **7 méthode(s) ajoutée(s)**.

Ajoutées : `GrokBotService.GetGrokBotAgentPreference`, `GrokBotService.HasGrokBotAgentPastMeetings`, `GrokBotService.RemoveGrokBotOnboardingDripStepInternal`, `GrokBotService.RunGrokBotAgentAutomationNow`, `GrokBotService.SetGrokBotAgentAutoReview`, `GrokBotService.SetGrokBotAgentPreference`, `GrokBotService.SetGrokBotOnboardingDripDisplayNameInternal`

Rapport : [reports/2026-11-06_18-08.md](reports/2026-11-06_18-08.md).



## 2026-11-05 18:11 (Paris) — `2026-11-04_18-03` → `2026-11-05_18-11`

Changement de surface RPC détecté : **17 méthode(s) ajoutée(s)**, **3 supprimée(s)**.

Ajoutées : `AppPublishService.PublishMerge`, `AppPublishService.RebuildProposalPreview`, `DashboardService.ConfirmStepUpLink`, `DashboardService.ConfirmUnifyGrokUnlinkAtConflict`, `DashboardService.DescribeStepUpLink`, `DashboardService.GetOrganizationGrokBotWallSnapshotControls`, `DashboardService.GetStepUpStatus`, `DashboardService.GetUserServedPlanInternal`, `DashboardService.ListCursorApiKeys`, `DashboardService.StartStepUp`, `DashboardService.UpdateOrganizationGrokBotWallSnapshotControls`, `GrokBotService.BeginOnePasswordConnection`, `GrokBotService.CompleteOnePasswordConnection`, `GrokBotService.GenerateGrokBotTryOnImage`, `GrokBotService.ListGrokBotTryOnImages`, `GrokBotService.RerunGrokBotAgentAutomationFromCard`, `GrokBotService.SetOnePasswordPreviousIntegration`

Supprimées : `DashboardService.CopyTeamIngressIpAllowlistToOrganization`, `DashboardService.GetTeamSandHibernatedOwnerVmTerminationSettings`, `DashboardService.UpdateTeamSandHibernatedOwnerVmTerminationSettings`

Rapport : [reports/2026-11-05_18-11.md](reports/2026-11-05_18-11.md).


## 2026-11-04 18:03 (Paris) — `2026-11-03_18-09` → `2026-11-04_18-03`

Changement de surface RPC détecté : **35 méthode(s) ajoutée(s)**, **2 supprimée(s)**.

Ajoutées : `AutomationsService.RerunGrokBotAutomationFromCard`, `BackgroundComposerService.UpdateProjectWorkerModel`, `DashboardService.GetGrokBotUsageBreakdown`, `DashboardService.GetSandEnterpriseSpendLimit`, `DashboardService.GetTeamUnifiedStatusInternal`, `DashboardService.GetTeamXaiDataRetention`, `DashboardService.ListGrokRequestLogBodies`, `DashboardService.ListGrokRequestLogs`, `DashboardService.SearchApiKeys`, `DashboardService.StaffArchiveServiceAccount`, `DashboardService.StaffDeleteServiceAccount`, `DashboardService.StaffListOwnerApiKeys`, `DashboardService.StaffListServiceAccountApiKeys`, `DashboardService.StaffResumeServiceAccount`, `DashboardService.StaffRevokeApiKey`, `DashboardService.StaffSearchApiKeys`, `DashboardService.StaffSuspendServiceAccount`, `DashboardService.UnlinkXaiUnificationCursorAccount`, `GrokBotService.ActOnGrokBotMeetingPrompt`, `GrokBotService.AnswerGrokBotLabelerQuestion`, `GrokBotService.CreateGrokBotOnboardingDripRouteInternal`, `GrokBotService.GetGrokBotPlaceDetails`, `GrokBotService.ListGrokBotCartItems`, `GrokBotService.MakeGrokBotAgentFilePrivate`, `GrokBotService.MakeGrokBotAgentSkillPrivate` …

Supprimées : `GrokBotService.BeginOnePasswordConnection`, `GrokBotService.CompleteOnePasswordConnection`

Rapport : [reports/2026-11-04_18-03.md](reports/2026-11-04_18-03.md).


## 2026-11-03 18:09 (Paris) — `2026-11-02_18-07` → `2026-11-03_18-09`

Changement de surface RPC détecté : **19 méthode(s) ajoutée(s)**.

Ajoutées : `AutomationsService.SetAutomationProject`, `DashboardService.CopyTeamIngressIpAllowlistToOrganization`, `DashboardService.GetOrganizationIngressIpAllowlist`, `DashboardService.GetTeamSandLocalEgressPrivateDestinations`, `DashboardService.ListAppMemberAppsApiKeys`, `DashboardService.ListOrganizationMemberSessions`, `DashboardService.ListServiceTokenApps`, `DashboardService.ListTeamAppsApiKeys`, `DashboardService.ListTeamMemberSessions`, `DashboardService.RevokeOrganizationMemberSession`, `DashboardService.RevokeTeamMemberSession`, `DashboardService.SetSandAutomationSlackWorkspace`, `DashboardService.SetServiceTokenApps`, `DashboardService.UpdateOrganizationIngressIpAllowlist`, `DashboardService.UpdateTeamSandLocalEgressPrivateDestinations`, `GrokBotService.AdminGetGrokBotAccountMigrationHeldSummary`, `GrokBotService.AdminGetGrokBotAgentResources`, `GrokBotService.LookupGrokBotShoppingProduct`, `GrokBotService.SetGrokBotAgentRoutineForTeam`

Rapport : [reports/2026-11-03_18-09.md](reports/2026-11-03_18-09.md).


## 2026-11-02 18:07 (Paris) — `2026-11-01_18-09` → `2026-11-02_18-07`

Changement de surface RPC détecté : **11 méthode(s) ajoutée(s)**, **3 supprimée(s)**.

Ajoutées : `DashboardService.GetTeamIngressIpAllowlist`, `DashboardService.MarkGrokBotTrainingOptInPromptSeen`, `DashboardService.ShouldShowGrokBotTrainingOptInPopup`, `DashboardService.UpdateTeamIngressIpAllowlist`, `GrokBotService.AdminCleanGrokBotAccountMigration`, `GrokBotService.AdminListGrokBotAccountMigrationLedger`, `GrokBotService.AdminOverrideGrokBotAccountMigrationItem`, `GrokBotService.AdminProposeGrokBotAccountMigrationPairs`, `GrokBotService.DeleteGrokBotUserProfilePhoto`, `GrokBotService.GetGrokBotUserProfile`, `GrokBotService.SetGrokBotUserProfilePhoto`

Supprimées : `BackgroundComposerService.CheckCustomerTunnelOwnership`, `BackgroundComposerService.ListCustomerPrivatelinkAllowedHosts`, `BackgroundComposerService.MintCustomerPrivatelinkProxyToken`

Rapport : [reports/2026-11-02_18-07.md](reports/2026-11-02_18-07.md).


## 2026-11-01 18:09 (Paris) — `2026-10-30_18-06` → `2026-11-01_18-09`

Changement de surface RPC détecté : **10 méthode(s) ajoutée(s)**.

Ajoutées : `DashboardService.GetTeamCloudAgentMcpAllowlist`, `DashboardService.UpdateTeamCloudAgentMcpAllowlist`, `GrokBotService.AdminApplyGrokBotAccountMigration`, `GrokBotService.AdminCancelGrokBotAccountMigrations`, `GrokBotService.AdminFailGrokBotAccountMigrationExportStart`, `GrokBotService.AdminGetGrokBotAccountMigration`, `GrokBotService.AdminListGrokBotAccountMigrations`, `GrokBotService.AdminResolveGrokBotAccountMigrationPairs`, `GrokBotService.AdminRetryGrokBotAccountMigrations`, `GrokBotService.AdminSubmitGrokBotAccountMigrations`

Rapport : [reports/2026-11-01_18-09.md](reports/2026-11-01_18-09.md).


## 2026-10-30 18:06 (Paris) — `2026-10-28_18-13` → `2026-10-30_18-06`

Changement de surface RPC détecté : **15 méthode(s) ajoutée(s)**.

Ajoutées : `AutomationsService.GetDebugBotTeamSettings`, `BackgroundComposerService.CheckCustomerTunnelOwnership`, `BackgroundComposerService.EnsureProjectGroupChat`, `DashboardService.GetCanvasPageBrand`, `DashboardService.GetProfileOriginContributions`, `DashboardService.GetSlackLinkPreview`, `DashboardService.ListSlackLinkedWorkspaces`, `DashboardService.RequestSandSeatTierUpgrade`, `DashboardService.RunCanvasConnectorAction`, `GrokBotService.AdminStartGrokBotDream`, `GrokBotService.AnswerGrokBotMessagingCard`, `GrokBotService.GetGrokBotInviteByEmailAccess`, `GrokBotService.GetGrokBotMessagingCard`, `GrokBotService.ReportGrokBotMessagingCapabilities`, `GrokBotService.SummarizeGrokBotTeamReviewItems`

Rapport : [reports/2026-10-30_18-06.md](reports/2026-10-30_18-06.md).


## 2026-10-28 18:13 (Paris) — `2026-10-28_06-05` → `2026-10-28_18-13`

Changement de surface RPC détecté : **8 méthode(s) ajoutée(s)**.

Ajoutées : `AppPublishService.ArmMergeRepublish`, `AppPublishService.DisarmMergeRepublish`, `AppPublishService.PrepareDefaultBranchPublish`, `GrokBotService.AdminGetSandBoxActiveVm`, `GrokBotService.ResolveGrokBotFlightBookingLink`, `GrokBotService.SearchGrokBotSimilarProducts`, `GrokBotService.SearchGrokBotTranscripts`, `GrokBotService.UpdateGrokBotMeetingSettings`

Rapport : [reports/2026-10-28_18-13.md](reports/2026-10-28_18-13.md).


## 2026-10-28 06:05 (Paris) — `2026-10-27_20-11` → `2026-10-28_06-05`

Changement de surface RPC détecté : **15 méthode(s) ajoutée(s)**.

Ajoutées : `DashboardService.GetMcpBackendAccessToken`, `DashboardService.GetTeamGrokBotPublishedWorkerImage`, `DashboardService.GetTeamGrokBotSlackWorkspaceAccess`, `DashboardService.ListTeamGrokBotCustomerHostedWorkers`, `DashboardService.RefreshCanvasSources`, `DashboardService.SetTeamGrokBotSlackWorkspaceAccess`, `DashboardService.UpdateTeamPersonalApiKeyCreationSetting`, `GrokBotService.CreateGrokBotOnboardingDripInternal`, `GrokBotService.CreateGrokBotOnboardingDripStepInternal`, `GrokBotService.GetGrokBotOnboardingDripInternal`, `GrokBotService.ListGrokBotOnboardingDripsInternal`, `GrokBotService.ReorderGrokBotOnboardingDripStepsInternal`, `GrokBotService.SetGrokBotOnboardingDripStatusInternal`, `GrokBotService.UpdateGrokBotOnboardingDripInternal`, `GrokBotService.UpdateGrokBotOnboardingDripStepInternal`

Rapport : [reports/2026-10-28_06-05.md](reports/2026-10-28_06-05.md).


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
