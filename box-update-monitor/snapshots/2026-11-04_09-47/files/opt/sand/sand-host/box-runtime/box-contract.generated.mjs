// @generated SignedSource<<2c6b9283704646f07797801a806c33f39687acb6788093a939e420bb5b952d11>>
// Emitted by scripts/gen-box-contract.mts from sand/src/shared/box/box-contract.ts;
// regenerate with `pnpm --filter sand run gen:box-contract`.
export const BOX_BROWSER_COHERENT_FINGERPRINT_MARKER_PATH = "/tmp/sand-box-browser-coherent-fingerprint";
export const BOX_BROWSER_UNGOOGLED_MARKER_PATH = "/tmp/sand-box-browser-ungoogled";
export const BOX_EXPERIMENTS_READY_MARKER_PATH = "/tmp/sand-box-experiments-ready";
export const EGRESS_TUNNEL_FAIL_CLOSED_MARKER_PATH = "/tmp/sand-egress-tunnel-fail-closed";
export const BOX_BROWSER_WEBRTC_PROXY_ONLY_MARKER_PATH = "/tmp/sand-box-browser-webrtc-proxy-only";
export const BOX_BROWSER_NO_LOCAL_DNS_MARKER_PATH = "/tmp/sand-box-browser-no-local-dns";
export const BOX_BROWSER_DISABLE_QUIC_MARKER_PATH = "/tmp/sand-chrome-disable-quic";
export const BOX_BROWSER_FINGERPRINT_LOG_SITES_PATH = "/tmp/sand-browser-fingerprint-log-sites";
export const BOX_BROWSER_NO_SWIFTSHADER_MARKER_PATH = "/tmp/sand-box-browser-no-swiftshader";
export const BOX_BROWSER_STOCK_FONTS_MARKER_PATH = "/tmp/sand-box-browser-stock-fonts";
export const BOX_BROWSER_STOCK_NOTIFICATIONS_MARKER_PATH = "/tmp/sand-box-browser-stock-notifications";
export const BOX_CHROME_EXTRA_FLAGS_PATH = "/tmp/sand-box-chrome-extra-flags";
export const BOX_BROWSER_RUNNING_PATH_PREFIX = "/tmp/sand-box-browser-running";
export const BOX_BROWSER_GPU_MODE_PATH_PREFIX = "/tmp/sand-box-browser-gpu-mode";
export const BOX_BROWSER_GPU_MODES = ["swiftshader","angle-gl","legacy","stock","llvmpipe"];
export const BROWSER_ACCEPT_LANGUAGE_EN_US_MARKER_PATH = "/tmp/sand-browser-accept-language-en-us";
export const BOX_BROWSER_ACCEPT_LANGUAGE_PATH_PREFIX = "/tmp/sand-box-browser-accept-language";
export const BROWSER_EGRESS_TIMEZONE_MARKER_PATH = "/tmp/sand-browser-egress-timezone";
export const WEB_BOT_AUTH_SIGNED_CACHE_PATH = "/tmp/sand-web-bot-auth-signed.json";
export const WEB_BOT_AUTH_SIGNED_TTL_MS = 120000;
export const WEB_BOT_AUTH_SIGNATURE_SOURCES = ["fresh","box_cache","fleet_cache"];
export const WEB_BOT_AUTH_SIGNED_CACHE_MAX_BYTES = 64000;
export const WEB_BOT_AUTH_NOT_SIGNED_REASONS = ["origin_relayed","origin_denied","origin_unlisted","cooldown_relayed","cooldown_denied","cooldown_unlisted","backoff","probe_in_flight","http_401","http_429","http_other","no_token","bad_headers","error"];
export const WEB_BOT_AUTH_NOT_CONSIDERED_REASONS = ["tunnel_attached","marker_absent","no_signer","disabled_mid_sign","tailnet","non_https","daemon_absent"];
export const WEB_BOT_AUTH_STATE_READY_PATH = "/tmp/sand-web-bot-auth-state-ready";
export const WALL_EVIDENCE_WRITER_VERSION = 2;
export const WALL_EVIDENCE_MAIN_FRAME_RESPONSES_MAX_BYTES = 8192;
export const WALL_EVIDENCE_PROFILE_DIR_MAX_BYTES = 256;
export const WALL_EVIDENCE_PROFILE_ORIGIN_MAX_BYTES = 192;
export const WALL_EVIDENCE_COOKIE_LINEAGE_MAX_BYTES = 256;
export const WALL_EVIDENCE_HEADER_NAME_MAX_CHARS = 64;
export const WALL_EVIDENCE_HEADER_VALUE_MAX_CHARS = 128;
export const WALL_EVIDENCE_REDIRECT_CHAIN_MAX_HOPS = 8;
export const WALL_EVIDENCE_COOKIE_FLAGS = ["session-id","ubid-main","at-main","sess-at-main","x-main","csm-hit","__cf_bm","cf_clearance","datadome","_abck","bm_sz","_px3","_pxvid"];
export const WALL_EVIDENCE_HEADER_NAMES = ["server","content-type","content-length","cache-control","retry-after","via","age","vary","x-cache","accept-ch","critical-ch","cf-ray","cf-mitigated","cf-cache-status","x-datadome","akamai-grn","x-iinfo","x-cdn","x-amz-rid","x-amz-cf-id","x-amz-cf-pop","set-cookie","location","refresh","www-authenticate","proxy-authenticate","authentication-info"];
export const WALL_EVIDENCE_PRESENCE_ONLY_HEADER_NAMES = ["set-cookie","location","refresh","www-authenticate","proxy-authenticate","authentication-info"];
export const WALL_EVIDENCE_COOKIE_FACT_DOMAINS = ["amazon.com","amazon.ae","amazon.ca","amazon.co.jp","amazon.co.uk","amazon.com.au","amazon.com.be","amazon.com.br","amazon.com.mx","amazon.com.tr","amazon.de","amazon.eg","amazon.es","amazon.fr","amazon.ie","amazon.in","amazon.it","amazon.nl","amazon.pl","amazon.sa","amazon.se","amazon.sg"];
export const WALL_EVIDENCE_COOKIE_FACTS_MAX_BYTES = 4096;
export const WALL_EVIDENCE_COOKIE_COUNT_MAX = 10000;
export const BOX_WALL_EVIDENCE_RECORDERS_MARKER_PATH = "/tmp/sand-box-wall-evidence-recorders";
export function webBotAuthStatePath(port) {
  return "/tmp/sand-web-bot-auth-state-" + port + ".json";
}
export function wallEvidenceCookieFactsPath(port) {
  return "/tmp/sand-cookie-facts-" + port + ".json";
}
export function wallEvidenceMainFrameResponsesPath(port) {
  return "/tmp/sand-main-frame-responses-" + port + ".json";
}
export function wallEvidenceMainFrameResponseNamesPath(port) {
  return "/tmp/sand-main-frame-responses-" + port + ".names.json";
}
export function wallEvidenceMainFrameResponsesLockPath(port) {
  return "/tmp/sand-main-frame-responses-" + port + ".lock";
}
export function wallEvidenceProfileDirPath(port) {
  return "/tmp/sand-profile-dir-" + port + "";
}
export const WALL_EVIDENCE_COOKIE_LINEAGE_PATH = "/home/box/chrome-profile/.sand-cookie-lineage.json";
export const WALL_EVIDENCE_BOX_BOOT_ID_PATH = "/tmp/sand-box-boot-id";
export const WALL_EVIDENCE_BOX_BOOT_ID_MAX_BYTES = 64;
export const WALL_EVIDENCE_BOX_BOOT_ID_SOURCE = "^[A-Za-z0-9-]{1,64}$";
export const WALL_EVIDENCE_BOX_BOOT_ID_UNKNOWN = "unknown";
export const WALL_EVIDENCE_PROFILE_ORIGIN_FILE_NAME = ".sand-profile-origin.json";
export const WALL_EVIDENCE_RESTORE_FACTS_PATH = "/tmp/sand-cookie-restore-facts.json";
export const WALL_EVIDENCE_RESTORE_FACTS_MAX_BYTES = 256;
export const SAND_BOX_PORTS = {"primaryExecDaemon":1337,"primaryPty":1338,"primaryNovncWebsockify":6080,"windowRouter":1339,"forkNovncWebsockify":6081,"hostGateway":1340,"egressTunnelWebSocket":8790,"egressConnectProxy":8791};
export const SAND_BOX_FORK_PORT_BASES = {"execDaemon":14000,"pty":13600,"vnc":5900};
export const SAND_BOX_CDP_PORT_BASE = 9222;
export const SAND_BOX_TAILNET_CHROME_CDP_PORT = 9600;
export const SAND_BOX_WINDOW_TOKEN_DIR = "/tmp/sand-window-tokens.d";
export const SAND_BOX_NOVNC_TOKEN_DIR = "/tmp/sand-novnc-tokens.d";
export const SAND_BOX_DISPLAY_HEADER = "x-sand-display";
export const SAND_BOX_WINDOW_OWNER_HEADER = "x-sand-window-owner";
export const TAILNET_HOST_SUFFIXES_PATH = "/var/lib/sand/tailnet/host-suffixes";
