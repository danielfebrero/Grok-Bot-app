// @generated SignedSource<<3721c0661b4787b927bff82d2f06817040a71a915fd0caa77c34ad297e229e7b>>
// Emitted by scripts/gen-box-contract.mts from sand/src/shared/box/box-contract.ts;
// regenerate with `pnpm --filter sand run gen:box-contract`.
export const ENABLE_SPOOF_GPU_MARKER_PATH = "/tmp/sand-enable-spoof-gpu";
export const BROWSER_HARDWARE_PROFILE_MARKER_PATH = "/tmp/sand-browser-hardware-profile";
export const BOX_BROWSER_UNGOOGLED_MARKER_PATH = "/tmp/sand-box-browser-ungoogled";
export const BOX_EXPERIMENTS_READY_MARKER_PATH = "/tmp/sand-box-experiments-ready";
export const BOX_BROWSER_RUNNING_PATH_PREFIX = "/tmp/sand-box-browser-running";
export const BROWSER_ACCEPT_LANGUAGE_EN_US_MARKER_PATH = "/tmp/sand-browser-accept-language-en-us";
export const BROWSER_EGRESS_TIMEZONE_MARKER_PATH = "/tmp/sand-browser-egress-timezone";
export const WEB_BOT_AUTH_SIGNED_CACHE_PATH = "/tmp/sand-web-bot-auth-signed.json";
export const WEB_BOT_AUTH_SIGNED_TTL_MS = 120000;
export const WEB_BOT_AUTH_SIGNATURE_SOURCES = ["fresh","box_cache","fleet_cache"];
export const SAND_BOX_PORTS = {"primaryExecDaemon":1337,"primaryPty":1338,"primaryNovncWebsockify":6080,"windowRouter":1339,"forkNovncWebsockify":6081,"hostGateway":1340,"egressTunnelWebSocket":8790,"egressConnectProxy":8791};
export const SAND_BOX_FORK_PORT_BASES = {"execDaemon":14000,"pty":13600,"vnc":5900};
export const SAND_BOX_CDP_PORT_BASE = 9222;
export const SAND_BOX_WINDOW_TOKEN_DIR = "/tmp/sand-window-tokens.d";
export const SAND_BOX_NOVNC_TOKEN_DIR = "/tmp/sand-novnc-tokens.d";
export const SAND_BOX_DISPLAY_HEADER = "x-sand-display";
export const SAND_BOX_WINDOW_OWNER_HEADER = "x-sand-window-owner";
export const TAILNET_HOST_SUFFIXES_PATH = "/var/lib/sand/tailnet/host-suffixes";
