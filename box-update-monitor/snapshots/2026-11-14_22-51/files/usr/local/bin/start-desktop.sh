#!/usr/bin/env bash
set -euo pipefail

scrub_gateway_token_env() {
	local gateway_token="${SAND_GATEWAY_TOKEN:-}" tunnel_bearer="${SAND_EGRESS_TUNNEL_BEARER:-}"
	local entry env_name token_value
	GATEWAY_TOKEN_ENV_UNSETS=()
	while IFS= read -r -d '' entry; do
		env_name="${entry%%=*}"
		[ "${env_name}" != "${entry}" ] || continue
		for token_value in "${gateway_token}" "${tunnel_bearer}"; do
			if [ -n "${token_value}" ] && [[ "${entry#*=}" == *"${token_value}"* ]]; then
				if [[ "${env_name}" =~ ^[A-Za-z_][A-Za-z0-9_]*$ ]]; then
					unset "${env_name}"
				else
					GATEWAY_TOKEN_ENV_UNSETS+=(-u "${env_name}")
				fi
				break
			fi
		done
	done < <(env -0)
	unset SAND_GATEWAY_TOKEN SAND_EGRESS_TUNNEL_BEARER
}
scrub_gateway_token_env
[ "${#GATEWAY_TOKEN_ENV_UNSETS[@]}" -eq 0 ] || exec env "${GATEWAY_TOKEN_ENV_UNSETS[@]}" "${BASH}" "$0" "$@"

export DISPLAY="${DISPLAY:-:1}"
export HOME="${HOME:-/home/box}"
export XDG_RUNTIME_DIR="${SAND_XDG_RUNTIME_DIR:-${XDG_RUNTIME_DIR:-/tmp/xdg-runtime}}"
mkdir -p "${XDG_RUNTIME_DIR}"
chmod 700 "${XDG_RUNTIME_DIR}" || true

if [ -r /usr/local/bin/sand-desktop-supervise.sh ]; then
	# shellcheck source=/dev/null
	. /usr/local/bin/sand-desktop-supervise.sh
fi
if ! command -v sand_desktop_register >/dev/null 2>&1; then
	sand_desktop_register() { :; }
fi

SAND_BOX_CGROUPS_LIB="${SAND_BOX_CGROUPS_LIB:-/usr/local/bin/box-cgroups.sh}"
if [ -r "${SAND_BOX_CGROUPS_LIB}" ]; then
	# shellcheck source=/dev/null
	. "${SAND_BOX_CGROUPS_LIB}"
fi
if ! command -v sand_cgroup_join >/dev/null 2>&1; then
	sand_cgroup_join() { :; }
fi
sand_cgroup_join "${SAND_CGROUP_INTERACTIVE_NAME:-interactive}"

SAND_DESKTOP_NUM="${DISPLAY#:}"
SAND_DESKTOP_NUM="${SAND_DESKTOP_NUM%%.*}"
SAND_DESKTOP_GROUP="d${SAND_DESKTOP_NUM:-1}"

/usr/local/bin/ensure-machine-id || true

# >>> box screen (from sand/src/shared/box/box-contract.ts; regenerate: pnpm --filter sand run gen:box-ports) >>>
SCREEN_WIDTH="${SCREEN_WIDTH:-1280}"
SCREEN_HEIGHT="${SCREEN_HEIGHT:-800}"
# <<< box screen <<<

# >>> box port table (from sand/src/shared/box/box-contract.ts; regenerate: pnpm --filter sand run gen:box-ports) >>>
SAND_BOX_NOVNC_TOKEN_DIR="/tmp/sand-novnc-tokens.d"
VNC_PORT="${SAND_VNC_PORT:-5900}"
NOVNC_PORT="${SAND_NOVNC_PORT:-6080}"
# <<< box port table <<<
NOVNC_HEARTBEAT_S="${SAND_NOVNC_HEARTBEAT_S:-30}"

if [ -z "${DBUS_SESSION_BUS_ADDRESS:-}" ]; then
	eval "$(dbus-launch --sh-syntax)"
fi

BOX_DISPLAY_NUM="${DISPLAY#:}"
BOX_DISPLAY_NUM="${BOX_DISPLAY_NUM%%.*}"
if [ "${BOX_DISPLAY_NUM:-1}" -ge 2 ] 2>/dev/null; then
	BOX_USER_XDG_DIR="/tmp/xdg-runtime-box-${BOX_DISPLAY_NUM}"
else
	BOX_USER_XDG_DIR="/tmp/xdg-runtime-box"
fi
if [ "$(id -u)" -eq 0 ]; then
	install -d -o box -g box -m 700 "${BOX_USER_XDG_DIR}" 2>/dev/null || true
	rm -f "${BOX_USER_XDG_DIR}/dbus-session-address"
	BOX_DBUS_ADDRESS="$(runuser -u box -- env \
		XDG_RUNTIME_DIR="${BOX_USER_XDG_DIR}" DISPLAY="${DISPLAY}" HOME=/home/box \
		dbus-launch --sh-syntax 2>/dev/null \
		| sed -n "s/^DBUS_SESSION_BUS_ADDRESS='\(.*\)';/\1/p")"
	if [ -n "${BOX_DBUS_ADDRESS}" ]; then
		printf '%s' "${BOX_DBUS_ADDRESS}" >"${BOX_USER_XDG_DIR}/dbus-session-address"
		chown box:box "${BOX_USER_XDG_DIR}/dbus-session-address" 2>/dev/null || true
	fi
else
	mkdir -p "${BOX_USER_XDG_DIR}" 2>/dev/null || true
	chmod 700 "${BOX_USER_XDG_DIR}" 2>/dev/null || true
	printf '%s' "${DBUS_SESSION_BUS_ADDRESS:-}" >"${BOX_USER_XDG_DIR}/dbus-session-address" 2>/dev/null || true
fi

rm -f /tmp/xdg-runtime-box/sand-gateway-credential /tmp/xdg-runtime-box-*/sand-gateway-credential || true

/usr/local/bin/box-bounded-log --run "/tmp/xvfb${DISPLAY}.log" -- \
	/usr/local/bin/box-xvfb "${DISPLAY}" -screen 0 "${SCREEN_WIDTH}x${SCREEN_HEIGHT}x24" -ac +extension GLX +render -noreset &
sand_desktop_register "${SAND_DESKTOP_GROUP}" xvfb 0 "/tmp/xvfb${DISPLAY}.log" "$!" -- \
	/usr/local/bin/box-bounded-log --run "/tmp/xvfb${DISPLAY}.log" -- \
	/usr/local/bin/box-xvfb "${DISPLAY}" -screen 0 "${SCREEN_WIDTH}x${SCREEN_HEIGHT}x24" -ac +extension GLX +render -noreset

DISPLAY_READY=false
for _ in $(seq 1 100); do
	if xdpyinfo -display "${DISPLAY}" >/dev/null 2>&1; then
		DISPLAY_READY=true
		break
	fi
	sleep 0.2
done

if [ "${DISPLAY_READY}" != "true" ]; then
	echo "X display did not become ready" >&2
	exit 1
fi

if [ -x /usr/local/bin/sand-wallpaper ]; then
	/usr/local/bin/sand-wallpaper paint || true
else
	hsetroot -cover /usr/share/backgrounds/cursor-box-wallpaper.jpg \
		|| hsetroot -solid "#1e2330" \
		|| xsetroot -solid "#1e2330" \
		|| true
fi

X11VNC_ARGS=(
	-display "${DISPLAY}"
	-localhost
	-nopw
	-shared
	-forever
	-noxdamage
	-rfbport "${VNC_PORT}"
	-quiet
)
/usr/local/bin/box-bounded-log --run "/tmp/x11vnc${DISPLAY}.log" -- \
	/usr/local/bin/box-x11vnc "${X11VNC_ARGS[@]}" &
sand_desktop_register "${SAND_DESKTOP_GROUP}" x11vnc 3 "/tmp/x11vnc${DISPLAY}.log" "$!" -- \
	/usr/local/bin/box-bounded-log --run "/tmp/x11vnc${DISPLAY}.log" -- \
	/usr/local/bin/box-x11vnc "${X11VNC_ARGS[@]}"

if [ -n "${SAND_NOVNC_TOKEN:-}" ]; then
	TOKEN_DIR="${SAND_NOVNC_TOKEN_DIR:-${SAND_BOX_NOVNC_TOKEN_DIR}}"
	mkdir -p "${TOKEN_DIR}"
	printf '%s: localhost:%s\n' "${SAND_NOVNC_TOKEN}" "${VNC_PORT}" \
		>"${TOKEN_DIR}/${SAND_NOVNC_TOKEN}"
else
	WEBSOCKIFY_ARGS=(
		--web=/usr/share/novnc
		--heartbeat="${NOVNC_HEARTBEAT_S}"
		"0.0.0.0:${NOVNC_PORT}"
		"localhost:${VNC_PORT}"
	)
	/usr/local/bin/box-bounded-log --run "/tmp/novnc${DISPLAY}.log" -- \
		websockify "${WEBSOCKIFY_ARGS[@]}" &
	sand_desktop_register "${SAND_DESKTOP_GROUP}" websockify 4 "/tmp/novnc${DISPLAY}.log" "$!" -- \
		/usr/local/bin/box-bounded-log --run "/tmp/novnc${DISPLAY}.log" -- \
		websockify "${WEBSOCKIFY_ARGS[@]}"
fi

/usr/local/bin/box-bounded-log --run "/tmp/xfwm4${DISPLAY}.log" -- \
	/usr/local/bin/box-xfwm4 --compositor=off &
sand_desktop_register "${SAND_DESKTOP_GROUP}" xfwm4 1 "/tmp/xfwm4${DISPLAY}.log" "$!" -- \
	/usr/local/bin/box-bounded-log --run "/tmp/xfwm4${DISPLAY}.log" -- \
	/usr/local/bin/box-xfwm4 --compositor=off

for _ in $(seq 1 50); do
	if xprop -root _NET_SUPPORTING_WM_CHECK >/dev/null 2>&1; then
		break
	fi
	sleep 0.2
done

PICOM_ARGS=(
	--backend xrender
	--no-vsync
	--no-frame-pacing
	--no-use-damage
)
/usr/local/bin/box-bounded-log --run "/tmp/picom${DISPLAY}.log" -- \
	/usr/local/bin/box-picom "${PICOM_ARGS[@]}" &
sand_desktop_register "${SAND_DESKTOP_GROUP}" picom 2 "/tmp/picom${DISPLAY}.log" "$!" -- \
	/usr/local/bin/box-bounded-log --run "/tmp/picom${DISPLAY}.log" -- \
	/usr/local/bin/box-picom "${PICOM_ARGS[@]}"

# GTK3's org.gtk.Settings.FileChooser schema defines window-size and window-position
# as the GtkFileChooserDialog's window size and position, and the dconf writes below
# seed those keys (https://gitlab.gnome.org/GNOME/gtk/-/raw/gtk-3-24/gtk/org.gtk.Settings.FileChooser.gschema.xml).
dconf write /org/gtk/settings/file-chooser/window-size "(1100, 680)" 2>/dev/null || true
dconf write /org/gtk/settings/file-chooser/window-position "(90, 60)" 2>/dev/null || true
if [ "$(id -u)" -eq 0 ]; then
	runuser -u box -- env HOME=/home/box \
		dconf write /org/gtk/settings/file-chooser/window-size "(1100, 680)" 2>/dev/null || true
	runuser -u box -- env HOME=/home/box \
		dconf write /org/gtk/settings/file-chooser/window-position "(90, 60)" 2>/dev/null || true
fi

mkdir -p "${HOME}/.local/share/applications"
# StartupWMClass declares the WM class the launched application's window carries
# (https://specifications.freedesktop.org/desktop-entry-spec/latest/recognized-keys.html),
# so it must equal the class box-chrome passes with --class.
cat >"${HOME}/.local/share/applications/box-chrome.desktop" <<'EOF'
[Desktop Entry]
Version=1.0
Type=Application
Name=Google Chrome
Exec=/usr/local/bin/box-chrome --new-window %U
Icon=google-chrome
StartupWMClass=box-chrome
Categories=Network;WebBrowser;
MimeType=x-scheme-handler/http;x-scheme-handler/https;
EOF

mkdir -p "${HOME}/.config" 2>/dev/null || true
cat >"${HOME}/.config/mimeapps.list" <<'EOF'
[Default Applications]
x-scheme-handler/http=box-chrome.desktop
x-scheme-handler/https=box-chrome.desktop
EOF

write_dconf_setting() {
	local path="$1"
	local value="$2"
	for _ in $(seq 1 20); do
		dconf write "${path}" "${value}" >/dev/null 2>&1 || true
		if [ "$(dconf read "${path}" 2>/dev/null || true)" = "${value}" ]; then
			return 0
		fi
		sleep 0.1
	done
	echo "start-desktop: failed to persist dconf ${path}=${value}" >&2
	return 1
}

# At start Plank loads only the .dockitem files present in its launchers folder, in
# its dock-items order, then rewrites that setting from what it loaded, so a seeded
# dock-items entry whose file is missing at start is dropped (https://github.com/ricotz/plank/blob/master/lib/DockController.vala).
configure_plank_dock() {
	local dock_name="$1"
	local launchers="${HOME}/.config/plank/${dock_name}/launchers"
	local dconf_path="/net/launchpad/plank/docks/${dock_name}"
	local chrome_launcher="${HOME}/.local/share/applications/box-chrome.desktop"
	local thunar_launcher="/usr/share/applications/thunar.desktop"
	local terminal_launcher="/usr/share/applications/xfce4-terminal.desktop"
	mkdir -p "${launchers}"
	printf '[PlankDockItemPreferences]\nLauncher=file://%s\n' \
		"${chrome_launcher}" \
		>"${launchers}/chrome.dockitem"
	printf '[PlankDockItemPreferences]\nLauncher=file://%s\n' \
		"${thunar_launcher}" \
		>"${launchers}/thunar.dockitem"
	printf '[PlankDockItemPreferences]\nLauncher=file://%s\n' \
		"${terminal_launcher}" \
		>"${launchers}/xfce4-terminal.dockitem"
	write_dconf_setting "${dconf_path}/dock-items" "['chrome.dockitem', 'thunar.dockitem', 'xfce4-terminal.dockitem']" || true
	write_dconf_setting "${dconf_path}/position" "'bottom'" || true
	write_dconf_setting "${dconf_path}/theme" "'Transparent'" || true
	write_dconf_setting "${dconf_path}/icon-size" 48 || true
	# Plank's 'none' reserves a strut that shrinks maximized windows, and window-dodge hides the
	# dock while any window overlaps it, not just the focused one (https://github.com/ricotz/plank/blob/master/lib/HideManager.vala).
	write_dconf_setting "${dconf_path}/hide-mode" "'window-dodge'" || true
	write_dconf_setting "${dconf_path}/unhide-delay" 250 || true
	write_dconf_setting "${dconf_path}/hide-delay" 0 || true
}

start_dock_once() {
	local pid_file="$1" log="$2"
	shift 2
	local lock_file="${pid_file}.lock"
	(
		flock -w 10 9 2>/dev/null || true
		local existing=""
		if [ -f "${pid_file}" ]; then
			existing="$(tr -dc '0-9' <"${pid_file}" 2>/dev/null || true)"
		fi
		if [ -n "${existing}" ] && kill -0 "${existing}" >/dev/null 2>&1; then
			exit 0
		fi
		# A flock(2) lock lives on the open file description and is released only when
		# every duplicate fd, a forked child's copy included, is closed
		# (https://man7.org/linux/man-pages/man2/flock.2.html), so 9>&- keeps the lock fd
		# out of the dock's process.
		setsid nohup /usr/local/bin/box-bounded-log --run "${log}" -- "$@" \
			>/dev/null 2>&1 9>&- &
		echo "$!" >"${pid_file}"
	) 9>"${lock_file}"
}

configure_plank_dock dock1

start_dock_once "/tmp/plank${DISPLAY}-dock1.pid" "/tmp/plank${DISPLAY}-dock1.log" \
	/usr/local/bin/box-plank --name dock1
sand_desktop_register "${SAND_DESKTOP_GROUP}" dock 5 "/tmp/plank${DISPLAY}-dock1.log" \
	"$(tr -dc '0-9' <"/tmp/plank${DISPLAY}-dock1.pid" 2>/dev/null || true)" \
	"/tmp/plank${DISPLAY}-dock1.pid" -- \
	/usr/local/bin/box-bounded-log --run "/tmp/plank${DISPLAY}-dock1.log" -- \
	/usr/local/bin/box-plank --name dock1

mkdir -p /workspace
cd /workspace

tail -f /dev/null
