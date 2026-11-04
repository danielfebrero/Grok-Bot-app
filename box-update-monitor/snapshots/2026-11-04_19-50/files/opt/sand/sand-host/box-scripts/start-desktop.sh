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
# Copy-in may restore packages into the user site, and websockify is Python.
export PYTHONNOUSERSITE=1
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

BOX_DISPLAY_NUM="${DISPLAY#:}"
BOX_DISPLAY_NUM="${BOX_DISPLAY_NUM%%.*}"
if [ "${BOX_DISPLAY_NUM:-1}" -ge 2 ] 2>/dev/null; then
	BOX_USER_XDG_DIR="/tmp/xdg-runtime-box-${BOX_DISPLAY_NUM}"
else
	BOX_USER_XDG_DIR="/tmp/xdg-runtime-box"
fi
if [ "$(id -u)" -eq 0 ]; then
	if [ -z "${DBUS_SESSION_BUS_ADDRESS:-}" ]; then
		eval "$(dbus-launch --sh-syntax)"
	fi
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
	if [ -z "${DBUS_SESSION_BUS_ADDRESS:-}" ]; then
		# Each daemon start listens under a new guid (https://dbus.freedesktop.org/doc/dbus-specification.html#addresses),
		# so the published address names only the socket, which a relaunched bus reuses.
		export DBUS_SESSION_BUS_ADDRESS="unix:path=${BOX_USER_XDG_DIR}/bus"
		BOX_DBUS_ARGV=(/usr/local/bin/box-dbus --session --nofork --address="${DBUS_SESSION_BUS_ADDRESS}")
		/usr/local/bin/box-bounded-log --run "/tmp/dbus${DISPLAY}.log" -- "${BOX_DBUS_ARGV[@]}" &
		BOX_DBUS_PID=$!
		sand_desktop_register "${SAND_DESKTOP_GROUP}" bus -1 "/tmp/dbus${DISPLAY}.log" "${BOX_DBUS_PID}" -- \
			/usr/local/bin/box-bounded-log --run "/tmp/dbus${DISPLAY}.log" -- "${BOX_DBUS_ARGV[@]}"
		for _ in $(seq 1 50); do
			BOX_DBUS_OWNER="$(dbus-send --session --print-reply=literal --dest=org.freedesktop.DBus /org/freedesktop/DBus \
				org.freedesktop.DBus.GetConnectionUnixProcessID string:org.freedesktop.DBus 2>/dev/null)" || BOX_DBUS_OWNER=""
			BOX_DBUS_OWNER="${BOX_DBUS_OWNER##* }"
			if [ -n "${BOX_DBUS_OWNER}" ] && { [ "${BOX_DBUS_OWNER}" = "${BOX_DBUS_PID}" ] || ! kill -0 "${BOX_DBUS_PID}" 2>/dev/null; }; then
				break
			fi
			sleep 0.1
		done
	fi
	printf '%s' "${DBUS_SESSION_BUS_ADDRESS:-}" >"${BOX_USER_XDG_DIR}/dbus-session-address" 2>/dev/null || true
fi

rm -f /tmp/xdg-runtime-box/sand-gateway-credential /tmp/xdg-runtime-box-*/sand-gateway-credential || true

SAND_DESKTOP_PANEL_MARKER="${BOX_USER_XDG_DIR}/sand-desktop-panel"
if [ "${SAND_BOX_DESKTOP_PANEL:-}" = "1" ] && command -v xfce4-panel >/dev/null 2>&1; then
	: >"${SAND_DESKTOP_PANEL_MARKER}" || true
else
	rm -f "${SAND_DESKTOP_PANEL_MARKER}" || true
fi

# Persistent user keyring (sand_box_persistent_keyring gate, stamped as
# SAND_BOX_PERSISTENT_KEYRING=1 at pod create): unlock the on-disk `login`
# GNOME keyring on the box session bus. box-keyring-unlock execs
# gnome-keyring-daemon in the foreground, so it runs as a supervised desktop
# component like the rest of this session: the supervisor restarts it when it
# dies (the image removes Secret Service D-Bus activation, so nothing else
# would), and the script replaces the daemon from any earlier session before
# it starts, so two daemons never share the keyring files. Only on the primary
# display: the keyring files are shared on disk, and a second daemon on a
# forked display's bus would race the first. No-op when the env is absent
# (today's default) or the image lacks gnome-keyring: host-bundle sync can
# land this script on an image without the daemon, and registering it there
# would make the supervisor relaunch its clean no-op exit forever, so the
# component is not registered at all until the image ships the binary.
start_persistent_keyring() {
	[ "${SAND_BOX_PERSISTENT_KEYRING:-}" = "1" ] || return 0
	[ "${BOX_DISPLAY_NUM:-1}" -lt 2 ] 2>/dev/null || return 0
	[ -x /usr/local/bin/box-keyring-unlock ] || return 0
	command -v gnome-keyring-daemon >/dev/null 2>&1 || return 0
	local bus="${DBUS_SESSION_BUS_ADDRESS:-}" runner=()
	if [ "$(id -u)" -eq 0 ]; then
		bus="${BOX_DBUS_ADDRESS:-}"
		runner=(runuser -u box --)
	fi
	[ -n "${bus}" ] || return 0
	local cmd=(
		"${runner[@]}" env
		HOME=/home/box
		XDG_RUNTIME_DIR="${BOX_USER_XDG_DIR}"
		DBUS_SESSION_BUS_ADDRESS="${bus}"
		/usr/local/bin/box-bounded-log --run /tmp/box-keyring.log --
		/usr/local/bin/box-keyring-unlock
	)
	"${cmd[@]}" &
	sand_desktop_register "${SAND_DESKTOP_GROUP}" keyring 6 /tmp/box-keyring.log "$!" -- "${cmd[@]}"
}
start_persistent_keyring

# Audio devices (sand_box_audio_devices gate, stamped as SAND_BOX_AUDIO_DEVICES=1
# at pod create): a PulseAudio daemon per display in the box runtime dir that
# box-chrome hands this display's Chrome. It starts before Xvfb so its socket is
# up before any Chrome on the display can start, because Chrome picks its audio
# backend once at startup
# (https://chromium.googlesource.com/chromium/src/+/refs/tags/156.0.8078.47/media/audio/linux/audio_manager_linux.cc).
# Registers nothing on an image without pulseaudio, like the keyring above.
start_box_audio() {
	[ "${SAND_BOX_AUDIO_DEVICES:-}" = "1" ] || return 0
	[ -x /usr/local/bin/box-audio.sh ] || return 0
	command -v pulseaudio >/dev/null 2>&1 || return 0
	local log="/tmp/pulseaudio${DISPLAY}.log"
	local cmd=(/usr/local/bin/box-bounded-log --run "${log}" -- /usr/local/bin/box-audio.sh "${BOX_USER_XDG_DIR}")
	if [ "$(id -u)" -eq 0 ]; then
		cmd=(runuser -u box -- env HOME=/home/box "${cmd[@]}")
	fi
	"${cmd[@]}" &
	sand_desktop_register "${SAND_DESKTOP_GROUP}" audio 7 "${log}" "$!" -- "${cmd[@]}"
	# box-audio never replaces a live daemon, so a socket that accepts a
	# connection is served by the daemon that keeps serving it, and a socket
	# file a dead daemon left in a kept /tmp refuses connections
	# (https://man7.org/linux/man-pages/man7/unix.7.html).
	for _ in $(seq 1 50); do
		python3 -I -c 'import socket, sys; socket.socket(socket.AF_UNIX).connect(sys.argv[1])' \
			"${BOX_USER_XDG_DIR}/pulse/native" 2>/dev/null && return 0
		sleep 0.1
	done
	echo "start-desktop: PulseAudio for ${DISPLAY} did not accept connections within 5 s; Chrome started on this display now gets no audio devices (see ${log})" >&2
}
start_box_audio

/usr/local/bin/box-bounded-log --run "/tmp/xvfb${DISPLAY}.log" -- \
	/usr/local/bin/box-xvfb "${DISPLAY}" -screen 0 "${SCREEN_WIDTH}x${SCREEN_HEIGHT}x24" -ac +extension GLX +render -noreset &
sand_desktop_register "${SAND_DESKTOP_GROUP}" xvfb 0 "/tmp/xvfb${DISPLAY}.log" "$!" -- \
	/usr/local/bin/box-bounded-log --run "/tmp/xvfb${DISPLAY}.log" -- \
	/usr/local/bin/box-xvfb "${DISPLAY}" -screen 0 "${SCREEN_WIDTH}x${SCREEN_HEIGHT}x24" -ac +extension GLX +render -noreset

DISPLAY_READY=false
for _ in $(seq 1 150); do
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
	-norc
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
	--config /dev/null
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

# xfce4-panel reads its layout from the xfce4-panel xfconf channel, and a panel
# that is never hidden reserves its size along its edge as a strut, which the
# window manager subtracts from the work area it maximizes windows into
# (https://docs.xfce.org/xfce/xfce4-panel/preferences,
# https://specifications.freedesktop.org/wm-spec/1.5/ar01s05.html#id-1.6.10).
# Each launcher plugin starts the .desktop items kept in its own launcher-<id>
# folder (https://docs.xfce.org/xfce/xfce4-panel/launcher).
configure_desktop_panel() {
	local channel_dir="${HOME}/.config/xfce4/xfconf/xfce-perchannel-xml"
	local launchers="${HOME}/.config/xfce4/panel" id=2 item
	mkdir -p "${channel_dir}" || return 1
	for item in \
		"${HOME}/.local/share/applications/box-chrome.desktop" \
		/usr/share/applications/thunar.desktop \
		/usr/share/applications/xfce4-terminal.desktop; do
		mkdir -p "${launchers}/launcher-${id}" &&
			cp -f "${item}" "${launchers}/launcher-${id}/" 2>/dev/null || true
		id=$((id + 1))
	done
	cat >"${channel_dir}/.xfce4-panel.xml.$$" <<'EOF' || return 1
<?xml version="1.0" encoding="UTF-8"?>
<channel name="xfce4-panel" version="1.0">
  <property name="configver" type="int" value="2"/>
  <property name="panels" type="array">
    <value type="int" value="1"/>
    <property name="panel-1" type="empty">
      <property name="position" type="string" value="p=10;x=0;y=0"/>
      <property name="position-locked" type="bool" value="true"/>
      <property name="length" type="double" value="100"/>
      <property name="size" type="uint" value="40"/>
      <property name="icon-size" type="uint" value="24"/>
      <property name="autohide-behavior" type="uint" value="0"/>
      <property name="enable-struts" type="bool" value="true"/>
      <property name="plugin-ids" type="array">
        <value type="int" value="1"/>
        <value type="int" value="2"/>
        <value type="int" value="3"/>
        <value type="int" value="4"/>
        <value type="int" value="5"/>
        <value type="int" value="6"/>
        <value type="int" value="7"/>
        <value type="int" value="8"/>
      </property>
    </property>
  </property>
  <property name="plugins" type="empty">
    <property name="plugin-1" type="string" value="applicationsmenu"/>
    <property name="plugin-2" type="string" value="launcher">
      <property name="items" type="array">
        <value type="string" value="box-chrome.desktop"/>
      </property>
    </property>
    <property name="plugin-3" type="string" value="launcher">
      <property name="items" type="array">
        <value type="string" value="thunar.desktop"/>
      </property>
    </property>
    <property name="plugin-4" type="string" value="launcher">
      <property name="items" type="array">
        <value type="string" value="xfce4-terminal.desktop"/>
      </property>
    </property>
    <property name="plugin-5" type="string" value="separator"/>
    <property name="plugin-6" type="string" value="tasklist"/>
    <property name="plugin-7" type="string" value="separator">
      <property name="expand" type="bool" value="true"/>
      <property name="style" type="uint" value="0"/>
    </property>
    <property name="plugin-8" type="string" value="clock"/>
  </property>
</channel>
EOF
	mv -f "${channel_dir}/.xfce4-panel.xml.$$" "${channel_dir}/xfce4-panel.xml"
}

start_desktop_panel() {
	[ -e "${SAND_DESKTOP_PANEL_MARKER:-}" ] || return 1
	if ! configure_desktop_panel; then
		rm -f "${SAND_DESKTOP_PANEL_MARKER}"
		return 1
	fi
	start_dock_once "/tmp/xfce4-panel${DISPLAY}.pid" "/tmp/xfce4-panel${DISPLAY}.log" xfce4-panel
	sand_desktop_register "${SAND_DESKTOP_GROUP}" dock 5 "/tmp/xfce4-panel${DISPLAY}.log" \
		"$(tr -dc '0-9' <"/tmp/xfce4-panel${DISPLAY}.pid" 2>/dev/null || true)" \
		"/tmp/xfce4-panel${DISPLAY}.pid" -- \
		/usr/local/bin/box-bounded-log --run "/tmp/xfce4-panel${DISPLAY}.log" -- xfce4-panel
}

if ! start_desktop_panel; then
	configure_plank_dock dock1

	start_dock_once "/tmp/plank${DISPLAY}-dock1.pid" "/tmp/plank${DISPLAY}-dock1.log" \
		/usr/local/bin/box-plank --name dock1
	sand_desktop_register "${SAND_DESKTOP_GROUP}" dock 5 "/tmp/plank${DISPLAY}-dock1.log" \
		"$(tr -dc '0-9' <"/tmp/plank${DISPLAY}-dock1.pid" 2>/dev/null || true)" \
		"/tmp/plank${DISPLAY}-dock1.pid" -- \
		/usr/local/bin/box-bounded-log --run "/tmp/plank${DISPLAY}-dock1.log" -- \
		/usr/local/bin/box-plank --name dock1
fi

mkdir -p /workspace
cd /workspace

tail -f /dev/null
