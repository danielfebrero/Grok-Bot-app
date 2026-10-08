# shellcheck shell=bash

SAND_DESKTOP_DIR="${SAND_DESKTOP_DIR:-/tmp/sand-desktop}"

# >>> desktop supervision predicate (from packages/constants/src/sand-supervisor.ts; regenerate: pnpm --filter sand run gen:box-ports) >>>
sand_desktop_supervision_enabled() {
	case "$(printf '%s' "${SAND_DESKTOP_SUPERVISION_DISABLED:-}" | tr '[:upper:]' '[:lower:]' | tr -d '[:space:]')" in
	1 | true | yes) return 1 ;;
	*) return 0 ;;
	esac
}
# <<< desktop supervision predicate <<<

sand_desktop_register() {
	sand_desktop_supervision_enabled || return 0
	command -v jq >/dev/null 2>&1 || return 0
	local group="$1" name="$2" order="$3" logfile="$4" pid="$5"
	shift 5
	local pidfile="" listen_port=""
	if [ "${1:-}" != "--" ] && [ "${1:-}" != "--listen-port" ]; then
		pidfile="$1"
		shift
	fi
	if [ "${1:-}" = "--listen-port" ]; then
		[ "$#" -ge 2 ] || return 0
		listen_port="$2"
		shift 2
		case "${listen_port}" in
		'' | *[!0-9]*) return 0 ;;
		esac
	fi
	[ "${1:-}" = "--" ] && shift
	[ "$#" -ge 1 ] || return 0
	case "${pid}" in
	'' | *[!0-9]*) return 0 ;;
	esac
	local dir="${SAND_DESKTOP_DIR}/${group}"
	mkdir -p "${dir}" 2>/dev/null || return 0
	[ -n "${pidfile}" ] || pidfile="${dir}/${name}.pid"
	local argv_json
	argv_json="$(printf '%s\0' "$@" | jq -Rs 'split("\u0000")[:-1]' 2>/dev/null)" || return 0
	[ -n "${argv_json}" ] || return 0
	local tmp="${dir}/.${name}.json.$$"
	if jq -n \
		--arg name "${name}" \
		--arg group "${group}" \
		--argjson order "${order}" \
		--arg logfile "${logfile}" \
		--arg pidfile "${pidfile}" \
		--argjson listen_port "${listen_port:-null}" \
		--arg display "${DISPLAY:-}" \
		--arg home "${HOME:-}" \
		--arg xdg "${XDG_RUNTIME_DIR:-}" \
		--arg dbus "${DBUS_SESSION_BUS_ADDRESS:-}" \
		--argjson argv "${argv_json}" \
		'({name: $name, group: $group, order: $order, logFile: $logfile, pidFile: $pidfile,
		  env: ({DISPLAY: $display, HOME: $home, XDG_RUNTIME_DIR: $xdg, DBUS_SESSION_BUS_ADDRESS: $dbus}
		        | with_entries(select(.value != ""))),
		  argv: $argv}
		  + if $listen_port == null then {} else {listenPort: $listen_port} end)' >"${tmp}" 2>/dev/null; then
		mv -f "${tmp}" "${dir}/${name}.json" 2>/dev/null || rm -f "${tmp}"
		printf '%s' "${pid}" >"${pidfile}" 2>/dev/null || true
	else
		rm -f "${tmp}" 2>/dev/null || true
	fi
}

sand_desktop_unregister_group() {
	local group="$1"
	[ -n "${group}" ] || return 0
	rm -rf "${SAND_DESKTOP_DIR:?}/${group}" 2>/dev/null || true
}
