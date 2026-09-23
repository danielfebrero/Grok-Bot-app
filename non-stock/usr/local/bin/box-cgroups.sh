# shellcheck shell=bash

SAND_CGROUP_ROOT="${SAND_CGROUP_ROOT:-/sys/fs/cgroup}"
SAND_CGROUP_INTERACTIVE_NAME="interactive"
SAND_CGROUP_AGENT_NAME="agent"

sand_cgroups_enabled() {
	case "$(printf '%s' "${SAND_BOX_CGROUPS_DISABLED:-}" | tr '[:upper:]' '[:lower:]' | tr -d '[:space:]')" in
	1 | true | yes) return 1 ;;
	*) return 0 ;;
	esac
}

sand_cgroup_log() {
	echo "[box-cgroups] $*"
}

sand_cgroup_write() {
	local value="$1" path="$2"
	printf '%s' "${value}" >"${path}" 2>/dev/null || return 1
	return 0
}

sand_cgroup_v2_cpu_available() {
	local controllers="${SAND_CGROUP_ROOT}/cgroup.controllers"
	[ -r "${controllers}" ] || return 1
	grep -qw cpu "${controllers}" 2>/dev/null || return 1
	return 0
}

# A cgroup created under a threaded cgroup is "domain invalid": it cannot be
# populated or have controllers enabled until its cgroup.type is written
# "threaded", and operations on it fail with EOPNOTSUPP
# (https://docs.kernel.org/admin-guide/cgroup-v2.html#threads), so setup is
# skipped when the box's own cgroup is threaded (the dev-box case).
sand_cgroup_is_threaded() {
	local type_file="${SAND_CGROUP_ROOT}/cgroup.type"
	[ -r "${type_file}" ] || return 1
	case "$(cat "${type_file}" 2>/dev/null)" in
	threaded) return 0 ;;
	*) return 1 ;;
	esac
}

# Non-root cgroups can enable a domain controller in cgroup.subtree_control
# only while they hold no processes of their own
# (https://docs.kernel.org/admin-guide/cgroup-v2.html#no-internal-process-constraint),
# and the box's cgroup, shown as "/" inside its cgroup namespace, is the cgroupns
# root and a non-root cgroup to the kernel (https://docs.kernel.org/admin-guide/cgroup-v2.html#namespace),
# so sand_cgroup_migrate_root_procs moves its processes to a leaf before +cpu is written.
sand_cgroup_migrate_root_procs() {
	local group="$1"
	local dest="${SAND_CGROUP_ROOT}/${group}/cgroup.procs"
	local src="${SAND_CGROUP_ROOT}/cgroup.procs"
	[ -r "${src}" ] || return 0
	[ -d "${SAND_CGROUP_ROOT}/${group}" ] || return 0
	local pid
	# cgroup.procs is a live listing whose PIDs are unordered and may repeat when a
	# process moves during the read (https://docs.kernel.org/admin-guide/cgroup-v2.html#core-interface-files),
	# so the list is snapshotted before the migration writes.
	local pids
	pids="$(cat "${src}" 2>/dev/null || true)"
	for pid in ${pids}; do
		case "${pid}" in
		'' | *[!0-9]*) continue ;;
		esac
		printf '%s' "${pid}" >"${dest}" 2>/dev/null || true
	done
	return 0
}

sand_cgroup_apply_weight() {
	local group="$1" weight="$2"
	[ -n "${weight}" ] || return 0
	case "${weight}" in
	'' | *[!0-9]*)
		sand_cgroup_log "ignoring non-numeric cpu.weight '${weight}' for ${group}"
		return 0
		;;
	esac
	if [ "${weight}" -lt 1 ] || [ "${weight}" -gt 10000 ]; then
		sand_cgroup_log "ignoring out-of-range cpu.weight ${weight} for ${group} (1..10000)"
		return 0
	fi
	if sand_cgroup_write "${weight}" "${SAND_CGROUP_ROOT}/${group}/cpu.weight"; then
		sand_cgroup_log "${group}: cpu.weight=${weight}"
	fi
	return 0
}

sand_cgroup_setup() {
	sand_cgroups_enabled || {
		sand_cgroup_log "disabled by SAND_BOX_CGROUPS_DISABLED; skipping"
		return 0
	}
	if ! sand_cgroup_v2_cpu_available; then
		sand_cgroup_log "no cgroup v2 cpu controller at ${SAND_CGROUP_ROOT}; skipping (box runs unpartitioned)"
		return 0
	fi
	if sand_cgroup_is_threaded; then
		sand_cgroup_log "cgroup at ${SAND_CGROUP_ROOT} is threaded (dev box on a cloud VM); skipping (box runs unpartitioned)"
		return 0
	fi
	local group
	for group in "${SAND_CGROUP_INTERACTIVE_NAME}" "${SAND_CGROUP_AGENT_NAME}"; do
		if ! mkdir -p "${SAND_CGROUP_ROOT}/${group}" 2>/dev/null; then
			sand_cgroup_log "cannot create ${SAND_CGROUP_ROOT}/${group} (read-only cgroupfs?); skipping"
			return 0
		fi
	done
	sand_cgroup_migrate_root_procs "${SAND_CGROUP_AGENT_NAME}"
	if ! sand_cgroup_write "+cpu" "${SAND_CGROUP_ROOT}/cgroup.subtree_control"; then
		sand_cgroup_log "could not enable +cpu in cgroup.subtree_control; leaves exist but carry no cpu accounting"
	fi
	sand_cgroup_apply_weight "${SAND_CGROUP_INTERACTIVE_NAME}" "${SAND_CGROUP_INTERACTIVE_WEIGHT:-}"
	sand_cgroup_apply_weight "${SAND_CGROUP_AGENT_NAME}" "${SAND_CGROUP_AGENT_WEIGHT:-}"
	sand_cgroup_log "ready: cpu.stat=$(
		[ -r "${SAND_CGROUP_ROOT}/${SAND_CGROUP_INTERACTIVE_NAME}/cpu.stat" ] && echo yes || echo no
	) cpu.pressure=$(
		[ -r "${SAND_CGROUP_ROOT}/${SAND_CGROUP_INTERACTIVE_NAME}/cpu.pressure" ] && echo yes || echo no
	) interactive.weight=$(
		cat "${SAND_CGROUP_ROOT}/${SAND_CGROUP_INTERACTIVE_NAME}/cpu.weight" 2>/dev/null || echo unset
	) agent.weight=$(
		cat "${SAND_CGROUP_ROOT}/${SAND_CGROUP_AGENT_NAME}/cpu.weight" 2>/dev/null || echo unset
	)"
	return 0
}

sand_cgroup_place() {
	local group="$1" pid="$2"
	sand_cgroups_enabled || return 0
	case "${pid}" in
	'' | *[!0-9]*) return 0 ;;
	esac
	[ -d "${SAND_CGROUP_ROOT}/${group}" ] || return 0
	{ printf '%s' "${pid}" >"${SAND_CGROUP_ROOT}/${group}/cgroup.procs"; } 2>/dev/null || true
	return 0
}

sand_cgroup_join() {
	local group="$1"
	sand_cgroup_place "${group}" "$$"
	return 0
}
