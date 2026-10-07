#!/usr/bin/env bash
# Publication automatique du monitor sur le repo GitHub défini dans .publish.conf (local, gitignoré),
# dossier box-update-monitor/.
# Garde-fou BLOQUANT : scan données perso / secrets sur exactement ce qui serait poussé.
#
# Usage : publish.sh [--dry-run]
# Codes de sortie :
#   0  publié (affiche "PUBLISHED commit=<sha> ...")  ou rien à publier ("NOTHING_TO_PUBLISH")
#      ou --dry-run propre ("DRY_RUN_CLEAN ...")
#   2  BLOQUÉ par le scan -> rien n'est poussé ; reports/PUBLISH-BLOCKED.flag écrit (chemins + n° de règle,
#      jamais la valeur perso brute)
#   3  verrou occupé (snapshot/check en cours)       1  autre erreur (git, réseau...)
#
# Clone de travail persistant : $PUBLISH_REPO_DIR (défaut /workspace/.box-update-monitor-publish-repo)
set -uo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
CONF="$DIR/.publish.conf"   # REPO_SLUG, BRANCH, GIT_NAME, GIT_EMAIL (local : jamais publié)
[[ -f "$CONF" ]] || { echo "ERROR: $CONF absent" >&2; exit 1; }
# shellcheck disable=SC1090
source "$CONF"
: "${REPO_SLUG:?}" "${GIT_NAME:?}" "${GIT_EMAIL:?}"; BRANCH="${BRANCH:-main}"
SUBDIR="box-update-monitor"
CLONE="${PUBLISH_REPO_DIR:-/workspace/.box-update-monitor-publish-repo}"
BLOCK_FLAG="$DIR/reports/PUBLISH-BLOCKED.flag"
LOG="$DIR/logs/publish.log"
DRY=0; [[ "${1:-}" == "--dry-run" ]] && DRY=1
mkdir -p "$DIR/logs" "$DIR/reports"
log(){ echo "[$(date '+%F %T %Z')] $*" >> "$LOG"; }
die(){ echo "ERROR: $*" >&2; log "ERROR $*"; exit 1; }

# Compte GitHub propriétaire du repo (le GH_TOKEN par défaut de la box = un autre compte, sans accès)
unset GH_TOKEN GITHUB_TOKEN
# git-lfs absent de la box : neutraliser le filtre (on ne touche jamais aux fichiers LFS de non-stock/)
G=(git -C "$CLONE" -c filter.lfs.process= -c filter.lfs.required=false -c filter.lfs.smudge=cat -c filter.lfs.clean=cat)

# 4. verrou partagé avec snapshot.sh / check-update.sh
exec 9>"$DIR/.lock"
flock -w 600 9 || { echo "LOCKED: snapshot/check en cours" >&2; log "LOCKED"; exit 3; }

# clone persistant
if [[ ! -d "$CLONE/.git" ]]; then
  GIT_LFS_SKIP_SMUDGE=1 git -c filter.lfs.process= -c filter.lfs.required=false -c filter.lfs.smudge=cat \
    clone -q "https://github.com/$REPO_SLUG.git" "$CLONE" || die "clone impossible"
  "${G[@]}" config filter.lfs.process ""; "${G[@]}" config filter.lfs.required false
  "${G[@]}" config filter.lfs.smudge cat;  "${G[@]}" config filter.lfs.clean cat
fi
"${G[@]}" config user.name "$GIT_NAME"
"${G[@]}" config user.email "$GIT_EMAIL"
gh auth setup-git >/dev/null 2>&1 || true

reset_clone(){ "${G[@]}" reset -q --hard "origin/$BRANCH" && "${G[@]}" clean -fdq -- "$SUBDIR"; }
"${G[@]}" fetch -q origin "$BRANCH" || die "fetch impossible"
reset_clone || die "reset du clone impossible"

# miroir de l'état publiable (le .gitignore du dossier s'applique en plus)
rm -rf "$CLONE/$SUBDIR"; mkdir -p "$CLONE/$SUBDIR"
( cd "$DIR" && tar --exclude=./objects --exclude=./logs --exclude=./.lock --exclude=./.pii-denylist --exclude=./.publish.conf \
    --exclude=./reports/NOUVEAU-DIFF.flag --exclude=./reports/PUBLISH-BLOCKED.flag --exclude='*.partial' \
    -cf - . ) | ( cd "$CLONE/$SUBDIR" && tar -xf - ) || die "copie impossible"
touch "$CLONE/$SUBDIR/reports/.gitkeep"
"${G[@]}" add -A -- "$SUBDIR" || die "git add impossible"

mapfile -d '' CHANGED < <("${G[@]}" diff --cached --name-only -z --diff-filter=ACMR -- "$SUBDIR")
N_ALL=$("${G[@]}" diff --cached --name-only -- "$SUBDIR" | wc -l)
if [[ "$N_ALL" -eq 0 ]]; then
  reset_clone
  echo "NOTHING_TO_PUBLISH (dépôt déjà à jour : $("${G[@]}" rev-parse --short HEAD))"
  log "nothing to publish"
  exit 0
fi

# 1. scan sur EXACTEMENT ce qui serait poussé (fichiers ajoutés/modifiés)
STAGE="$(mktemp -d /tmp/bum-publish-scan.XXXXXX)"
trap 'rm -rf "$STAGE"' EXIT
for f in "${CHANGED[@]}"; do
  mkdir -p "$STAGE/$(dirname "$f")"; cp -p "$CLONE/$f" "$STAGE/$f"
done
mkdir -p "$STAGE/.empty"   # cible jamais vide
SCAN_OUT="$("$DIR/scan-personal-data.sh" --redact "$STAGE" 2>&1)"; SCAN_RC=$?
SCAN_OUT="${SCAN_OUT//$STAGE\//}"

if [[ $SCAN_RC -ne 0 ]]; then
  # 2. BLOQUÉ : rien n'est poussé
  reset_clone
  {
    echo "PUBLISH BLOQUÉ — $(date '+%F %T %Z')"
    echo "Fichiers candidats : ${#CHANGED[@]} (rien n'a été commité ni poussé)"
    echo "Détails (chemins relatifs au repo : nb d'occurrences ; règle = n° de ligne active de .pii-denylist ; valeurs masquées) :"
    echo "$SCAN_OUT"
    echo
    echo "Action : retirer/masquer la donnée dans /workspace/box-update-monitor (ou exclure le chemin),"
    echo "relancer publish.sh, puis supprimer ce fichier."
  } > "$BLOCK_FLAG"
  log "BLOCKED rc=$SCAN_RC files=${#CHANGED[@]}"
  echo "BLOCKED flag=$BLOCK_FLAG"
  echo "$SCAN_OUT" | grep -E '^## ' >&2
  exit 2
fi

# 3. message de commit
NEW_SNAPS=$("${G[@]}" diff --cached --name-only --diff-filter=A -- "$SUBDIR/snapshots" | sed -n "s#^$SUBDIR/snapshots/\([^/]*\)/SNAPSHOT-INFO.txt\$#\1#p")
NEW_REPORTS=$("${G[@]}" diff --cached --name-only --diff-filter=AM -- "$SUBDIR/reports" | grep -E '\.md$' || true)
MSG=""
for s in $NEW_SNAPS; do
  label=$(sed -n 's/^label\t//p' "$DIR/snapshots/$s/SNAPSHOT-INFO.txt")
  MSG+="snapshot $s${label:+ ($label)}; "
done
for r in $NEW_REPORTS; do
  rf="$CLONE/$r"
  a=$(sed -n 's/^| Ajoutés | \([0-9]*\) |$/\1/p' "$rf"); d=$(sed -n 's/^| Supprimés | \([0-9]*\) |$/\1/p' "$rf")
  m=$(sed -n 's/^| Modifiés (hash) | \([0-9]*\) |$/\1/p' "$rf")
  MSG+="diff $(basename "$r" .md) +${a:-?}/-${d:-?}/~${m:-?}; "
done
if [[ -z "$MSG" ]]; then
  MSG="box-update-monitor: update $(echo "${CHANGED[@]}" | tr ' ' '\n' | sed "s#^$SUBDIR/##" | head -6 | paste -sd, -)"
else
  MSG="box-update-monitor: ${MSG%; }"
fi
MSG="${MSG:0:200}"

if [[ $DRY -eq 1 ]]; then
  echo "DRY_RUN_CLEAN scan=PROPRE files=$N_ALL"
  echo "commit (non créé) : $MSG"
  "${G[@]}" diff --cached --name-status -- "$SUBDIR" | head -20
  reset_clone
  log "dry-run clean files=$N_ALL"
  exit 0
fi

"${G[@]}" commit -q -m "$MSG" -m "Auto-publié par publish.sh après scan données perso (PROPRE). $N_ALL fichier(s)." \
  || die "commit impossible"
if ! "${G[@]}" push -q origin "HEAD:$BRANCH" 2>>"$LOG"; then
  "${G[@]}" pull -q --rebase origin "$BRANCH" && "${G[@]}" push -q origin "HEAD:$BRANCH" 2>>"$LOG" \
    || { reset_clone; die "push refusé (voir $LOG)"; }
fi
SHA=$("${G[@]}" rev-parse HEAD)
[[ -f "$BLOCK_FLAG" ]] && { mv "$BLOCK_FLAG" "$DIR/logs/PUBLISH-BLOCKED.resolved.$(date +%Y%m%d-%H%M%S)"; }
log "PUBLISHED $SHA files=$N_ALL msg=$MSG"
echo "PUBLISHED commit=$SHA files=$N_ALL"
echo "message: $MSG"
echo "url: https://github.com/$REPO_SLUG/commit/$SHA"
exit 0
