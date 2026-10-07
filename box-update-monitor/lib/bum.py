#!/usr/bin/env python3
"""box-update-monitor — snapshots + diffs du périmètre app Grok Bot sur la box.

Sous-commandes :
  snapshot [--label TXT]              crée snapshots/<YYYY-MM-DD_HH-MM>/
  diff [OLD NEW]                      rapport Markdown (défaut : 2 snapshots les plus récents)
  check [--deep] [--quiet]            détecte un changement depuis le dernier snapshot ;
                                      si changement : snapshot + diff + reports/NOUVEAU-DIFF.flag
  fingerprint                         affiche l'empreinte rapide courante
  versions                            affiche les versions/build courantes
"""
import fnmatch, gzip, hashlib, json, os, re, shutil, stat, subprocess, sys, tempfile, time
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SNAP_DIR = os.path.join(ROOT, "snapshots")
REPORT_DIR = os.path.join(ROOT, "reports")
OBJ_DIR = os.path.join(ROOT, "objects")
LOG_DIR = os.path.join(ROOT, "logs")
PERIMETER = os.path.join(ROOT, "perimeter.conf")
FLAG = os.path.join(REPORT_DIR, "NOUVEAU-DIFF.flag")
VERSION_FILE = os.path.join(ROOT, "VERSION-actuelle")

TEXT_COPY_MAX = 2 * 1024 * 1024        # texte <= 2 Mo : copie brute (diff direct)
BIG_TEXT_MAX = 64 * 1024 * 1024        # texte 2–64 Mo (bundles JS) : copie .gz
INLINE_DIFF_LINES = 150                # lignes de diff incluses dans le .md par fichier
LONG_LINE = 1000                       # au-delà : fichier minifié -> normalisation avant diff

# Fichiers clés pour la détection rapide (check-update)
KEY_FILES = [
    "/etc/sand-box-image-sha",
    "/usr/local/share/sand-box-scripts/version",
    "/opt/sand/sand-host/version",
    "/opt/sand/sand-host/host-main.cjs",
    "/opt/sand/sand-host/sand-eval-runner.cjs",
    "/opt/sand/deps/runtime-deps-manifest.json",
    "/exec-daemon/package.json",
    "/exec-daemon/index.js",
    "/exec-daemon/agent-sdk/canvas-sdk-version",
    "/usr/local/bin/start-sand-box",
    "/usr/local/bin/sand-supervisor.mjs",
    "/usr/local/bin/sand-supervisor-contract.mjs",
    "/usr/local/bin/box-doctor",
    "/usr/local/bin/cursor_agent_store_fuse_version",
    "/usr/local/libexec/sand/sand-daemon-supervisor",
    "/var/lib/sand-daemon-supervisor/orbitd/bundled/manifest.json",
]
KEY_TREES = [  # contenu hashé (petits arbres)
    "/home/box/reference",
    "/home/box/sand-data/managed-skills",
    "/var/lib/sand-daemon-supervisor/orbitd/releases",
]


def now_local():
    return datetime.now().astimezone()


def log(msg):
    os.makedirs(LOG_DIR, exist_ok=True)
    line = f"[{now_local().strftime('%Y-%m-%d %H:%M:%S %Z')}] {msg}"
    with open(os.path.join(LOG_DIR, "monitor.log"), "a") as f:
        f.write(line + "\n")


def read_small(path, limit=4096):
    try:
        with open(path, "rb") as f:
            return f.read(limit).decode("utf-8", "replace").strip()
    except Exception as e:
        return f"<absent: {type(e).__name__}>"


def run(cmd, timeout=20):
    try:
        return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout).stdout.strip()
    except Exception as e:
        return f"<erreur: {type(e).__name__}>"


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def load_perimeter():
    roots, excludes = [], []
    with open(PERIMETER) as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            mode, _, path = line.partition(" ") if " " in line else line.partition("\t")
            mode, path = mode.strip(), path.strip()
            if mode == "exclude":
                excludes.append(path)
            elif mode in ("full", "hash"):
                roots.append((mode, path))
    return roots, excludes


def excluded(path, excludes):
    return any(fnmatch.fnmatch(path, pat) for pat in excludes)


def walk(root, excludes):
    """Yield chemins (fichiers + symlinks) sous root, sans suivre les symlinks de dossiers."""
    if excluded(root, excludes):
        return
    if os.path.islink(root) or os.path.isfile(root):
        yield root
        return
    if not os.path.isdir(root):
        return
    for dirpath, dirnames, filenames in os.walk(root, onerror=lambda e: None):
        keep = []
        for d in sorted(dirnames):
            p = os.path.join(dirpath, d)
            if excluded(p, excludes):
                continue
            if os.path.islink(p):
                yield p          # symlink vers dossier : listé comme entrée, non suivi
            else:
                keep.append(d)
        dirnames[:] = keep
        for fn in sorted(filenames):
            p = os.path.join(dirpath, fn)
            if not excluded(p, excludes):
                yield p


def is_text(path, size):
    try:
        with open(path, "rb") as f:
            head = f.read(8192)
    except Exception:
        return False
    if b"\x00" in head:
        return False
    if not head:
        return True
    try:
        head.decode("utf-8")
        return True
    except UnicodeDecodeError as e:
        return e.start > len(head) - 4  # coupure multi-octets en fin de bloc


def store_object(path, sha, big):
    """Store adressé par contenu (dédup entre snapshots). Retourne le chemin de l'objet."""
    sub = os.path.join(OBJ_DIR, sha[:2])
    os.makedirs(sub, exist_ok=True)
    dst = os.path.join(sub, sha + (".gz" if big else ""))
    if not os.path.exists(dst):
        tmp = dst + ".tmp"
        if big:
            with open(path, "rb") as s, gzip.open(tmp, "wb", compresslevel=6) as d:
                shutil.copyfileobj(s, d)
        else:
            shutil.copyfile(path, tmp)
        os.replace(tmp, dst)
    return dst


def link_into(obj, dst):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    try:
        os.link(obj, dst)
    except OSError:
        shutil.copyfile(obj, dst)


# ---------------------------------------------------------------- versions / meta
def collect_versions():
    v = {}
    v["image_sha (/etc/sand-box-image-sha)"] = read_small("/etc/sand-box-image-sha")
    v["box_scripts_version (/usr/local/share/sand-box-scripts/version)"] = read_small("/usr/local/share/sand-box-scripts/version")
    v["sand_host_version (/opt/sand/sand-host/version)"] = read_small("/opt/sand/sand-host/version")
    try:
        st = json.load(open("/tmp/sand-supervisor/status.json"))
        v["supervisor_status.hostVersion (runtime)"] = str(st.get("hostVersion"))
        v["supervisor_status.lastCommandId (runtime)"] = str(st.get("lastCommandId"))
        v["supervisor_status.pendingUpgradeVersion (runtime)"] = str(st.get("pendingUpgradeVersion"))
    except Exception:
        v["supervisor_status (runtime)"] = "<absent>"
    try:
        pj = json.load(open("/exec-daemon/package.json"))
        v["exec_daemon.name"] = pj.get("name", "")
        v["exec_daemon.gitCommit"] = pj.get("gitCommit", "")
        v["exec_daemon.buildTimestamp"] = pj.get("buildTimestamp", "")
    except Exception:
        v["exec_daemon"] = "<absent>"
    v["canvas_sdk_version"] = read_small("/exec-daemon/agent-sdk/canvas-sdk-version")
    try:
        m = json.load(open("/var/lib/sand-daemon-supervisor/orbitd/bundled/manifest.json"))
        v["orbitd.bundled_revision"] = m.get("revision", "")
    except Exception:
        v["orbitd.bundled_revision"] = "<absent>"
    rel = "/var/lib/sand-daemon-supervisor/orbitd/releases"
    v["orbitd.releases"] = ",".join(sorted(os.listdir(rel))) if os.path.isdir(rel) else "<absent>"
    fuse = read_small("/usr/local/bin/cursor_agent_store_fuse_version")
    mm = re.search(r"([0-9a-f]{40})", fuse)
    v["agent_store_fuse.commit"] = mm.group(1) if mm else fuse
    v["chrome"] = run(["/opt/google/chrome/chrome", "--version"])
    v["node (/exec-daemon/node)"] = run(["/exec-daemon/node", "--version"])
    v["node (system)"] = run(["node", "--version"])
    v["python3"] = run(["python3", "--version"])
    v["debian_version"] = read_small("/etc/debian_version")
    v["kernel (runtime)"] = run(["uname", "-r"])
    v["runtime"] = "docker" if os.path.exists("/.dockerenv") else "anyrun"
    return v


def versions_text(v):
    return "".join(f"{k}\t{val}\n" for k, val in v.items())


def collect_meta():
    meta = {}
    meta["versions.txt"] = versions_text(collect_versions())
    meta["dpkg-packages.txt"] = run(["dpkg-query", "-W", "-f=${Package}\t${Version}\t${Architecture}\n"], 60) + "\n"
    sp = "/usr/local/lib/python3.13/site-packages"
    alt = run(["python3", "-c", "import site;print('\\n'.join(site.getsitepackages()))"]).splitlines()
    dists = []
    for d in [sp] + alt:
        if os.path.isdir(d):
            dists += [f"{d}\t{x}" for x in os.listdir(d) if x.endswith((".dist-info", ".egg-info"))]
    meta["python-dists.txt"] = "\n".join(sorted(set(dists))) + "\n"
    npm = []
    base = "/usr/local/lib/node_modules"
    if os.path.isdir(base):
        for n in sorted(os.listdir(base)):
            pkgs = [os.path.join(base, n)]
            if n.startswith("@"):
                pkgs = [os.path.join(base, n, s) for s in sorted(os.listdir(os.path.join(base, n)))]
            for p in pkgs:
                try:
                    j = json.load(open(os.path.join(p, "package.json")))
                    npm.append(f"{j.get('name')}@{j.get('version')}")
                except Exception:
                    npm.append(os.path.relpath(p, base) + "@?")
    meta["npm-global.txt"] = "\n".join(npm) + "\n"
    names = sorted(k for k in os.environ if k.startswith(("SAND_", "CURSOR_", "BOX_")))
    meta["env-var-names.txt"] = "# noms seulement (valeurs non capturées : peuvent contenir des secrets)\n" + "\n".join(names) + "\n"
    return meta


def quick_fingerprint():
    fp = {}
    v = collect_versions()
    for k, val in v.items():
        if "(runtime)" in k:
            continue
        fp["ver:" + k] = val
    for p in KEY_FILES:
        try:
            fp["sha:" + p] = sha256_file(p)
        except Exception as e:
            fp["sha:" + p] = f"<absent: {type(e).__name__}>"
    for t in KEY_TREES:
        h = hashlib.sha256()
        n = 0
        for p in walk(t, []):
            try:
                if os.path.isfile(p) and not os.path.islink(p):
                    h.update(p.encode() + b"\0" + sha256_file(p).encode() + b"\n")
                    n += 1
            except Exception:
                h.update(p.encode() + b"\0<unreadable>\n")
        fp["tree:" + t] = f"{h.hexdigest()} ({n} fichiers)"
    return fp


def fp_text(fp):
    return "".join(f"{k}\t{v}\n" for k, v in sorted(fp.items()))


def parse_kv(text):
    out = {}
    for line in text.splitlines():
        if "\t" in line:
            k, v = line.split("\t", 1)
            out[k] = v
    return out


# ---------------------------------------------------------------- snapshot
def snapshot_name():
    base = now_local().strftime("%Y-%m-%d_%H-%M")
    name, i = base, 2
    while os.path.exists(os.path.join(SNAP_DIR, name)):
        name = f"{base}-{i}"
        i += 1
    return name


def take_snapshot(label="", quiet=False):
    t0 = time.time()
    roots, excludes = load_perimeter()
    name = snapshot_name()
    final = os.path.join(SNAP_DIR, name)
    work = final + ".partial"
    os.makedirs(work)
    files_dir = os.path.join(work, "files")
    n_files = n_text = n_big = n_err = 0
    total = 0
    with open(os.path.join(work, "manifest.tsv"), "w") as man:
        man.write("path\tsize\tsha256\tmtime\tmode\tkind\tperimeter\n")
        for mode, root in roots:
            for p in walk(root, excludes):
                try:
                    st = os.lstat(p)
                    mtime = datetime.fromtimestamp(st.st_mtime).astimezone().strftime("%Y-%m-%dT%H:%M:%S%z")
                    perm = oct(stat.S_IMODE(st.st_mode))
                    if stat.S_ISLNK(st.st_mode):
                        tgt = os.readlink(p)
                        man.write(f"{p}\t{len(tgt)}\t{hashlib.sha256(tgt.encode()).hexdigest()}\t{mtime}\t{perm}\tsymlink->{tgt}\t{mode}\n")
                        n_files += 1
                        continue
                    if not stat.S_ISREG(st.st_mode):
                        continue
                    sha = sha256_file(p)
                    kind = "bin"
                    total += st.st_size
                    if mode == "full" and is_text(p, st.st_size):
                        if st.st_size <= TEXT_COPY_MAX:
                            kind = "text"
                            link_into(store_object(p, sha, False), files_dir + p)
                            n_text += 1
                        elif st.st_size <= BIG_TEXT_MAX:
                            kind = "text-gz"
                            link_into(store_object(p, sha, True), files_dir + p + ".gz")
                            n_big += 1
                        else:
                            kind = "text-toobig"
                    elif mode == "hash":
                        kind = "hashonly"
                    man.write(f"{p}\t{st.st_size}\t{sha}\t{mtime}\t{perm}\t{kind}\t{mode}\n")
                    n_files += 1
                except Exception as e:
                    man.write(f"{p}\t-\tUNREADABLE:{type(e).__name__}\t-\t-\terror\t{mode}\n")
                    n_err += 1
    meta_dir = os.path.join(work, "meta")
    os.makedirs(meta_dir)
    meta = collect_meta()
    for fn, content in meta.items():
        with open(os.path.join(meta_dir, fn), "w") as f:
            f.write(content)
    fp = quick_fingerprint()
    with open(os.path.join(work, "fingerprint.txt"), "w") as f:
        f.write(fp_text(fp))
    dur = time.time() - t0
    info = (f"snapshot\t{name}\n"
            f"created\t{now_local().strftime('%Y-%m-%d %H:%M:%S %Z')}\n"
            f"label\t{label}\n"
            f"files\t{n_files}\ntext_copies\t{n_text}\nbig_text_gz\t{n_big}\nunreadable\t{n_err}\n"
            f"bytes_hashed\t{total}\nduration_s\t{dur:.1f}\n")
    with open(os.path.join(work, "SNAPSHOT-INFO.txt"), "w") as f:
        f.write(info)
    os.rename(work, final)
    write_version_file(meta["versions.txt"], name)
    log(f"snapshot {name} files={n_files} text={n_text} gz={n_big} err={n_err} {dur:.1f}s label={label!r}")
    if not quiet:
        print(info, end="")
        print(f"path\t{final}")
    return name


def write_version_file(versions_txt, snap):
    with open(VERSION_FILE, "w") as f:
        f.write(f"# Version/build courante de l'app Grok Bot sur la box\n")
        f.write(f"# capturée le {now_local().strftime('%Y-%m-%d %H:%M:%S %Z')} (snapshot {snap})\n")
        f.write("# Box update (Update Computer) = change image_sha ; hot-update host = change sand_host_version\n\n")
        f.write(versions_txt)


def list_snapshots():
    if not os.path.isdir(SNAP_DIR):
        return []
    return sorted(d for d in os.listdir(SNAP_DIR)
                  if os.path.isdir(os.path.join(SNAP_DIR, d)) and not d.endswith(".partial")
                  and os.path.exists(os.path.join(SNAP_DIR, d, "manifest.tsv")))


# ---------------------------------------------------------------- diff
def load_manifest(snap):
    m = {}
    with open(os.path.join(SNAP_DIR, snap, "manifest.tsv")) as f:
        next(f)
        for line in f:
            parts = line.rstrip("\n").split("\t")
            if len(parts) >= 7:
                m[parts[0]] = dict(size=parts[1], sha=parts[2], mtime=parts[3], mode=parts[4], kind=parts[5], per=parts[6])
    return m


def text_of(snap, path, entry):
    base = os.path.join(SNAP_DIR, snap, "files") + path
    if entry["kind"] == "text" and os.path.exists(base):
        return open(base, "rb").read().decode("utf-8", "replace")
    if entry["kind"] == "text-gz" and os.path.exists(base + ".gz"):
        return gzip.open(base + ".gz", "rb").read().decode("utf-8", "replace")
    return None


def normalize_minified(txt):
    return re.sub(r"([;{}])", r"\1\n", txt)


def udiff(a_txt, b_txt, a_label, b_label):
    long = any(len(l) > LONG_LINE for l in (a_txt or "").splitlines()[:5000]) or \
           any(len(l) > LONG_LINE for l in (b_txt or "").splitlines()[:5000])
    if long:
        a_txt, b_txt = normalize_minified(a_txt), normalize_minified(b_txt)
    with tempfile.TemporaryDirectory() as td:
        pa, pb = os.path.join(td, "a"), os.path.join(td, "b")
        open(pa, "w").write(a_txt)
        open(pb, "w").write(b_txt)
        out = subprocess.run(["diff", "-u", "--label", a_label, "--label", b_label, pa, pb],
                             capture_output=True, text=True, errors="replace").stdout
    return out, long


def safe_name(path):
    return path.strip("/").replace("/", "__")


def component(path, roots):
    best = ""
    for _, r in roots:
        if (path == r or path.startswith(r.rstrip("/") + "/")) and len(r) > len(best):
            best = r
    return best or "?"


def make_diff(old=None, new=None, quiet=False):
    snaps = list_snapshots()
    if old is None or new is None:
        if len(snaps) < 2:
            msg = f"Il faut au moins 2 snapshots (trouvés : {len(snaps)}). Lance d'abord snapshot.sh."
            print(msg)
            return None
        old, new = snaps[-2], snaps[-1]
    for s in (old, new):
        if s not in snaps:
            raise SystemExit(f"snapshot inconnu : {s} (dispo : {', '.join(snaps)})")
    roots, _ = load_perimeter()
    A, B = load_manifest(old), load_manifest(new)
    added = sorted(set(B) - set(A))
    removed = sorted(set(A) - set(B))
    modified = sorted(p for p in set(A) & set(B) if A[p]["sha"] != B[p]["sha"])
    perm_only = sorted(p for p in set(A) & set(B) if A[p]["sha"] == B[p]["sha"] and A[p]["mode"] != B[p]["mode"])

    os.makedirs(REPORT_DIR, exist_ok=True)
    adjacent = snaps.index(new) - snaps.index(old) == 1
    rname = new if adjacent else f"{new}_vs_{old}"
    rpath = os.path.join(REPORT_DIR, rname + ".md")
    ddir = os.path.join(REPORT_DIR, rname + "-diffs")
    if os.path.isdir(ddir):
        shutil.rmtree(ddir)
    os.makedirs(ddir)

    va = parse_kv(open(os.path.join(SNAP_DIR, old, "meta", "versions.txt")).read())
    vb = parse_kv(open(os.path.join(SNAP_DIR, new, "meta", "versions.txt")).read())

    L = []
    L.append(f"# Diff box — `{old}` → `{new}`\n")
    L.append(f"Généré le {now_local().strftime('%Y-%m-%d %H:%M %Z')} par box-update-monitor.\n")
    L.append("## Résumé\n")
    L.append(f"| | Nombre |\n|---|---|\n| Ajoutés | {len(added)} |\n| Supprimés | {len(removed)} |\n| Modifiés (hash) | {len(modified)} |\n| Permissions seules | {len(perm_only)} |\n")
    L.append("## Versions / build\n")
    L.append("| Clé | Avant | Après | |\n|---|---|---|---|")
    for k in sorted(set(va) | set(vb), key=lambda k: list(vb).index(k) if k in vb else 999):
        a, b = va.get(k, "—"), vb.get(k, "—")
        L.append(f"| {k} | `{a}` | `{b}` | {'**CHANGÉ**' if a != b else ''} |")
    L.append("")

    # par composant
    comp = {}
    for kind, lst in (("ajoutés", added), ("supprimés", removed), ("modifiés", modified)):
        for p in lst:
            comp.setdefault(component(p, roots), {"ajoutés": 0, "supprimés": 0, "modifiés": 0})[kind] += 1
    if comp:
        L.append("## Par composant\n")
        L.append("| Racine | Ajoutés | Supprimés | Modifiés |\n|---|---|---|---|")
        for r in sorted(comp):
            c = comp[r]
            L.append(f"| `{r}` | {c['ajoutés']} | {c['supprimés']} | {c['modifiés']} |")
        L.append("")

    # meta diffs (paquets, etc.)
    L.append("## Paquets / environnement (meta)\n")
    any_meta = False
    for fn in sorted(os.listdir(os.path.join(SNAP_DIR, new, "meta"))):
        pa = os.path.join(SNAP_DIR, old, "meta", fn)
        pb = os.path.join(SNAP_DIR, new, "meta", fn)
        ta = open(pa).read() if os.path.exists(pa) else ""
        tb = open(pb).read()
        if ta != tb:
            any_meta = True
            out, _ = udiff(ta, tb, f"{old}/meta/{fn}", f"{new}/meta/{fn}")
            open(os.path.join(ddir, "meta__" + fn + ".diff"), "w").write(out)
            lines = out.splitlines()
            L.append(f"### `{fn}`\n\n```diff\n" + "\n".join(lines[:INLINE_DIFF_LINES]) + "\n```")
            if len(lines) > INLINE_DIFF_LINES:
                L.append(f"_… {len(lines) - INLINE_DIFF_LINES} lignes de plus : `{os.path.relpath(ddir, ROOT)}/meta__{fn}.diff`_")
            L.append("")
    if not any_meta:
        L.append("_Aucun changement._\n")

    def fmt(p, e):
        return f"| `{p}` | {e['size']} | `{e['sha'][:12]}` | {e['kind']} |"

    L.append(f"## Fichiers ajoutés ({len(added)})\n")
    if added:
        L.append("| Chemin | Taille | sha256 | Type |\n|---|---|---|---|")
        L += [fmt(p, B[p]) for p in added]
    else:
        L.append("_Aucun._")
    L.append("")
    L.append(f"## Fichiers supprimés ({len(removed)})\n")
    if removed:
        L.append("| Chemin | Taille | sha256 | Type |\n|---|---|---|---|")
        L += [fmt(p, A[p]) for p in removed]
    else:
        L.append("_Aucun._")
    L.append("")
    L.append(f"## Fichiers modifiés ({len(modified)})\n")
    if modified:
        L.append("| Chemin | Taille avant → après | sha256 avant → après | Type |\n|---|---|---|---|")
        for p in modified:
            a, b = A[p], B[p]
            L.append(f"| `{p}` | {a['size']} → {b['size']} | `{a['sha'][:12]}` → `{b['sha'][:12]}` | {b['kind']} |")
    else:
        L.append("_Aucun._")
    L.append("")
    if perm_only:
        L.append(f"## Permissions modifiées ({len(perm_only)})\n")
        L += [f"- `{p}` : {A[p]['mode']} → {B[p]['mode']}" for p in perm_only]
        L.append("")

    # diffs texte
    L.append("## Diff texte ligne à ligne\n")
    n_diff = 0
    for p in modified + added:
        a = A.get(p)
        b = B[p]
        tb = text_of(new, p, b)
        ta = text_of(old, p, a) if a else ""
        if tb is None or (a and ta is None):
            continue
        out, minified = udiff(ta, tb, f"{old}:{p}", f"{new}:{p}")
        if not out:
            continue
        n_diff += 1
        dfile = os.path.join(ddir, safe_name(p) + ".diff")
        open(dfile, "w").write(out)
        lines = out.splitlines()
        plus = sum(1 for l in lines if l.startswith("+") and not l.startswith("+++"))
        minus = sum(1 for l in lines if l.startswith("-") and not l.startswith("---"))
        tag = " (minifié : découpé sur ; { } avant diff)" if minified else ""
        status = "nouveau" if not a else "modifié"
        L.append(f"### `{p}` — {status}, +{plus} / -{minus}{tag}\n")
        L.append("```diff\n" + "\n".join(lines[:INLINE_DIFF_LINES]) + "\n```")
        if len(lines) > INLINE_DIFF_LINES:
            L.append(f"_… {len(lines) - INLINE_DIFF_LINES} lignes de plus : `{os.path.relpath(dfile, ROOT)}`_")
        L.append("")
    if n_diff == 0:
        L.append("_Aucun fichier texte modifié._\n")
    binmods = [p for p in modified if B[p]["kind"] not in ("text", "text-gz")]
    if binmods:
        L.append(f"## Binaires / hash-only modifiés ({len(binmods)}) — pas de diff texte\n")
        L += [f"- `{p}`" for p in binmods[:500]]
        if len(binmods) > 500:
            L.append(f"- … {len(binmods) - 500} de plus")
        L.append("")
    with open(rpath, "w") as f:
        f.write("\n".join(L) + "\n")
    if not os.listdir(ddir):
        os.rmdir(ddir)
    log(f"diff {old} -> {new}: +{len(added)} -{len(removed)} ~{len(modified)} report={rpath}")
    if not quiet:
        print(f"report\t{rpath}")
        print(f"added\t{len(added)}\nremoved\t{len(removed)}\nmodified\t{len(modified)}\ntext_diffs\t{n_diff}")
    return rpath, len(added), len(removed), len(modified)


# ---------------------------------------------------------------- check
def check(deep=False, quiet=False):
    snaps = list_snapshots()
    if not snaps:
        name = take_snapshot("baseline (auto: aucun snapshot)", quiet=True)
        print(f"BASELINE snapshot={name}")
        return 0
    last = snaps[-1]
    old_fp = parse_kv(open(os.path.join(SNAP_DIR, last, "fingerprint.txt")).read())
    cur_fp = quick_fingerprint()
    changed = sorted(k for k in set(old_fp) | set(cur_fp) if old_fp.get(k) != cur_fp.get(k))
    if not changed and deep:
        name = take_snapshot("deep-check", quiet=True)
        A, B = load_manifest(last), load_manifest(name)
        same = A.keys() == B.keys() and all(A[p]["sha"] == B[p]["sha"] and A[p]["mode"] == B[p]["mode"] for p in A)
        ma = open(os.path.join(SNAP_DIR, last, "meta", "dpkg-packages.txt")).read()
        mb = open(os.path.join(SNAP_DIR, name, "meta", "dpkg-packages.txt")).read()
        if same and ma == mb:
            shutil.rmtree(os.path.join(SNAP_DIR, name))
            write_version_file(open(os.path.join(SNAP_DIR, last, "meta", "versions.txt")).read(), last)
            log(f"check --deep: no change vs {last}")
            print(f"NO_CHANGE last={last} (deep)")
            return 0
        changed = ["deep:manifest"]
        new = name
    elif not changed:
        log(f"check: no change vs {last}")
        print(f"NO_CHANGE last={last}")
        return 0
    else:
        new = take_snapshot("auto: changement détecté (" + ", ".join(changed[:5]) + ")", quiet=True)
    res = make_diff(last, new, quiet=True)
    rpath = res[0]
    with open(FLAG, "a") as f:
        f.write(f"{rpath}\t{now_local().strftime('%Y-%m-%dT%H:%M:%S%z')}\t{last}->{new}\t+{res[1]} -{res[2]} ~{res[3]}\n")
    log(f"check: CHANGED {last}->{new} keys={changed}")
    print(f"CHANGED old={last} new={new} report={rpath} added={res[1]} removed={res[2]} modified={res[3]}")
    print("changed_keys=" + ",".join(changed))
    return 10


def main(argv):
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    cmd, args = argv[0], argv[1:]
    if cmd == "snapshot":
        label = ""
        if "--label" in args:
            label = args[args.index("--label") + 1]
        take_snapshot(label)
        return 0
    if cmd == "diff":
        pos = [a for a in args if not a.startswith("-")]
        if len(pos) == 1:
            snaps = list_snapshots()
            make_diff(pos[0], snaps[-1])
        elif len(pos) >= 2:
            make_diff(pos[0], pos[1])
        else:
            make_diff()
        return 0
    if cmd == "check":
        return check(deep="--deep" in args, quiet="--quiet" in args)
    if cmd == "fingerprint":
        print(fp_text(quick_fingerprint()), end="")
        return 0
    if cmd == "versions":
        print(versions_text(collect_versions()), end="")
        return 0
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
