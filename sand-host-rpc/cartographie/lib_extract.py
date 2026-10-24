#!/usr/bin/env python3
"""lib_extract — extracteur EXHAUSTIF des services/méthodes RPC du host.

Scanne TOUS les .cjs/.js/.mjs de /opt/sand (pas seulement host-main.cjs et
sand-eval-runner.cjs), repère chaque définition de service Connect/gRPC
(`typeName: "<pkg>.<Service>"` suivi d'un bloc `methods:{…}`), quel que soit le
package (aiserver.v1, agent.v1, origin.v1, orbit.v1, …), et énumère toutes les
méthodes (`name/I/O/kind`) via un balayeur d'accolades conscient des chaînes et
commentaires (évite les sur-dépassements des parseurs naïfs sur code minifié).
"""
import os, re

ROOTS = ["/opt/sand"]
EXTS = (".cjs", ".js", ".mjs")
TYPENAME = re.compile(r'typeName:\s*"([A-Za-z_][A-Za-z0-9_.]*\.[A-Za-z0-9_]+)"')
METHOD = re.compile(
    r'name:\s*"([A-Za-z0-9_]+)"\s*,\s*I:\s*([A-Za-z0-9_$.]+)\s*,\s*O:\s*([A-Za-z0-9_$.]+)'
    r'\s*,\s*kind:\s*MethodKind\.([A-Za-z]+)'
)


def _methods_block(data, start):
    mm = data.find("methods:", start)
    if mm < 0 or mm - start > 80:
        return None
    b = data.find("{", mm)
    if b < 0:
        return None
    i = b; depth = 0; n = len(data)
    while i < n:
        c = data[i]
        if c in '"\'`':
            q = c; i += 1
            while i < n:
                if data[i] == "\\": i += 2; continue
                if data[i] == q: break
                i += 1
            i += 1; continue
        if c == "/" and i + 1 < n and data[i + 1] == "/":
            j = data.find("\n", i); i = n if j < 0 else j + 1; continue
        if c == "/" and i + 1 < n and data[i + 1] == "*":
            j = data.find("*/", i); i = n if j < 0 else j + 2; continue
        if c == "{": depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return data[b:i + 1]
        i += 1
    return None


def extract(roots=None):
    """Retourne (services, files) où services = {fqn: {method: {'I','O','kind'}}}
    et files = liste des fichiers contenant des définitions de service."""
    roots = roots or ROOTS
    services = {}
    files = set()
    for root in roots:
        for dp, _, fs in os.walk(root):
            for fn in fs:
                if not fn.endswith(EXTS):
                    continue
                p = os.path.join(dp, fn)
                try:
                    data = open(p, "r", errors="replace").read()
                except Exception:
                    continue
                if "typeName:" not in data or "MethodKind" not in data:
                    continue
                hit = False
                for m in TYPENAME.finditer(data):
                    fqn = m.group(1)
                    block = _methods_block(data, m.end())
                    if not block:
                        continue
                    for me in METHOD.finditer(block):
                        services.setdefault(fqn, {}).setdefault(
                            me.group(1),
                            {"I": me.group(2), "O": me.group(3), "kind": me.group(4)},
                        )
                        hit = True
                if hit:
                    files.add(p)
    return services, sorted(files)


def triples(roots=None):
    """Liste plate [ [fqn, method, kind], … ] triée."""
    services, _ = extract(roots)
    out = []
    for fqn in services:
        for m, meta in services[fqn].items():
            out.append([fqn, m, meta["kind"]])
    out.sort()
    return out


if __name__ == "__main__":
    svc, files = extract()
    tot = sum(len(v) for v in svc.values())
    print(f"services={len(svc)} methodes={tot} fichiers={len(files)}")
    for f in files:
        print("  ", f)
    for s in sorted(svc):
        print(f"  {s}: {len(svc[s])}")
