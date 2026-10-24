# live-intercept — capture live du trafic RPC de la box (2 pattes)

Outillage pour capturer **en direct** le trafic RPC de la box sur les deux pattes
et journaliser **requête + réponse (corps complets)** dans les JSONL rédactés
partagés EG (`/workspace/shared-eg/box-vers-vps/sand-host-rpc-traffic/`).

> ⚠️ **Rien ici n'est armé.** `arm.sh` refuse de tourner sans jeton explicite et
> la bascule finale (qui peut couper l'egress) est laissée à l'opérateur, protégée
> par `watchdog.sh`. Ce dossier = scripts testés + runbook.

## 1. Les deux pattes (vérifié sur cette box)

| Patte | Transport | Preuve | Interception |
|-------|-----------|--------|--------------|
| **agent ↔ host** | **HTTP/1.1 EN CLAIR** sur `0.0.0.0:1340` (`SAND_HOST_PORT`) | `curl http://127.0.0.1:1340/` → `HTTP/1.1 404` JSON ; `openssl s_client` → *wrong version number* (pas de TLS) ; h2c refusé | **relai TCP plaintext** — aucune CA requise |
| **host ↔ backend** | **TLS direct par IP** vers `api2.cursor.sh` (52.21.130.202:443), S3, cloudflare — **PAS** via le tunnel `8791` | `ss -tnp` sur pid host-main : connexions `:443` directes, zéro vers `:8791` | **mitmproxy transparent** + host-main doit faire confiance à **notre CA** |

Conséquence : seule la patte **host↔backend** exige la CA et un **redémarrage de
host-main** (donc le risque). La patte **agent↔host** s'intercepte sans CA ni
redémarrage du core, juste en déplaçant host-main de 1340→1341 et en mettant le
relai sur 1340.

## 2. Topologie de supervision (pourquoi le revert est sûr)

```
pod-daemon(7) ─ sand-exit-watch(53) ─ supervise-sand-supervisor(21506)
                                         └ sand-supervisor.mjs(21521)
                                              └ host-main.cjs (LA bouée, uid 1000)
```

`sand-supervisor` **relance automatiquement** host-main dès qu'il meurt, avec
**son propre env d'origine** (port 1340, sans `NODE_EXTRA_CA_CERTS`). De plus il
**adopte** tout `host-main.cjs` vivant qu'il découvre via `gateway.json`
(`maybeAdoptOrphanHost`).

- **revert** exploite ça : tuer host-main ⇒ le superviseur en relance un **propre**
  sur 1340. On n'a donc jamais à « reconstruire » host-main nous-mêmes.
- **arm** exploite l'adoption : on lance NOTRE host-main (nouveau port + CA), il
  écrit `gateway.json`, puis on SIGTERM l'ancien ⇒ le superviseur adopte le nôtre
  au lieu d'en pondre un neuf.

## 3. Méthode CA retenue : `NODE_EXTRA_CA_CERTS` additif

Choisie car **la moins cassante** parmi les trois options :

- ✅ **`NODE_EXTRA_CA_CERTS`** (retenu) : **ajoute** notre CA aux racines bundlées
  de Node sans les remplacer → confiance backend d'origine **préservée**. Lu au
  démarrage de node → injecté en relançant **uniquement host-main** avec son env
  d'origine + cette seule variable. Pas de wrapper, pas de mutation du superviseur.
- ❌ **wrapper `/exec-daemon/node`** : persistant, impacterait **tout** process
  node de la box — interdit.
- ❌ **env du superviseur** : impossible à chaud ; remonterait toute la chaîne
  `pod-daemon` (≈ reboot stack).

## 4. Anti-boucle d'interception

host-main tourne en **uid 1000 (box)**, *le même* que l'agent/les interceptors.
On ne peut donc pas exclure nos propres sondes par uid si elles tournent aussi en
1000 → **les interceptors (mitm + relai) tournent en ROOT (uid 0)**. La règle nft
ne redirige que `meta skuid 1000 tcp dport 443` ; le trafic uid-0 du mitm vers le
vrai backend n'est **pas** repris → pas de boucle.

## 5. Netfilter : backend ACTIF = `nft`

La box est dans le **netns racine réel** (`net:[4026531840]`, tailscale + docker
présents). Le backend netfilter vivant est **nftables** (`table ip nat` avec les
chaînes tailscale ; `iptables` était absent, installé en `iptables-nft`). On
utilise donc **`nft`** et une **table dédiée `ip sandrpc`** — créée/hookée
seulement par `arm`, supprimée en bloc par `revert` (`nft delete table ip sandrpc`).
La patte agent↔host **n'utilise pas** nft (relai direct sur 1340).

## 6. Scripts

| Script | Rôle | Exécute-t-il qqch d'irréversible ? |
|--------|------|-----------------------------------|
| `watchdog.sh` | dead-man's switch **détaché (setsid)**, survit à la mort de l'agent | non (sauf s'il déclenche `revert`) |
| `revert.sh` | restaure TOUT (delete table nft, kill interceptors, bounce host-main→respawn propre 1340) ; idempotent ; `--dry-run` | oui (c'est son rôle) — mais sûr |
| `arm.sh` | la bascule ; **refuse de tourner** sans `LI_ARM_CONFIRM=yes-i-understand-brick-risk` (sinon dry-run) | oui, seulement avec le jeton |
| `relay_agent_host.py` | relai HTTP/1.1 keep-alive loggeur, patte agent↔host (`leg=agent-host`) | non |
| `start-mitm-host-backend.sh` | mitmproxy **transparent** patte host↔backend (`leg=host-backend`), en root | non |
| `lib.sh` | config + helpers communs | non |
| `test/run_test.sh` | PREUVE capture+rédaction+revert sur cibles **scratch** (ne touche pas host-main) | non |

Rédaction/décodage/rotation : **réutilisés** depuis `../rpc-traffic-monitor/proxy.py`
+ `../rpc-traffic-monitor/mitmproxy/addon_log.py` (headers sensibles, JWT, Bearer,
e-mails, clés GH/PEM, valeurs d'env secrètes ; dé-framing Connect/gRPC ; gunzip).

## 7. RUNBOOK

### Armer (opérateur, après revue)
```bash
cd .../sand-host-rpc/live-intercept
bash arm.sh                      # DRY : imprime le plan, ne touche à rien
LI_ARM_CONFIRM=yes-i-understand-brick-risk sudo -E bash arm.sh   # bascule réelle
```
`arm` fait, dans l'ordre : pose `ARM` → lance `watchdog.sh` (grace 90s) → mitm
transparent (root, :8900) → spawn host-main sur 1341 (+CA) → SIGTERM l'ancien →
relai sur 1340 → règle nft redirect 443.

### Ce qui peut casser
Si host-main **épingle le certificat** (cert-pinning) du backend, il rejettera
notre CA ⇒ plus d'egress host↔backend ⇒ **l'agent meurt** (host-main est la bouée).

### Comment le watchdog sauve
Détaché (`setsid`), il probe `api2.cursor.sh:443` toutes les 5s. Si injoignable
3 probes d'affilée **ou** si `grace` (90s) s'écoule **sans `DISARM`**, il lance
`revert.sh` tout seul → la box se répare même si l'agent est mort.

### Confirmer / désarmer / annuler
```bash
touch /workspace/.sand-live-intercept/DISARM   # « arm OK, annule le dead-man » (garde la capture)
bash revert.sh                                  # annuler/teardown propre (pose aussi DISARM)
bash revert.sh --dry-run                        # voir ce que revert ferait
```

### Exit codes
- `watchdog.sh` : 0 sortie propre (DISARM/ARM retiré) · 10 revert (backend KO) · 11 revert (grace)
- `revert.sh` : 0 host-main sain sur 1340 + backend joignable · 1 sinon
- `arm.sh` : 0 dry-run · 2 pas de host-main · 3 host-main KO sur 1341 (revert auto)

## 8. Estimation du risque de pinning

**Modéré-à-élevé, non vérifiable sans armer.** host-main est un bundle Node
(`application/json` Connect). Les clients Node/undici standards suivent le magasin
de CA (donc `NODE_EXTRA_CA_CERTS` fonctionnerait) ; le pinning explicite est
possible côté backend sécurisé mais rare en pur Node sans code dédié. Le test
scratch prouve que l'addon capture un TLS ré-signé par notre CA quand le client
l'accepte (`--cacert`). **Le seul vrai test est l'arm**, d'où : watchdog armé
AVANT la coupure + grace courte + revert auto. Mon conseil reste de privilégier la
capture host↔backend **côté EG** (env maîtrisé) et de n'armer la box que pour la
patte **agent↔host** (sans CA, sans redémarrage du core → risque quasi nul).

## 9. Limites connues
- `proxy.py` : `PKEY_RE` rédige désormais le bloc PEM COMPLET (en-tête + corps base64 + pied),
  pas le corps base64 d'une vraie clé. À durcir dans `rpc-traffic-monitor/`
  (hors scope de ce commit, strictement `live-intercept/`).
- La patte host↔backend ne capte que les hôtes redirigés (uid 1000 : host-main) ;
  S3/CDN par IP directe sont redirigés aussi (dport 443) mais étiquetés `autre:`.
