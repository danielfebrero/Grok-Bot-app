# live-intercept — capture live du trafic RPC de la box (2 pattes)

Outillage pour capturer **en direct** le trafic RPC de la box sur les deux pattes
et journaliser **requête + réponse (corps complets)** dans les JSONL rédactés
partagés EG (`/workspace/shared-eg/box-vers-vps/sand-host-rpc-traffic/`).

> ⚠️ **Rien ici n'est armé.** `arm.sh` et `hb-setup-keylog.sh` refusent de tourner
> sans jeton explicite. La bascule (un restart de host-main) est laissée à
> l'opérateur, protégée par `watchdog.sh`. Ce dossier = scripts testés + runbook.

## 1. Les deux pattes (vérifié sur cette box)

| Patte | Transport | Preuve | Interception |
|-------|-----------|--------|--------------|
| **agent ↔ host** | **HTTP/1.1 EN CLAIR** sur `0.0.0.0:1340` (`SAND_HOST_PORT`) | `curl http://127.0.0.1:1340/` → `HTTP/1.1 404` JSON ; `openssl s_client` → *wrong version number* (pas de TLS) ; h2c refusé | **relai TCP plaintext** — aucune CA, aucun keylog |
| **host ↔ backend** | **TLS direct** vers `api2`/`api3`/…`cursor.sh:443` — **PAS** via le tunnel egress `198.18.0.1` | connexions `:443` directes du pid host-main, zéro vers le tunnel | **capture passive** : `--tls-keylog` + tcpdump/tshark. Pas de mitm, pas de CA, pas de redirect |

La patte **agent↔host** s'intercepte sans redémarrer le core (§10 : relai + nft
sur le seul client agent). La patte **host↔backend** demande **un** restart de
host-main, une seule fois, pour ajouter `--tls-keylog` — ensuite on/off =
démarrer ou arrêter la capture, sans nouveau restart (§11).

## 2. Topologie de supervision (pourquoi le revert est sûr)

```
pod-daemon(7) ─ sand-exit-watch(53) ─ supervise-sand-supervisor(21506)
                                         └ sand-supervisor.mjs(21521)
                                              └ host-main.cjs (LA bouée, uid 1000)
```

`sand-supervisor` **relance automatiquement** host-main dès qu'il meurt, avec
**son propre env d'origine** (port 1340, **sans** `NODE_EXTRA_CA_CERTS`, **sans**
`SSLKEYLOGFILE`, **sans** `--tls-keylog`). De plus il **adopte** tout
`host-main.cjs` vivant qu'il découvre via `gateway.json` (`maybeAdoptOrphanHost`).

- **hb-revert / revert** exploitent ça : tuer host-main ⇒ le superviseur en relance
  un **propre** sur 1340. On n'a donc jamais à « reconstruire » host-main nous-mêmes.
- **hb-setup-keylog** exploite l'adoption : on lance NOTRE host-main (même port
  1340, argv `--tls-keylog` seulement), il écrit `gateway.json`, puis on SIGTERM
  l'ancien ⇒ le superviseur adopte le nôtre au lieu d'en pondre un neuf.

## 3. Déchiffrement host↔backend : `--tls-keylog`, pas une CA

Vérifié sur `/exec-daemon/node` **v22.23.4** (OpenSSL 3.5, undici intégré) :

- L'env **`SSLKEYLOGFILE` seul ne crée aucun fichier.** Le binaire ne contient
  pas cette chaîne ; le callback keylog n'est armé que par l'option.
- **`--tls-keylog=FICHIER`** écrit le format NSS que Wireshark appelle
  SSLKEYLOGFILE (`CLIENT_HANDSHAKE_TRAFFIC_SECRET`, `CLIENT_TRAFFIC_SECRET_0`,
  `SERVER_*`, `EXPORTER_SECRET`). Node avertit que ce fichier contient des
  secrets de session. Mode **0600** (append).
- On pose **aussi** `SSLKEYLOGFILE=` vers le même chemin : no-op sur ce build,
  utile si un Node futur honore l'env. On **retire** `NODE_EXTRA_CA_CERTS` et
  tout `SSLKEYLOGFILE` précédent.
- Pas de wrapper sur `/exec-daemon/node` (impacterait tout node de la box).
  Pas de mutation de l'env du superviseur.

Le TLS vers le vrai backend n'est pas terminé par nous : pas de certificat
maison, pas de pinning à satisfaire. tshark déchiffre **après coup** avec
`tls.keylog_file`.

## 4. Pas de boucle : la capture est passive

Rien n'est redirigé. tcpdump copie les paquets ; le trafic de host-main continue
d'aller au vrai `:443`. Le filtre BPF par défaut est
`tcp port 443 and not host 198.18.0.1` (le tunnel egress, donc les Chrome
agents, reste dehors). Le décodeur ne **journalise** que les échanges dont le
SNI (sinon `:authority` / `Host`) matche `HB_ALLOW_HOSTS` et dont
`addon_log._direction` vaut `host->backend`. Le reste peut rester chiffré dans
un pcap éphémère mode 0600, détruit par `hb-off`, et n'entre pas dans le JSONL.

Les clés n'existent que pour le process Node lancé avec `--tls-keylog`. Un autre
client TLS de la box n'est pas déchiffrable avec ce fichier.

## 5. Netfilter

La box est dans le **netns racine réel** (tailscale + docker présents). Le
backend netfilter vivant est **nftables**.

- **host↔backend n'ajoute aucune règle.** L'ancienne table mitm `ip sandrpc_hb`
  (redirect `:443` → mitm) est **supprimée** si elle traîne encore
  (`hb-setup-keylog`, `hb-off`, `hb-revert`). Elle n'est plus jamais créée.
- **agent↔host (§10)** utilise la table `ip sandrpc` (prerouting, un seul
  client). `hb-*` ne la touche pas. `revert.sh` (chemin combiné historique) la
  supprime encore : ne pas le lancer pour « éteindre » seulement le keylog —
  utiliser `hb-revert.sh`.

## 6. Scripts

| Script | Rôle | Exécute-t-il qqch d'irréversible ? |
|--------|------|-----------------------------------|
| `watchdog.sh` | dead-man's switch **détaché (setsid)**, survit à la mort de l'agent | non (sauf s'il déclenche le revert qu'on lui donne) |
| `hb-setup-keylog.sh` | **un** restart host-main : argv `--tls-keylog` + env `SSLKEYLOGFILE`, capture démarrée **avant** le restart. Jeton `LI_ARM_CONFIRM` | oui, seulement avec le jeton |
| `hb-on.sh` / `hb-off.sh` | démarre / arrête tcpdump + décodeur. **Aucun** restart, **aucune** nft | non (`hb-off` détruit le pcap, garde le keylog) |
| `hb-revert.sh` | stop capture, shred pcap **et** keylog, SIGTERM host-main → respawn sans flag | oui (c'est son rôle) — host propre au retour |
| `start-capture-host-backend.sh` | tcpdump (`--immediate-mode`) + `decode_keylog.py`. Root. Appelé par setup/on | non |
| `decode_keylog.py` | tshark + keylog NSS → JSONL rédigé `leg=host-backend`. `--plaintext` : WebSocket CDP loopback | non |
| `cdp-on.sh` / `cdp-off.sh` / `start-capture-cdp.sh` | tcpdump `lo` sur les ports CDP + décodeur WS. Pas de restart, pas de nft | `cdp-on` oui, avec `CDP_ON_CONFIRM=yes` |
| `arm-agent-host.sh` / `revert-agent-host.sh` / `relay_agent_host.py` | patte agent↔host, **inchangés** (§10) | `arm-agent-host` oui, avec le jeton |
| `arm.sh` / `revert.sh` | chemin combiné **historique**. `arm.sh` **n'arme plus** le host↔backend (pas de CA, pas de `--tls-keylog`, pas de nft :443). Ne pas l'enchaîner après `hb-setup-keylog` : il relance host-main **sans** le flag | oui, avec le jeton |
| `hb-setup-ca.sh`, `start-mitm-host-backend.sh` | **retirés** (exit 1, message vers les scripts keylog) | non |
| `lib.sh` | config + helpers communs | non |
| `test/test-host-backend.sh` | preuve déchiffrement + rédaction sur cibles **scratch** (pas de host-main, pas de nft) | non |
| `test/test_decode_bodies.py` | réassemblage HTTP/2 (corps proto non null) + décodage catalogue ; pcap scratch si tshark/tcpdump | non |

Rédaction/décodage/rotation : **réutilisés** depuis `../rpc-traffic-monitor/proxy.py`
+ `../rpc-traffic-monitor/mitmproxy/addon_log.py` (headers sensibles, JWT, Bearer,
e-mails, clés GH/PEM, valeurs d'env secrètes ; dé-framing Connect/gRPC ; gunzip ;
protobuf → JSON via `../cartographie/data/rpc_catalog.json`).

## 7. RUNBOOK

### Armer (opérateur, après revue)
```bash
cd .../sand-host-rpc/live-intercept
bash arm.sh                      # DRY : imprime le plan, ne touche à rien
LI_ARM_CONFIRM=yes-i-understand-brick-risk sudo -E bash arm.sh   # bascule réelle
```
`arm.sh` **ne fait plus** la patte host↔backend (ni mitm, ni CA, ni nft :443, ni
`--tls-keylog`). Pour cette patte, suivre le §11. Pour l'agent↔host sans
restart, suivre le §10 (`arm-agent-host.sh`), pas `arm.sh`.

### Ce qui peut casser
`arm.sh` déplace encore host-main de 1340 vers 1341. Ce n'est pas le chemin
host↔backend. Le risque de la patte backend est décrit au §8 et au §11 : un
restart / une course d'adoption, pas un pinning de CA.

### Comment le watchdog sauve
Détaché (`setsid`), il probe `api2.cursor.sh:443` toutes les 5s. Si injoignable
3 probes d'affilée **ou** si `grace` (90s) s'écoule sans le flag de désarmement,
il lance le revert qu'on lui a donné. `arm.sh` utilise `DISARM` + `revert.sh`.
`hb-setup-keylog.sh` utilise `HB_DISARM` + `hb-revert.sh` (flags dédiés, la
patte agent↔host n'est pas emportée).

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

## 8. Risque résiduel (plus de pinning)

Le TLS n'est plus intercepté : host-main parle au certificat **réel** du backend.
Le pinning ne s'applique plus.

Ce qui reste :

- **Un restart** de host-main pour injecter `--tls-keylog` (et seulement ça).
  Course d'adoption possible, la même que l'ancien `arm` : le watchdog est armé
  **avant**, et `hb-revert.sh` SIGTERM host-main pour que le superviseur en
  relance un sans le flag. Tant que le backend répond, le dead-man ne tire pas.
- **Le fichier de clés** contient des secrets de session TLS. Il vit dans
  `STATE_DIR/keylog/sslkeys.log` (défaut `/workspace/.sand-live-intercept/keylog/`),
  mode **0600**, propriétaire uid host-main (1000), **hors** dépôt et **hors**
  `TRAFFIC_LOG_DIR`. Jamais recopié dans les JSONL. `hb-revert` le shred.
  L'environnement de host-main est lu dans `/proc/PID/environ` **en mémoire** et
  n'est pas écrit sur disque.
- **hb-off ne retire pas les clés** : host-main continue d'append au fichier
  0600. C'est voulu (un `hb-on` suivant n'a pas besoin d'un restart). Pour
  couper l'écriture : `hb-revert.sh`.

## 9. Limites connues
- `proxy.py` : `PKEY_RE` rédige le **BLOC COMPLET**
  `-----BEGIN … PRIVATE KEY----- … -----END … PRIVATE KEY-----` (`re.DOTALL`),
  en-tête et corps base64.
- **Wireshark a besoin du ClientHello dans le pcap** pour associer les secrets
  TLS 1.3. Les pools HTTP/2 ouverts **avant** la capture restent opaques
  jusqu'à une reconnexion. `hb-setup-keylog` démarre donc tcpdump **avant** le
  restart, pour prendre les handshakes du nouveau process.
- **`SSLKEYLOGFILE` seul est ignoré** sur le Node de cette box. Le flag
  `--tls-keylog` est obligatoire. Les deux sont posés vers le même chemin.
- tcpdump est lancé avec **`--immediate-mode`**. Sans ça, TPACKET_V3 ne livre
  les paquets qu'au bout d'environ une seconde : un `hb-off` dans la foulée
  voyait « N received, 0 captured » et un pcap vide.
- Dépendances : `tcpdump` et `tshark` (Wireshark). Sans eux, `hb-on` / la
  capture sortent en erreur ; host-main n'est pas modifié par `hb-on`.
- HTTP/1.1 et HTTP/2 sont reassemblés. Seuls les SNI de `HB_ALLOW_HOSTS` sont
  journalisés (`leg=host-backend`). Le pcap est éphémère (plafond
  `HB_PCAP_MAX_BYTES`, défaut 64 Mio) et **shreddé par hb-off**.
- Le décodeur ne déchiffre que ce que **ce** Node a mis dans le keylog. Le BPF
  peut voir d'autre TCP/443 ; il n'est pas écrit dans le JSONL.

## 10. Variante « sans restart host-main » (patte agent↔host) — RECOMMANDÉE

Fichiers : `arm-agent-host.sh`, `revert-agent-host.sh`, relai `relay_agent_host.py`
(args `--listen HOST:PORT --upstream HOST:PORT`), `watchdog.sh --revert …`.

La patte agent↔host étant du **HTTP/1.1 en clair**, on n'a besoin ni de CA ni de
redémarrer le core. On intercale juste un relai loggeur devant le **host-main réel
(intact)** et on redirige, en `nft` PREROUTING, le **seul** trafic du client agent.

```
agent (172.30.0.1) ──[nft prerouting: saddr 172.30.0.1, daddr 172.30.0.2:1340 → :1342]──▶
   relai loggeur (0.0.0.0:1342) ──▶ host-main RÉEL (172.30.0.2:1340, NON redémarré)
```

**Pourquoi aucune boucle (vérifié sur cette box) :**
- `172.30.0.2` est une adresse **locale** (`enp0s4`, route via `lo`, `src 172.30.0.2`).
  La connexion sortante du relai vers `172.30.0.2:1340` part donc avec
  `saddr=172.30.0.2`, **jamais** `172.30.0.1` → la règle (qui ne matche que
  `saddr 172.30.0.1`) ne la redirige pas.
- Un paquet localement généré vers une adresse locale passe par **OUTPUT**, pas par
  **PREROUTING** : la connexion du relai ne traverse même pas la chaîne.
- `skuid 0` est **inopérant en prerouting** (paquet entrant d'un autre netns = pas de
  socket propriétaire) : c'est le match `saddr` qui protège. Défense en profondeur :
  scoping `iifname "enp0s4"` ajouté à la règle.

**Bind du relai :** la cible `redirect` réécrit la daddr vers l'adresse de l'iface
d'entrée (`172.30.0.2`), **pas** `127.0.0.1`. Le relai écoute donc `0.0.0.0:1342`
(couvre 127.0.0.1 **et** 172.30.0.2) pour que le paquet redirigé atterrisse bien.

**Armer réellement :**
```
LI_ARM_CONFIRM=yes-i-understand-brick-risk sudo -E bash arm-agent-host.sh
```
(sans jeton = DRY-RUN, n'écrit rien). Confirmer : `touch /workspace/.sand-live-intercept/DISARM`.
Annuler : `bash revert-agent-host.sh` (idempotent, `--dry-run` dispo). **host-main
n'est jamais touché par cette variante** ; risque ≈ nul (pas de CA, pas de 443, pas
de restart). Watchdog armé AVANT la règle, dead-man `grace` 90 s → revert auto si
pas de DISARM.

**Test (cibles scratch, n'arme rien) :** `bash test/test-agent-host.sh` prouve
(a) req+resp capturés, (b) `leg=agent-host`, (c) grep anti-secret VIDE (clé privée
en bloc complet), (d) `revert-agent-host.sh --dry-run` OK.

## 11. Patte host↔backend : keylog NSS + capture passive

La patte host↔backend est du **TLS direct** vers `api2.cursor.sh` (et le set
backend), **hors** tunnel egress. On ne termine plus le TLS. Node écrit les
secrets de session (`--tls-keylog`), tcpdump copie les paquets, tshark déchiffre,
et le JSONL rédigé part dans le dossier partagé EG avec `leg=host-backend`.

Prérequis : `tcpdump` et `tshark` (`apt-get install tcpdump tshark`).

### 11.1 Qui est journalisé

Le BPF par défaut ne prend que `tcp port 443` **sauf** `198.18.0.1`. Le
décodeur ne garde ensuite que le SNI (sinon `:authority` / `Host`) qui matche
`HB_ALLOW_HOSTS`, et seulement si la direction est `host->backend`.

| Hôte (SNI, sinon :authority / Host) | Journal |
|---|---|
| `api` / `api2` / `api3` / `api4` / `repo42.cursor.sh`, `cursor.sh`, `cursor.com`, `www.cursor.com` | **oui**, `leg=host-backend`, req + resp rédigés |
| `198.18.0.1` (tunnel egress → Chrome agents) | **hors BPF** |
| S3, Cloudflare/CDN, tout autre nom | **non** (le pcap peut les contenir chiffrés ; il est détruit par `hb-off`) |

`HB_ALLOW_HOSTS` (ancré, anti-spoof — `settings.cursor.sh.evil.com` rejeté) :
```
(^|\.)(api|api2|api3|api4|repo42)\.cursor\.sh(:443)?$|(^|\.)(www\.)?cursor\.com(:443)?$|(^|\.)cursor\.sh(:443)?$
```

Iface par défaut : `HB_IFACE=any`. Le test scratch force `lo`.

### 11.2 Le seul moment à risque = `hb-setup-keylog.sh` (un restart)

C'est **la seule** commande host↔backend qui touche host-main. Sans jeton elle
imprime le plan et sort.

Elle lit `/proc/PID/environ` en mémoire (jamais écrit, jamais loggé), retire
`NODE_EXTRA_CA_CERTS` et un ancien `SSLKEYLOGFILE`, ajoute
`SSLKEYLOGFILE=<keylog>`, et relance **le seul** host-main sur **le même port
1340** avec `--tls-keylog=<même fichier>`. Le relai agent↔host, s'il est en
place, reste sur 1340. Le superviseur adopte ce process.

Ordre : dead-man `watchdog.sh` (flags `HB_ARM` / `HB_DISARM`) → keylog vide
0600 → suppression d'un reliquat `ip sandrpc_hb` / mitmdump :8900 → **capture
démarrée** → restart → santé sur :1340 + cmdline contient `--tls-keylog` →
marqueur `hb_keylog_ready`. Si la capture ne démarre pas, host-main n'est pas
touché. Si le backend ne répond pas après le restart, le watchdog lance
`hb-revert.sh`.

```
cd .../sand-host-rpc/live-intercept
bash hb-setup-keylog.sh          # DRY-RUN
LI_ARM_CONFIRM=yes-i-understand-brick-risk sudo -E bash hb-setup-keylog.sh
touch /workspace/.sand-live-intercept/HB_DISARM    # « setup OK, annule le dead-man »
```

Ne pas enchaîner `arm.sh` ensuite : il relance host-main **sans** `--tls-keylog`.

### 11.3 ON / OFF = capture seule (aucun restart, aucune nft)

```
HB_ON_CONFIRM=yes sudo -E bash hb-on.sh     # tcpdump + décodeur, si pas déjà up
bash hb-off.sh                              # SIGTERM, passe finale, shred du pcap
```

`hb-on` exige le marqueur `hb_keylog_ready` **et** `--tls-keylog` dans la
cmdline du host-main vivant (`HB_FORCE=1` saute le contrôle, déconseillé).
Sans `HB_ON_CONFIRM=yes` (ou le même `LI_ARM_CONFIRM`), c'est un dry-run.
`LI_ARM_CONFIRM` déjà posé compte comme confirmation.

`hb-off` garde le keylog et le marqueur. Les handshakes pris pendant que la
capture tourne sont déchiffrables ; ceux d'avant `hb-on` ne le sont pas
(ClientHello absent du pcap).

### 11.4 Revert complet
```
bash hb-revert.sh            # stop capture, shred pcap + keylog, SIGTERM host-main
bash hb-revert.sh --dry-run  # plan, n'exige pas root, n'exécute rien
```
Le superviseur relance host-main **sans** `--tls-keylog` et **sans**
`SSLKEYLOGFILE`. Ne touche **jamais** `ip sandrpc` ni le relai agent↔host.
Supprime un reliquat mitm (`ip sandrpc_hb`, mitmdump :8900) s'il existe.

### 11.5 Sortir de l'ancien mitm + CA

`hb-setup-ca.sh` et `start-mitm-host-backend.sh` sortent **1** et n'installent
plus rien. Si un ancien arm mitm tourne encore :

1. `bash hb-off.sh` puis `bash hb-revert.sh` — retire la table `ip sandrpc_hb`,
   tue mitmdump, et le respawn de host-main **enlève** `NODE_EXTRA_CA_CERTS`.
2. `apt-get install tcpdump tshark`
3. Le `hb-setup-keylog.sh` du §11.2 (un restart, capture déjà armée, CA absente).
4. `touch …/HB_DISARM`, puis `hb-on` / `hb-off` sans autre restart.

### 11.6 Test (cibles scratch, n'arme rien de réel)

`bash test/test-host-backend.sh` ne lance pas host-main et ne pose pas de nft.
Trois faux serveurs sur `127.0.0.1`, SNI forcé par le client Node
(`--tls-keylog`), capture via le script de prod :

- `scratchtest.cursor.sh` en HTTP/1.1 → journalisé, rédigé ;
- `api2.cursor.sh` en HTTP/2 → journalisé, rédigé ;
- `pass.scratch.test` → déchiffrable mais **absent** du JSONL ;
- grep anti-secret vide (Bearer, `ghp_`, PEM complet, e-mail, JWT, cookie,
  labels de keylog) ;
- `HB_ALLOW_HOSTS` matche le set backend et rejette tunnel / S3 / CDN / spoof ;
- dry-run de setup / on / off / revert ; les stubs mitm sortent 1 ;
- `arm-agent-host.sh` reste un dry-run (patte non armée) ;
- `hb-off` détruit le pcap et **garde** le keylog en 0600.

`bash test/run_test.sh` enchaîne le relai agent↔host scratch, ce test, et
`revert.sh --dry-run`.

## 12. Corps JSONL : protobuf lisible, HTTP/2 complet

`application/proto` (Connect unaire) n'est plus du texte UTF-8. `req_body` /
`resp_body` sont des objets JSON : noms de champs du catalogue, tags inconnus
en `_unknown_<n>`. Exemple : `GetManagedSkills` → `skills[].id`, `skills[].content`.

tshark réassemble les DATA HTTP/2 et **recolle le corps entier** dans la
dernière trame (`http2.body.reassembled.data`, et souvent `http2.data.data`
plus long que `http2.length`) **en plus** des fragments. Le décodeur garde ce
corps réassemblé. En `--follow`, une réponse n'est écrite qu'une fois les deux
côtés terminés (END_STREAM, ou `Content-Length` atteint) — sinon la ligne
partait avec `resp_body: null` alors que `content-length` était déjà grand.

Limites : un stream HTTP/2 unaire qui ne finit pas pendant la capture n'est
pas écrit. Une enveloppe `application/connect+proto` / gRPC complète est
dé-tramée, chaque payload décodé. Le WebSocket est traité à part (§13).

## 13. WebSocket

Trois canaux, pas le même transport. Le JSONL reste le même fichier
(`protocol: websocket`, `ws_opcode`, `ws_path`, payload dans `req_body` ou
`resp_body`). Le dashboard lit ces lignes comme les échanges unaires.

| Canal | Où ça passe | Comment les trames sont obtenues |
|---|---|---|
| **CDP** (le canal utile) | `ws://127.0.0.1:<port>/devtools/browser/<id>` et `/devtools/page/<targetId>`, en clair. `host-main` fait `new WebSocket(webSocketDebuggerUrl)`. Ports vus : **9229, 9257, 9302**, aussi via `GET /json/version` et `/json/list` | tcpdump sur `lo` + tshark `-d tcp.port==PORT,http`. La couche `websocket` donne le texte déjà démasqué (`websocket.payload.text`) ; un opcode 2 (binaire) est le `data.data` voisin. `cdp-on.sh` / `cdp-off.sh`. leg=`host-cdp` |
| **agent↔host** | le relai `:1342→:1340` ne voyait pas l'`Upgrade` | si `Upgrade: websocket` et réponse **101**, le relai fait suivre les octets **tels quels** et parse les trames en Python (RFC 6455). leg=`agent-host` |
| **host↔backend** (best-effort) | TLS tcp/443, filtre inchangé (`not host 198.18.0.1`) | même décodeur, après le keylog. HTTP/1.1 `Upgrade` ou HTTP/2 extended CONNECT (`:protocol=websocket`) si tshark émet la couche `websocket` ou des DATA. Une session voix (`websocket_url` à côté de `client_secret`) peut passer par le **tunnel egress exclu** : elle n'apparaît alors pas |

Sens d'une trame : bit MASK (client→serveur est masqué). Pour le CDP, le client est host-main (`host->cdp` dans `req_body`) et Chrome répond non masqué (`cdp->host` dans `resp_body`). Le JSON CDP (`method` / `params` / `result`) est parsé ; un `method` CDP (ex. `Page.navigate`) remplace le path dans la colonne méthode. Binaire : proto nommé si le path est dans le catalogue, sinon hex. Contrôle (close / ping / pong) : résumé. Rédaction identique aux autres corps.

```
cd .../sand-host-rpc/live-intercept
bash cdp-on.sh                         # dry-run, affiche les ports découverts
CDP_PORTS='9229 9302' CDP_ON_CONFIRM=yes sudo -E bash cdp-on.sh
bash cdp-off.sh                        # shred du pcap, ne touche pas le 443
```

Limites :

- **permessage-deflate** (bit RSV) n'est pas décompressé. La trame est notée
  `_ws: compressed` + hex. CDP ne négocie pas cette extension.
- Les **fragments** sont recollés (le relai bufferise ; tshark met le message
  entier sur la trame de continuation, le fragment FIN=0 est vide et ignoré).
  Un fragment coupé par la fin du pcap n'est pas écrit.
- La capture doit contenir le **handshake HTTP** (comme le ClientHello pour
  TLS) : sans lui tshark ne dissèque pas les trames suivantes.
- **noVNC / websockify (6080, 6081)** et **198.18.0.1** sont hors périmètre,
  même si on les met dans `CDP_PORTS`.
- Un flux unaire HTTP/2 qui ne se termine pas n'est toujours pas journalisé.
  Les trames WS, elles, partent au fil de l'eau (le socket reste ouvert).
