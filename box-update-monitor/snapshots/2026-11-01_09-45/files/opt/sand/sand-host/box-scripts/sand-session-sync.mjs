import { mkdirSync, readFileSync, renameSync, statSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { pathToFileURL } from "node:url";
import {
  attachToPage,
  CDP_PORT_BASE,
  connectBrowser,
  cookieDigest,
  cookieFingerprint,
  cookieKey,
  cookieRecency,
  deleteCookies,
  detachSession,
  discoverMonitorPorts,
  hostMatchesCookieDomain,
  hostOfOrigin,
  isConnectionRefused,
  isSignedOutCopy,
  listPageTargets,
  parseSignedOutLedger,
  pushCookies,
  readCookies,
  readPageStorage,
  reloadPage,
  serializeSignedOutLedger,
  setPageStorage,
  SIGNED_OUT_LEDGER_PATH,
  toCookieParam,
} from "./cdp-cookies.mjs";

const POLL_INTERVAL_MS = Number.parseInt(process.env.SAND_SESSION_SYNC_INTERVAL_MS ?? "1500", 10);

const STORAGE_SEP = "\u0000";
export function makeStorageKey(origin, key) {
  return `${origin}${STORAGE_SEP}${key}`;
}
export function splitStorageKey(storageKey) {
  const i = storageKey.indexOf(STORAGE_SEP);
  return [storageKey.slice(0, i), storageKey.slice(i + 1)];
}

export function mergeStorage(perMonitor) {
  const canonical = new Map();
  for (const origins of perMonitor) {
    for (const [origin, items] of origins) {
      for (const [key, value] of items) {
        const storageKey = makeStorageKey(origin, key);
        const existing = canonical.get(storageKey);
        if (existing == null || (existing === "" && value !== "")) {
          canonical.set(storageKey, value);
        }
      }
    }
  }
  return canonical;
}

export function selectStorageSeed(origin, canonical, currentItems, offered) {
  const entries = [];
  const offeredKeys = [];
  for (const [storageKey, value] of canonical) {
    const [o, key] = splitStorageKey(storageKey);
    if (o !== origin) continue;
    if (currentItems.has(key)) continue;
    if (offered?.get(storageKey) === value) continue;
    entries.push([key, value]);
    offeredKeys.push([storageKey, value]);
  }
  return { entries, offeredKeys };
}

export function mergeCookies(perMonitor) {
  const canonical = new Map();
  for (const cookies of perMonitor) {
    for (const [key, cookie] of cookies) {
      const existing = canonical.get(key);
      if (existing == null || cookieRecency(cookie) > cookieRecency(existing)) {
        canonical.set(key, cookie);
      }
    }
  }
  return canonical;
}

export function selectCookieSeed(canonical, currentCookies, offered) {
  const cookies = [];
  for (const [key, cookie] of canonical) {
    if (currentCookies.has(key)) continue;
    if (offered?.get(key) === cookieFingerprint(cookie)) continue;
    cookies.push(cookie);
  }
  return cookies;
}

const SIGNED_OUT_LEDGER_MAX_KEYS = 512;
const SIGNED_OUT_LEDGER_MAX_DIGESTS = 16;
const BOX_CHROME_POLICY_PATH = "/etc/opt/chrome/policies/managed/sand-webauthn.json";

function cookieHasExpired(cookie, nowMs) {
  return (
    cookie.session !== true &&
    typeof cookie.expires === "number" &&
    cookie.expires > 0 &&
    cookie.expires * 1000 <= nowMs
  );
}

function addToSetMap(map, key, value) {
  let set = map.get(key);
  if (set == null) {
    set = new Set();
    map.set(key, set);
  }
  set.add(value);
}

function addToListMap(map, key, value) {
  let list = map.get(key);
  if (list == null) {
    list = [];
    map.set(key, list);
  }
  list.push(value);
}

function liveCookies(jar, ledger) {
  const live = new Map();
  for (const [key, cookie] of jar) {
    if (!isSignedOutCopy(ledger, key, cookie)) live.set(key, cookie);
  }
  return live;
}

function planCookieSync({ views, reads, ledger, now, restoresSessions }) {
  const continuing = [];
  for (const [port, read] of reads) {
    const view = views.get(port);
    if (view?.jar != null && view.browserId === read.browserId) continuing.push([port, view]);
  }

  const removed = new Map();
  const written = new Map();
  const retired = new Map();
  for (const [port, view] of continuing) {
    const { jar } = reads.get(port);
    for (const [key, before] of view.jar) {
      if (jar.has(key)) continue;
      view.sent.delete(key);
      if (cookieHasExpired(before, now)) continue;
      if (before.session === true && !restoresSessions) continue;
      addToListMap(removed, key, before);
    }
    for (const [key, cookie] of jar) {
      const fingerprint = cookieFingerprint(cookie);
      const before = view.jar.get(key);
      if (before != null && cookieFingerprint(before) === fingerprint) continue;
      if (view.sent.get(key) === fingerprint) continue;
      if (before != null && isSignedOutCopy(ledger, key, before)) {
        addToListMap(retired, key, cookie);
      } else {
        addToSetMap(written, key, cookieDigest(cookie));
      }
    }
  }

  let ledgerChanged = false;
  for (const [key, digests] of written) {
    const entry = ledger.get(key);
    if (entry == null) continue;
    for (const digest of digests) ledgerChanged = entry.digests.delete(digest) || ledgerChanged;
    if (entry.digests.size === 0) ledger.delete(key);
  }

  for (const [key, gone] of removed) {
    for (const copy of gone) addToListMap(retired, key, copy);
    for (const view of views.values()) {
      const before = view.jar?.get(key);
      if (before != null) addToListMap(retired, key, before);
    }
    for (const read of reads.values()) {
      const current = read.jar.get(key);
      if (current != null) addToListMap(retired, key, current);
    }
  }
  for (const [key, copies] of retired) {
    const entry = ledger.get(key) ?? { digests: new Set(), at: now };
    entry.at = now;
    for (const copy of copies) {
      const digest = cookieDigest(copy);
      if (entry.digests.has(digest)) continue;
      entry.digests.add(digest);
      ledgerChanged = true;
    }
    if (entry.digests.size > 0) ledger.set(key, entry);
  }

  const canonical = mergeCookies([...reads.values()].map((read) => liveCookies(read.jar, ledger)));
  const deletes = new Map();
  const pushes = new Map();
  for (const [port, read] of reads) {
    for (const [key, cookie] of read.jar) {
      if (isSignedOutCopy(ledger, key, cookie)) addToListMap(deletes, port, cookie);
    }
    const sent = views.get(port)?.browserId === read.browserId ? views.get(port).sent : null;
    const cookies = selectCookieSeed(canonical, liveCookies(read.jar, ledger), sent);
    if (cookies.length > 0) pushes.set(port, cookies);
  }
  return { signedOut: [...removed.keys()], ledgerChanged, deletes, pushes };
}

function capSignedOutLedger(ledger) {
  for (const entry of ledger.values()) {
    if (entry.digests.size <= SIGNED_OUT_LEDGER_MAX_DIGESTS) continue;
    entry.digests = new Set([...entry.digests].slice(-SIGNED_OUT_LEDGER_MAX_DIGESTS));
  }
  if (ledger.size <= SIGNED_OUT_LEDGER_MAX_KEYS) return;
  const oldest = [...ledger].sort((a, b) => a[1].at - b[1].at);
  for (const [key] of oldest.slice(0, ledger.size - SIGNED_OUT_LEDGER_MAX_KEYS)) ledger.delete(key);
}

function mergeSignedOutLedger(ledger, file) {
  for (const [key, { digests, at }] of file) {
    const entry = ledger.get(key);
    if (entry == null) {
      ledger.set(key, { digests: new Set(digests), at });
    } else {
      entry.digests = new Set([...digests, ...entry.digests]);
      entry.at = Math.max(entry.at, at);
    }
  }
  for (const [key, { digests }] of ledger) {
    const saved = file.get(key)?.digests;
    if (saved == null || [...digests].some((digest) => !saved.has(digest))) return true;
  }
  return false;
}

function ledgerStamp({ ino, size, mtimeMs }) {
  return `${ino}:${size}:${mtimeMs}`;
}

export const MONITOR_BUSY_TTL_MS = 30_000;

export function readBusyMonitorPorts(ports, now = Date.now()) {
  const busy = new Set();
  for (const port of ports) {
    try {
      const { mtimeMs } = statSync(`/tmp/sand-monitor-busy-${port - CDP_PORT_BASE}`);
      if (now - mtimeMs < MONITOR_BUSY_TTL_MS) busy.add(port);
    } catch (error) {
      if (error?.code !== "ENOENT") {
        log(`busy marker for CDP port ${port} unreadable: ${String(error)}`);
      }
    }
  }
  return busy;
}

export const RELOAD_WINDOW_MS = 60_000;
export const RELOAD_MAX_PER_WINDOW = 3;
export const RELOAD_QUIET_MS = POLL_INTERVAL_MS * 10 + 5_000;

export class FreshIdentityReloadBreaker {
  constructor(now = Date.now) {
    this.now = now;
    this.byPort = new Map();
  }

  stateFor(port, host) {
    let perHost = this.byPort.get(port);
    if (perHost == null) {
      perHost = new Map();
      this.byPort.set(port, perHost);
    }
    let state = perHost.get(host);
    if (state == null) {
      state = { issued: [], open: false, lastRequestAt: null, pending: false };
      perHost.set(host, state);
    }
    return state;
  }

  clearPort(port) {
    this.byPort.delete(port);
  }

  hasPendingDeferrals(livePorts) {
    for (const port of livePorts) {
      const perHost = this.byPort.get(port);
      if (perHost == null) continue;
      for (const state of perHost.values()) {
        if (state.open && state.pending) return true;
      }
    }
    return false;
  }

  request(port, host) {
    const t = this.now();
    const state = this.stateFor(port, host);
    const quietForMs = state.lastRequestAt == null ? Infinity : t - state.lastRequestAt;
    state.lastRequestAt = t;
    if (state.open) {
      if (quietForMs < RELOAD_QUIET_MS) {
        state.pending = true;
        return "open";
      }
      state.open = false;
      state.pending = false;
      state.issued = [t];
      return "allow-after-quiet";
    }
    state.issued = state.issued.filter((at) => t - at < RELOAD_WINDOW_MS);
    if (state.issued.length >= RELOAD_MAX_PER_WINDOW) {
      state.open = true;
      return "trip";
    }
    state.issued.push(t);
    return "allow";
  }

  takeDueDeferredHosts(availableByPort) {
    const t = this.now();
    const due = [];
    for (const [port, perHost] of this.byPort) {
      for (const [host, state] of perHost) {
        if (!state.open || !state.pending) continue;
        if (state.lastRequestAt != null && t - state.lastRequestAt < RELOAD_QUIET_MS) {
          continue;
        }
        if (availableByPort != null) {
          const available = availableByPort.get(port);
          if (available == null || !available.has(host)) continue;
        }
        state.open = false;
        state.pending = false;
        state.issued = [t];
        due.push({ port, host });
      }
    }
    return due;
  }

  requeueDeferral(port, host) {
    const state = this.stateFor(port, host);
    state.open = true;
    state.pending = true;
  }
}

function log(message) {
  process.stderr.write(`sand-session-sync ${message}\n`);
}

// Chrome keeps an in-memory list of all cookies and reads the backing store only until it has fetched them all (https://chromium.googlesource.com/chromium/src/+/refs/tags/151.0.7922.169/net/cookies/cookie_monster.h), and the DevTools IndexedDB domain has no method that writes an entry (https://chromedevtools.github.io/devtools-protocol/tot/IndexedDB/).
export class SessionSyncer {
  constructor({
    ledgerPath = SIGNED_OUT_LEDGER_PATH,
    policyPath = BOX_CHROME_POLICY_PATH,
    fs = { mkdirSync, readFileSync, renameSync, statSync, writeFileSync },
    now = Date.now,
  } = {}) {
    this.browsers = new Map();
    this.cookieViews = new Map();
    this.storageOffered = new Map();
    this.browserIdByPort = new Map();
    this.reloadBreaker = new FreshIdentityReloadBreaker();
    this.busyWithheldReloads = new Map();
    this.ledgerPath = ledgerPath;
    this.policyPath = policyPath;
    this.fs = fs;
    this.now = now;
    this.ledger = new Map();
    this.ledgerStamp = null;
    this.ledgerReadFailed = false;
    this.ledgerSaveFailed = false;
    this.ledgerUnsaved = false;
    this.restoresSessions = null;
    this.policyReadFailed = false;
    this.loadLedger();
  }

  loadLedger() {
    let stamp;
    let text = null;
    try {
      stamp = ledgerStamp(this.fs.statSync(this.ledgerPath));
      if (stamp === this.ledgerStamp) return true;
      text = this.fs.readFileSync(this.ledgerPath, "utf8");
    } catch (error) {
      if (error?.code !== "ENOENT") {
        if (!this.ledgerReadFailed) {
          log(`sign-out ledger unreadable; keeping it and not saving: ${String(error)}`);
        }
        this.ledgerReadFailed = true;
        return false;
      }
      stamp = "absent";
    }
    this.ledgerReadFailed = false;
    if (stamp === this.ledgerStamp) return true;
    let file = text == null ? new Map() : parseSignedOutLedger(text);
    if (file == null) {
      log(`${this.ledgerPath} is not a sign-out ledger; replacing it at the next save`);
      file = new Map();
    }
    if (mergeSignedOutLedger(this.ledger, file)) this.ledgerUnsaved = true;
    this.ledgerStamp = stamp;
    return true;
  }

  readRestoresSessions() {
    let restores = false;
    let reason = "has no RestoreOnStartup: 1";
    try {
      restores = JSON.parse(this.fs.readFileSync(this.policyPath, "utf8"))?.RestoreOnStartup === 1;
      this.policyReadFailed = false;
    } catch (error) {
      if (error?.code !== "ENOENT" && this.restoresSessions != null) {
        if (!this.policyReadFailed) {
          log(`Chrome policy is unreadable (${String(error)}); keeping its last value`);
        }
        this.policyReadFailed = true;
        return this.restoresSessions;
      }
      reason = `is unreadable (${String(error)})`;
    }
    if (restores !== this.restoresSessions) {
      log(
        restores
          ? "Chrome keeps session cookies when a window closes; a lost one is a sign-out"
          : `Chrome policy ${reason}; a lost session cookie is not read as a sign-out`,
      );
    }
    this.restoresSessions = restores;
    return restores;
  }

  syncBrowserIdentity(port, browserId) {
    const prev = this.browserIdByPort.get(port);
    if (prev != null && browserId !== "" && prev !== browserId) {
      this.cookieViews.delete(port);
      this.storageOffered.delete(port);
      this.reloadBreaker.clearPort(port);
      this.busyWithheldReloads.delete(port);
      log(`monitor on CDP port ${port} was replaced; cleared offered state`);
    }
    if (browserId !== "") this.browserIdByPort.set(port, browserId);
  }

  saveLedger() {
    if (!this.loadLedger()) return;
    capSignedOutLedger(this.ledger);
    const tmp = `${this.ledgerPath}.tmp`;
    try {
      this.fs.mkdirSync(dirname(this.ledgerPath), { recursive: true });
      this.fs.writeFileSync(tmp, serializeSignedOutLedger(this.ledger));
      const stamp = ledgerStamp(this.fs.statSync(tmp));
      this.fs.renameSync(tmp, this.ledgerPath);
      this.ledgerStamp = stamp;
      this.ledgerSaveFailed = false;
      this.ledgerUnsaved = false;
    } catch (error) {
      if (!this.ledgerSaveFailed) log(`saving the sign-out ledger failed: ${String(error)}`);
      this.ledgerSaveFailed = true;
    }
  }

  async readJars(live) {
    const reads = new Map();
    for (const browser of live) {
      try {
        reads.set(browser.port, {
          browser,
          browserId: browser.browserId,
          jar: await readCookies(browser),
        });
      } catch {
        browser.fail(new Error("getCookies failed"));
      }
    }
    return reads;
  }

  async writeJar(read, view, planned) {
    const { browser } = read;
    const current = await readCookies(browser);
    const stale = [];
    const fill = [];
    for (const cookie of planned.pushes) {
      const held = current.get(cookieKey(cookie));
      if (held == null || isSignedOutCopy(this.ledger, cookieKey(cookie), held)) fill.push(cookie);
    }
    const filled = new Set(fill.map(cookieKey));
    for (const cookie of planned.deletes) {
      const key = cookieKey(cookie);
      const held = current.get(key);
      if (held != null && !filled.has(key) && isSignedOutCopy(this.ledger, key, held)) {
        stale.push(held);
      }
    }
    for (const cookie of stale) view.sent.delete(cookieKey(cookie));
    for (const cookie of fill) view.sent.set(cookieKey(cookie), cookieFingerprint(cookie));
    if (stale.length > 0) await deleteCookies(browser, stale);
    if (fill.length > 0) await pushCookies(browser, fill.map(toCookieParam));
    const after = await readCookies(browser);
    const jar = new Map(read.jar);
    for (const key of [...stale, ...fill].map(cookieKey)) {
      const landed = after.get(key);
      if (landed == null) jar.delete(key);
      else jar.set(key, landed);
    }
    view.jar = jar;
    const replaced = fill.filter((cookie) => current.has(cookieKey(cookie))).length;
    return { deleted: stale.length + replaced, fill };
  }

  async syncCookies(live) {
    const reads = await this.readJars(live);
    for (const [port, read] of reads) {
      if (this.cookieViews.get(port)?.browserId !== read.browserId) {
        this.cookieViews.set(port, { browserId: read.browserId, jar: null, sent: new Map() });
      }
    }
    this.loadLedger();
    const plan = planCookieSync({
      views: this.cookieViews,
      reads,
      ledger: this.ledger,
      now: this.now(),
      restoresSessions: this.readRestoresSessions(),
    });
    if (plan.signedOut.length > 0) {
      log(
        `read a sign-out of ${plan.signedOut.length} cookie(s); removing every copy from the ` +
          `other monitors`,
      );
    }
    const newCookieDomains = new Map();
    let pushed = 0;
    let deleted = 0;
    for (const [port, read] of reads) {
      const view = this.cookieViews.get(port);
      const planned = {
        deletes: plan.deletes.get(port) ?? [],
        pushes: plan.pushes.get(port) ?? [],
      };
      view.jar = read.jar;
      if (planned.deletes.length === 0 && planned.pushes.length === 0) continue;
      try {
        const done = await this.writeJar(read, view, planned);
        deleted += done.deleted;
        pushed += done.fill.length;
        if (done.fill.length > 0) {
          newCookieDomains.set(port, new Set(done.fill.map((cookie) => cookie.domain)));
        }
      } catch {
        read.browser.fail(new Error("cookie write failed"));
      }
    }
    if (plan.ledgerChanged) this.ledgerUnsaved = true;
    if (this.ledgerUnsaved) this.saveLedger();
    if (deleted > 0) log(`deleted ${deleted} signed-out cookie cop${deleted === 1 ? "y" : "ies"}`);
    if (pushed > 0) log(`mirrored ${pushed} cookie write(s) across ${reads.size} monitors`);
    return newCookieDomains;
  }

  recordStorageOffered(port, storageKey, value) {
    let perKey = this.storageOffered.get(port);
    if (perKey == null) {
      perKey = new Map();
      this.storageOffered.set(port, perKey);
    }
    perKey.set(storageKey, value);
  }

  async ensureConnections() {
    for (const port of discoverMonitorPorts()) {
      const existing = this.browsers.get(port);
      if (existing != null && !existing.isClosed) continue;
      try {
        const browser = await connectBrowser(port);
        this.browsers.set(port, browser);
        this.syncBrowserIdentity(port, browser.browserId);
        log(`connected to monitor on CDP port ${port}`);
      } catch (error) {
        if (!isConnectionRefused(error)) {
          log(`connecting to monitor on CDP port ${port} failed: ${String(error)}`);
        }
      }
    }
    for (const [port, browser] of this.browsers) {
      if (browser.isClosed) this.browsers.delete(port);
    }
  }

  async attachPages(browsers) {
    const byPort = new Map();
    for (const browser of browsers) {
      const pages = [];
      let targets;
      try {
        targets = await listPageTargets(browser);
      } catch {
        targets = [];
      }
      for (const { targetId } of targets) {
        const sessionId = await attachToPage(browser, targetId);
        if (sessionId == null) continue;
        const storage = await readPageStorage(browser, sessionId);
        if (storage == null) {
          await detachSession(browser, sessionId);
          continue;
        }
        pages.push({
          browser,
          sessionId,
          origin: storage.origin,
          items: storage.items,
        });
      }
      byPort.set(browser.port, pages);
    }
    return byPort;
  }

  async syncLocalStorage(pagesByPort) {
    const originsByPort = new Map();
    for (const [port, pages] of pagesByPort) {
      const origins = new Map();
      for (const page of pages) {
        let items = origins.get(page.origin);
        if (items == null) {
          items = new Map();
          origins.set(page.origin, items);
        }
        for (const [key, value] of page.items) items.set(key, value);
      }
      originsByPort.set(port, origins);
    }

    const canonical = mergeStorage([...originsByPort.values()]);
    const reloadHosts = new Map();
    let synced = 0;
    for (const [port, origins] of originsByPort) {
      const pages = pagesByPort.get(port) ?? [];
      for (const [origin, items] of origins) {
        const { entries, offeredKeys } = selectStorageSeed(
          origin,
          canonical,
          items,
          this.storageOffered.get(port),
        );
        if (entries.length === 0) continue;
        const originPage = pages.find((page) => page.origin === origin);
        if (originPage == null) continue;
        await setPageStorage(originPage.browser, originPage.sessionId, entries);
        for (const [storageKey, value] of offeredKeys) {
          this.recordStorageOffered(port, storageKey, value);
        }
        synced += entries.length;
        let hosts = reloadHosts.get(port);
        if (hosts == null) {
          hosts = new Set();
          reloadHosts.set(port, hosts);
        }
        hosts.add(hostOfOrigin(origin));
      }
    }
    if (synced > 0) {
      log(`mirrored ${synced} localStorage entr${synced === 1 ? "y" : "ies"} across monitors`);
    }
    return reloadHosts;
  }

  async tick() {
    await this.ensureConnections();
    const live = [...this.browsers.values()].filter((b) => !b.isClosed);
    const newCookieDomains = await this.syncCookies(live);
    if (live.length === 0) return;
    const livePorts = live.map((b) => b.port);
    if (
      live.length < 2 &&
      !this.reloadBreaker.hasPendingDeferrals(livePorts) &&
      !livePorts.some((port) => this.busyWithheldReloads.has(port))
    ) {
      return;
    }

    const pagesByPort = await this.attachPages(live);
    const busyPorts = readBusyMonitorPorts(livePorts);
    try {
      if (live.length >= 2) {
        const reloadHosts = await this.syncLocalStorage(pagesByPort);
        await this.reloadReceivers(pagesByPort, reloadHosts, newCookieDomains, busyPorts);
      }
      await this.deliverDeferredReloads(pagesByPort, busyPorts);
      await this.deliverWithheldReloads(pagesByPort, busyPorts);
    } finally {
      for (const pages of pagesByPort.values()) {
        for (const page of pages) await detachSession(page.browser, page.sessionId);
      }
    }
  }

  requestReload(port, host) {
    const status = this.reloadBreaker.request(port, host);
    if (status === "trip") {
      log(
        `reload breaker OPEN for ${host} on CDP port ${port}: ` +
          `${RELOAD_MAX_PER_WINDOW} session-sync reloads within ${RELOAD_WINDOW_MS}ms ` +
          `(page storage is not converging); suppressing further reloads until it ` +
          `pauses — cookie/localStorage seeding continues`,
      );
    } else if (status === "allow-after-quiet") {
      log(
        `reload breaker closed for ${host} on CDP port ${port} after a quiet ` +
          `period; session-sync reloads re-enabled`,
      );
    }
    return status === "allow" || status === "allow-after-quiet";
  }

  async deliverDeferredReloads(pagesByPort, busyPorts) {
    const availableByPort = new Map();
    for (const [port, pages] of pagesByPort) {
      if (busyPorts.has(port)) continue;
      const liveHosts = new Set();
      for (const page of pages) {
        if (!page.browser.isClosed) liveHosts.add(hostOfOrigin(page.origin));
      }
      availableByPort.set(port, liveHosts);
    }
    for (const { port, host } of this.reloadBreaker.takeDueDeferredHosts(availableByPort)) {
      let delivered = false;
      for (const page of pagesByPort.get(port) ?? []) {
        if (hostOfOrigin(page.origin) !== host || page.browser.isClosed) continue;
        await reloadPage(page.browser, page.sessionId);
        if (!page.browser.isClosed) delivered = true;
      }
      if (delivered) {
        log(
          `reload breaker closed for ${host} on CDP port ${port} after a quiet ` +
            `period; delivering the reload it deferred while open`,
        );
      } else {
        this.reloadBreaker.requeueDeferral(port, host);
      }
    }
  }

  async deliverWithheldReloads(pagesByPort, busyPorts) {
    for (const [port, hosts] of this.busyWithheldReloads) {
      if (busyPorts.has(port)) continue;
      const pages = pagesByPort.get(port);
      if (pages == null || pages.length === 0) continue;
      this.busyWithheldReloads.delete(port);
      for (const host of hosts) {
        if (!this.requestReload(port, host)) continue;
        log(
          `delivering the ${host} reload withheld while CDP port ${port}'s ` +
            `monitor was being driven`,
        );
        for (const page of pages) {
          if (hostOfOrigin(page.origin) !== host || page.browser.isClosed) continue;
          await reloadPage(page.browser, page.sessionId);
        }
      }
    }
  }

  async reloadReceivers(pagesByPort, reloadHosts, newCookieDomains, busyPorts) {
    for (const [port, pages] of pagesByPort) {
      const hosts = reloadHosts.get(port);
      const domains = newCookieDomains.get(port);
      if ((hosts == null || hosts.size === 0) && (domains == null || domains.size === 0)) {
        continue;
      }
      const allowedByHost = new Map();
      for (const page of pages) {
        const host = hostOfOrigin(page.origin);
        const cookieHit =
          domains != null && [...domains].some((d) => hostMatchesCookieDomain(host, d));
        const storageHit = hosts != null && hosts.has(host);
        if (!cookieHit && !storageHit) continue;
        let allowed = allowedByHost.get(host);
        if (allowed == null) {
          if (busyPorts.has(port)) {
            let withheld = this.busyWithheldReloads.get(port);
            if (withheld == null) {
              withheld = new Set();
              this.busyWithheldReloads.set(port, withheld);
            }
            withheld.add(host);
            log(
              `withholding the ${host} reload on CDP port ${port}: its monitor is ` +
                `being driven right now; seeding continues`,
            );
            allowed = false;
          } else {
            allowed = this.requestReload(port, host);
          }
          allowedByHost.set(host, allowed);
        }
        if (allowed) {
          await reloadPage(page.browser, page.sessionId);
        }
      }
    }
  }
}

async function main() {
  log("starting; mirroring cookies + localStorage across box monitors");
  const syncer = new SessionSyncer();
  for (;;) {
    try {
      await syncer.tick();
    } catch (error) {
      log(`tick failed: ${String(error)}`);
    }
    await new Promise((resolve) => setTimeout(resolve, POLL_INTERVAL_MS));
  }
}

if (process.argv[1] != null && import.meta.url === pathToFileURL(process.argv[1]).href) {
  void main();
}
