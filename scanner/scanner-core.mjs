export function parseInventoryID(value) {
  if (typeof value !== "string") return null;
  const id = value.trim();
  return /^(?:ITEM|BOX|LOC)-\d{1,20}$/.test(id) ? id : null;
}

export function parseReturnURL(value) {
  if (typeof value !== "string" || value.length > 2000) return null;
  try {
    const url = new URL(value);
    if (url.protocol !== "https:" || url.hostname !== "script.google.com" ||
        url.username || url.password || url.port ||
        !/^\/(?:macros\/s|a\/macros\/[a-zA-Z0-9.-]+\/s)\/[a-zA-Z0-9_-]+\/(?:exec|dev)$/.test(url.pathname)) {
      return null;
    }
    url.search = "";
    url.hash = "";
    url.searchParams.set("page", "scan");
    return url;
  } catch {
    return null;
  }
}

export function buildRecordURL(returnURL, recordID) {
  const safeReturn = parseReturnURL(String(returnURL));
  const safeID = parseInventoryID(recordID);
  if (!safeReturn || !safeID) return null;
  safeReturn.searchParams.set("record", safeID);
  return safeReturn.href;
}

export class CameraSession {
  constructor({ createScanner, config, onState, onMatch, onInvalid, onError }) {
    Object.assign(this, { createScanner, config, onState, onMatch, onInvalid, onError });
    this.scanner = null;
    this.starting = null;
    this.stopping = null;
    this.wanted = false;
    this.matched = false;
    this.state = "idle";
  }

  setState(state) {
    this.state = state;
    this.onState(state);
  }

  async start() {
    if (this.wanted || this.starting || this.stopping || this.state === "blocked") return;
    this.wanted = true;
    this.matched = false;
    this.setState("starting");
    try {
      this.scanner ||= this.createScanner();
      this.starting = Promise.resolve().then(() => this.scanner.start(
        { facingMode: "environment" }, this.config,
        text => { void this.accept(text); }, () => {}
      ));
      await this.starting;
      if (this.wanted && !this.matched) this.setState("scanning");
    } catch (error) {
      this.wanted = false;
      this.setState("idle");
      this.onError(error);
    } finally {
      this.starting = null;
    }
  }

  async stop() {
    this.wanted = false;
    if (this.stopping) return this.stopping;
    this.setState("stopping");
    this.stopping = Promise.resolve().then(async () => {
      try {
        if (this.starting) {
          try { await this.starting; } catch { /* No camera was acquired. */ }
        }
        if (this.scanner?.isScanning) await this.scanner.stop();
        this.setState("idle");
        return true;
      } catch (error) {
        this.setState("blocked");
        this.onError(error);
        return false;
      } finally {
        this.stopping = null;
      }
    });
    return this.stopping;
  }

  async accept(text) {
    if (!this.wanted || this.matched) return;
    const id = parseInventoryID(text);
    if (!id) {
      this.onInvalid();
      return;
    }
    this.matched = true;
    if (await this.stop()) this.onMatch(id);
  }
}
