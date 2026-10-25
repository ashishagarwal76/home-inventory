import test from "node:test";
import assert from "node:assert/strict";
import { CameraSession, parseInventoryID, parseReturnURL, buildRecordURL } from "../scanner/scanner-core.mjs";

const appURL = "https://script.google.com/macros/s/test_script-123/exec";
const tick = () => new Promise(resolve => setImmediate(resolve));

test("inventory IDs accept printed and legacy IDs but reject unrelated QR content", () => {
  for (const id of ["ITEM-00001", "ITEM-001", "BOX-001", "LOC-002"]) assert.equal(parseInventoryID(id), id);
  for (const id of [null, {}, "", "https://example.com", "<img src=x>", "ITEM-1&x=y", "ITEM-123456789012345678901"]) {
    assert.equal(parseInventoryID(id), null);
  }
});

test("return addresses only allow expected Google Apps Script hosts and deployment paths", () => {
  assert.equal(parseReturnURL(appURL + "?page=old&record=old#fragment").href, appURL + "?page=scan");
  assert.ok(parseReturnURL("https://script.google.com/a/macros/example.com/s/test_script/dev"));
  for (const value of ["javascript:alert(1)", "https://example.com", "http://script.google.com/macros/s/test/exec", "https://script.google.com.evil.test/macros/s/test/exec", "https://user@script.google.com/macros/s/test/exec", "https://script.google.com:8443/macros/s/test/exec", "https://script.google.com/anything"]) {
    assert.equal(parseReturnURL(value), null);
  }
});

test("return navigation carries the exact ID and discards previous query values", () => {
  assert.equal(buildRecordURL(appURL + "?unused=1", "ITEM-00001"), appURL + "?page=scan&record=ITEM-00001");
  assert.equal(buildRecordURL(appURL, "https://evil.test"), null);
});

function fixture(options = {}) {
  const matches = [], states = [], errors = [];
  let stops = 0, onDecoded;
  const scanner = {
    isScanning: false,
    async start(camera, config, onSuccess) {
      onDecoded = onSuccess;
      if (options.start) await options.start();
      this.isScanning = true;
    },
    async stop() {
      stops++;
      if (options.stop) await options.stop();
      this.isScanning = false;
    }
  };
  const session = new CameraSession({
    createScanner: () => scanner, config: {},
    onState: state => states.push(state), onMatch: id => matches.push(id),
    onInvalid() {}, onError: error => errors.push(error)
  });
  return { session, scanner, matches, states, errors, decode: text => onDecoded(text), get stops() { return stops; } };
}

test("duplicate successful frames stop once and return one record ID", async () => {
  const f = fixture();
  await f.session.start();
  f.decode("BOX-001");
  f.decode("BOX-001");
  await tick();
  assert.equal(f.stops, 1);
  assert.equal(f.scanner.isScanning, false);
  assert.deepEqual(f.matches, ["BOX-001"]);
});

test("stopping while camera permission is pending releases the acquired stream", async () => {
  let allowCamera;
  const f = fixture({ start: () => new Promise(resolve => { allowCamera = resolve; }) });
  const starting = f.session.start();
  await tick();
  const stopping = f.session.stop();
  allowCamera();
  await Promise.all([starting, stopping]);
  assert.equal(f.scanner.isScanning, false);
  assert.equal(f.stops, 1);
  assert.equal(f.session.state, "idle");
  assert.deepEqual(f.matches, []);
});

test("concurrent stop requests complete consistently and a fresh scan can start", async () => {
  const f = fixture();
  await f.session.start();
  assert.deepEqual(await Promise.all([f.session.stop(), f.session.stop()]), [true, true]);
  assert.equal(f.stops, 1);
  await f.session.start();
  assert.equal(f.scanner.isScanning, true);
  await f.session.stop();
});

test("permission failure leaves scanning retryable", async () => {
  let deny = true;
  const f = fixture({ start: async () => { if (deny) throw new Error("Permission denied"); } });
  await f.session.start();
  assert.equal(f.session.state, "idle");
  assert.equal(f.errors.length, 1);
  deny = false;
  await f.session.start();
  assert.equal(f.session.state, "scanning");
  await f.session.stop();
});

test("failed camera cleanup prevents returning a successful scan", async () => {
  const f = fixture({ stop: async () => { throw new Error("Stop failed"); } });
  await f.session.start();
  f.decode("ITEM-00001");
  await tick();
  assert.deepEqual(f.matches, []);
  assert.equal(f.session.state, "blocked");
  assert.equal(f.errors.length, 1);
});
