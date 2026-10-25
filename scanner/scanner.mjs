import { CameraSession, parseInventoryID, parseReturnURL, buildRecordURL } from "./scanner-core.mjs";

const byID = id => document.getElementById(id);
const returnURL = parseReturnURL(new URL(location.href).searchParams.get("return"));

function show(message, type = "") {
  byID("status").textContent = message;
  byID("status").className = "status" + (type ? " " + type : "");
}

function showRecord(id) {
  byID("resultTitle").textContent = id;
  const types = { ITEM: "Item label", BOX: "Box label", LOC: "Location label" };
  byID("resultType").textContent = types[id.split("-")[0]];
  const recordURL = returnURL ? buildRecordURL(returnURL, id) : null;
  byID("openRecord").hidden = !recordURL;
  if (recordURL) byID("openRecord").href = recordURL;
  else byID("openRecord").removeAttribute("href");
  byID("result").hidden = false;
  show("Label read. Open its inventory record to continue.", "success");
  if (!recordURL) show("Label read. Open the scanner from your inventory app to view this record.", "success");
}

function cameraError(error) {
  const name = error?.name || "";
  const text = String(error?.message || error || "");
  if (name === "NotAllowedError" || /permission|denied|notallowed/i.test(text)) {
    show("Camera access was blocked. Allow camera access for this page in your browser, then try again. You can also enter the printed ID below.", "error");
  } else if (name === "NotFoundError" || /notfound|no camera|not found/i.test(text)) {
    show("No camera was found. Open this page on a phone, or enter the printed ID below.", "error");
  } else {
    show("The camera could not start or stop. Close other camera apps and try again. If the camera remains on, close this tab. You can also enter the printed ID below.", "error");
  }
}

const session = new CameraSession({
  createScanner() {
    if (typeof window.Html5Qrcode !== "function") {
      throw new Error("Scanner library did not load.");
    }
    return new window.Html5Qrcode("reader", {
      formatsToSupport: [window.Html5QrcodeSupportedFormats.QR_CODE],
      verbose: false
    });
  },
  config: {
    fps: 8,
    qrbox(width, height) {
      const size = Math.max(1, Math.floor(Math.min(width, height, 360) * 0.75));
      return { width: size, height: size };
    }
  },
  onState(state) {
    const active = ["starting", "scanning", "stopping", "blocked"].includes(state);
    byID("startButton").disabled = active;
    byID("againButton").disabled = active;
    byID("stopButton").disabled = state !== "scanning" && state !== "starting" && state !== "blocked";
    byID("manualButton").disabled = state === "starting" || state === "stopping";
    byID("cameraHint").hidden = active;
    if (state === "starting") show("Opening camera…");
    else if (state === "scanning") show("Hold one QR label inside the scanning area.");
    else if (state === "stopping") show("Closing camera…");
  },
  onMatch: showRecord,
  onInvalid() {
    show("This QR code is not an inventory label. Scan an ITEM, BOX, or LOC label.", "error");
  },
  onError: cameraError
});

byID("returnHint").hidden = !!returnURL;
if (returnURL) {
  byID("backLink").href = returnURL.href;
  byID("backLink").hidden = false;
}

function startCamera() {
  if (!window.isSecureContext || !navigator.mediaDevices?.getUserMedia) {
    show("Open this scanner using its HTTPS address in your phone's browser. You can also enter the printed ID below.", "error");
    return;
  }
  if (typeof window.Html5Qrcode !== "function") {
    show("The scanner library did not load. Check your connection and refresh, or enter the printed ID below.", "error");
    return;
  }
  byID("result").hidden = true;
  void session.start();
}

byID("startButton").addEventListener("click", startCamera);
byID("againButton").addEventListener("click", startCamera);
byID("stopButton").addEventListener("click", async () => {
  if (await session.stop()) show("Camera stopped. Tap Start camera when you are ready.");
});
byID("manualForm").addEventListener("submit", async event => {
  event.preventDefault();
  const id = parseInventoryID(byID("manualID").value.trim().toUpperCase());
  if (!id) {
    show("Enter the ID printed on the label, such as ITEM-00001, BOX-001, or LOC-001.", "error");
    return;
  }
  if (await session.stop()) showRecord(id);
});

document.addEventListener("visibilitychange", () => {
  if (document.hidden && (session.wanted || session.scanner?.isScanning)) {
    void session.stop().then(stopped => {
      if (stopped) show("Camera stopped while you were away. Tap Start camera to resume.");
    });
  }
});
window.addEventListener("pagehide", () => {
  session.wanted = false;
  for (const video of byID("reader").querySelectorAll("video")) {
    video.srcObject?.getTracks().forEach(track => track.stop());
  }
});
