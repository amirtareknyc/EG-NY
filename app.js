/*
  Simple XLSX reader for this project.
  It reads database.xlsx directly in the browser and supports the
  normal XLSX ZIP structure used by the included spreadsheet.
  No external CDN/library is required.
*/

const CONFIG = {
  databaseFile: "database.xlsx",
  appointmentUrl: "https://example.com/book-appointment"
};

const form = document.getElementById("checkForm");
const requestIdInput = document.getElementById("requestId");
const phoneInput = document.getElementById("phone");
const button = document.getElementById("checkButton");
const message = document.getElementById("message");
const results = document.getElementById("results");
const offlineLoader = document.getElementById("offlineLoader");
const databasePicker = document.getElementById("databasePicker");

let databaseRows = null;

databasePicker.addEventListener("change", async () => {
  const file = databasePicker.files && databasePicker.files[0];
  if (!file) return;
  try {
    databaseRows = await loadXlsxFromBuffer(await file.arrayBuffer());
    offlineLoader.hidden = true;
    showMessage("Database loaded. You can now check the request.", "success");
  } catch (error) {
    console.error(error);
    showMessage("The selected Excel database could not be read.", "error");
  }
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  clearOutput();

  const requestId = requestIdInput.value.trim();
  const phone = phoneInput.value.trim();

  if (!requestId || !phone) {
    showMessage("Please enter both the request ID and phone number.", "error");
    return;
  }

  button.disabled = true;
  button.textContent = "Checking...";

  try {
    if (!databaseRows) {
      try {
        databaseRows = await loadXlsx(CONFIG.databaseFile);
      } catch (loadError) {
        offlineLoader.hidden = false;
        throw loadError;
      }
    }

    const wantedId = normalizeText(requestId);
    const wantedPhone = normalizePhone(phone);

    const matches = databaseRows.filter(row =>
      normalizeText(row.request_id) === wantedId &&
      normalizePhone(row.phone_number) === wantedPhone &&
      normalizeText(row.request_status) === "completed"
    );

    if (matches.length === 0) {
      showMessage(
        "There are no completed consular requests that match the data you've entered. " +
        "Please contact with us.",
        "error"
      );
      return;
    }

    showMessage("Completed request(s) found.", "success");
    renderResults(matches);
  } catch (error) {
    console.error(error);
    showMessage(
      "Unable to read the database. Make sure database.xlsx is in the same folder as this page.",
      "error"
    );
  } finally {
    button.disabled = false;
    button.textContent = "Check";
  }
});

function clearOutput() {
  message.className = "message";
  message.textContent = "";
  results.innerHTML = "";
}

function showMessage(text, type) {
  message.className = "message " + type;
  message.textContent = text;
}

function normalizeText(value) {
  return String(value ?? "").trim().toLowerCase();
}

function normalizePhone(value) {
  // Makes "+1 (212) 555-0101" and "1 212 555 0101" compare consistently.
  return String(value ?? "").replace(/\D/g, "");
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, ch => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;"
  }[ch]));
}

function renderResults(rows) {
  rows.forEach(row => {
    const div = document.createElement("div");
    div.className = "result";
    div.innerHTML = `
      <div class="result-row"><div class="result-label">Name</div><div>${escapeHtml(row.name)}</div></div>
      <div class="result-row"><div class="result-label">Request Type</div><div>${escapeHtml(row.request_type)}</div></div>
      <div class="result-row"><div class="result-label">Request Status</div><div>${escapeHtml(row.request_status)}</div></div>
      <div class="result-row"><div class="result-label">Notes</div><div>${escapeHtml(row.notes)}</div></div>
    `;
    results.appendChild(div);
  });

  const appointment = document.createElement("div");
  appointment.className = "appointment";
  appointment.innerHTML = `
    Please, visit our booking website to book an appointment to pickup your completed requests.
    <a href="${escapeHtml(CONFIG.appointmentUrl)}" target="_blank" rel="noopener noreferrer">
      Make an appointment now
    </a>
  `;
  results.appendChild(appointment);
}

/* ---------------- Minimal XLSX ZIP/XML reader ---------------- */

async function loadXlsx(url) {
  const response = await fetch(url, { cache: "no-store" });
  if (!response.ok) throw new Error("Database file could not be loaded.");
  return loadXlsxFromBuffer(await response.arrayBuffer());
}

async function loadXlsxFromBuffer(buffer) {
  const files = await unzip(buffer);

  const sharedStrings = files["xl/sharedStrings.xml"]
    ? parseSharedStrings(new TextDecoder().decode(files["xl/sharedStrings.xml"]))
    : [];

  const workbookXml = new TextDecoder().decode(files["xl/workbook.xml"]);
  const relsXml = new TextDecoder().decode(files["xl/_rels/workbook.xml.rels"]);
  const sheetPath = findFirstWorksheetPath(workbookXml, relsXml);

  if (!files[sheetPath]) throw new Error("Worksheet not found.");

  const sheetXml = new TextDecoder().decode(files[sheetPath]);
  const rows = parseWorksheet(sheetXml, sharedStrings);

  if (rows.length < 1) return [];
  const headers = rows[0].map(v => normalizeHeader(v));

  const required = ["request_id", "phone_number", "name", "request_type", "request_status", "notes"];
  const missing = required.filter(h => !headers.includes(h));
  if (missing.length) throw new Error("Missing database columns: " + missing.join(", "));

  return rows.slice(1)
    .filter(row => row.some(v => String(v ?? "").trim() !== ""))
    .map(row => {
      const obj = {};
      headers.forEach((header, i) => obj[header] = row[i] ?? "");
      return obj;
    });
}

function normalizeHeader(value) {
  return String(value ?? "").trim().toLowerCase().replace(/\s+/g, "_");
}

function xmlDoc(text) {
  return new DOMParser().parseFromString(text, "application/xml");
}

function parseSharedStrings(text) {
  const doc = xmlDoc(text);
  return Array.from(doc.getElementsByTagName("si")).map(si =>
    Array.from(si.getElementsByTagName("t")).map(t => t.textContent).join("")
  );
}

function findFirstWorksheetPath(workbookText, relsText) {
  const wb = xmlDoc(workbookText);
  const rels = xmlDoc(relsText);

  const sheet = wb.getElementsByTagName("sheet")[0];
  const rid = sheet?.getAttribute("r:id");
  const relList = Array.from(rels.getElementsByTagName("Relationship"));
  const rel = relList.find(r => r.getAttribute("Id") === rid);

  let target = rel?.getAttribute("Target") || "worksheets/sheet1.xml";
  target = target.replace(/^\/+/, "");
  if (!target.startsWith("xl/")) target = "xl/" + target;
  return target;
}

function parseWorksheet(text, sharedStrings) {
  const doc = xmlDoc(text);
  const rowNodes = Array.from(doc.getElementsByTagName("row"));
  const output = [];

  for (const rowNode of rowNodes) {
    const cells = Array.from(rowNode.getElementsByTagName("c"));
    const row = [];
    let currentIndex = 0;

    for (const cell of cells) {
      const ref = cell.getAttribute("r") || "";
      const colLetters = (ref.match(/[A-Z]+/) || ["A"])[0];
      const index = columnToIndex(colLetters);

      while (currentIndex < index) {
        row.push("");
        currentIndex++;
      }

      const type = cell.getAttribute("t");
      const v = cell.getElementsByTagName("v")[0];
      let value = v ? v.textContent : "";

      if (type === "s") {
        value = sharedStrings[Number(value)] ?? "";
      } else if (type === "inlineStr") {
        const t = cell.getElementsByTagName("t")[0];
        value = t ? t.textContent : "";
      }

      row[index] = value;
      currentIndex = index + 1;
    }

    output.push(row);
  }
  return output;
}

function columnToIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/*
  Basic ZIP reader:
  - reads the central directory
  - supports method 0 (stored) and method 8 (deflate)
  - uses the browser's native DecompressionStream for deflate
*/
async function unzip(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  const eocd = findSignatureFromEnd(bytes, 0x06054b50);

  if (eocd < 0) throw new Error("Invalid XLSX/ZIP file.");

  const centralOffset = view.getUint32(eocd + 16, true);
  const totalEntries = view.getUint16(eocd + 10, true);

  const files = {};
  let p = centralOffset;

  for (let i = 0; i < totalEntries; i++) {
    if (view.getUint32(p, true) !== 0x02014b50) {
      throw new Error("Invalid ZIP central directory.");
    }

    const method = view.getUint16(p + 10, true);
    const compressedSize = view.getUint32(p + 20, true);
    const fileNameLength = view.getUint16(p + 28, true);
    const extraLength = view.getUint16(p + 30, true);
    const commentLength = view.getUint16(p + 32, true);
    const localOffset = view.getUint32(p + 42, true);

    const nameBytes = bytes.slice(p + 46, p + 46 + fileNameLength);
    const name = new TextDecoder().decode(nameBytes);

    if (!name.endsWith("/")) {
      const localNameLength = view.getUint16(localOffset + 26, true);
      const localExtraLength = view.getUint16(localOffset + 28, true);
      const dataStart = localOffset + 30 + localNameLength + localExtraLength;
      const compressed = bytes.slice(dataStart, dataStart + compressedSize);

      let content;
      if (method === 0) {
        content = compressed;
      } else if (method === 8) {
        content = await inflateRaw(compressed);
      } else {
        throw new Error("Unsupported ZIP compression method: " + method);
      }

      files[name] = content;
    }

    p += 46 + fileNameLength + extraLength + commentLength;
  }

  return files;
}

function findSignatureFromEnd(bytes, signature) {
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (
      bytes[i] === (signature & 255) &&
      bytes[i + 1] === ((signature >>> 8) & 255) &&
      bytes[i + 2] === ((signature >>> 16) & 255) &&
      bytes[i + 3] === ((signature >>> 24) & 255)
    ) return i;
  }
  return -1;
}

async function inflateRaw(data) {
  if (!("DecompressionStream" in window)) {
    throw new Error("This browser does not support DecompressionStream. Please use a recent Chrome, Edge, Firefox, or Safari.");
  }

  const stream = new Blob([data]).stream().pipeThrough(
    new DecompressionStream("deflate-raw")
  );

  return new Uint8Array(await new Response(stream).arrayBuffer());
}
