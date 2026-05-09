//# allFunctionsCalledOnLoad

const SAMPLE_BYTES    = 16384;
const MAX_EXPORT_ROWS = 1_048_576; // Excel hard limit
const MAX_VIEW_ROWS   = 10_000;    // rows loaded into the view page
const JSON_WARN_BYTES = 50 * 1024 * 1024;
const ALLOWED      = ["csv", "txt", "json", "xml", "sxl"];
const DATA_ALLOWED = ["csv", "txt", "json", "xml"];
const DELIMITERS = [
  { char: "\t", label: "Tab" },
  { char: ",",  label: "Comma" },
  { char: ";",  label: "Semicolon" },
  { char: "|",  label: "Pipe" },
];
const DATA_TYPES = ["General", "Number", "Currency", "Accounting", "Date", "DateTime", "Time", "Percentage", "Fraction", "Scientific", "Text", "Boolean"];

// ── DOM refs ──────────────────────────────────────────────────────────────────
const card            = document.getElementById("card");
const dropZone        = document.getElementById("dropZone");
const fileInput       = document.getElementById("fileInput");
const errorMsg        = document.getElementById("errorMsg");
const tableSection    = document.getElementById("tableSection");
const tableBody       = document.getElementById("tableBody");
const resetBtn        = document.getElementById("resetBtn");
const updateFileInput = document.getElementById("updateFileInput");
const saveBtn         = document.getElementById("saveBtn");
const sxlNotice       = document.getElementById("sxlNotice");
const configSection   = document.getElementById("configSection");
const delimiterInput  = document.getElementById("delimiterInput");
const hasHeadersCheck = document.getElementById("hasHeadersCheck");
const fieldsPanel     = document.getElementById("fieldsPanel");
const activeList      = document.getElementById("activeList");
const excludedList    = document.getElementById("excludedList");
const bulkTypeSelect  = document.getElementById("bulkTypeSelect");
const applyTypeBtn    = document.getElementById("applyTypeBtn");
const viewBtn         = document.getElementById("viewBtn");
const exportBtn       = document.getElementById("exportBtn");
const exportMenu      = document.getElementById("exportMenu");
const exportCsvBtn    = document.getElementById("exportCsvBtn");
const exportExcelBtn  = document.getElementById("exportExcelBtn");
const samplePrevBtn      = document.getElementById("samplePrevBtn");
const sampleNextBtn      = document.getElementById("sampleNextBtn");
const sampleRowLabel     = document.getElementById("sampleRowLabel");
const qualifyHeadersCheck = document.getElementById("qualifyHeadersCheck");

// Cached selectors for controls that show/hide per file type
const delimiterFieldEl     = document.querySelector(".delimiter-field");
const hasHeadersLabelEl    = document.querySelector(".has-headers-label");
const qualifyHeadersLabelEl = document.querySelector(".qualify-headers-label");

// ── State ─────────────────────────────────────────────────────────────────────
const state = {
  currentFile:     null,
  sxlFileMeta:     null,   // { name, size, lastModified } when loaded from .sxl
  fileType:        null,   // "delimited" | "json"
  delimChar:       null,
  rawRows:         [],     // delimited: parsed sample rows
  jsonRawSample:   [],     // json: raw objects from first N records
  jsonRecords:     [],     // json: flattened rows from jsonRawSample
  activeColumns:   [],     // { header, type, originalIndex, key? }
  excludedColumns: [],
  sampleRowIndex:  0,
};

// ── Utilities ─────────────────────────────────────────────────────────────────
function ext(name) { return name.split(".").pop().toLowerCase(); }

function formatSize(bytes) {
  if (bytes === 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return (bytes / Math.pow(1024, i)).toFixed(i === 0 ? 0 : 1) + " " + units[i];
}

function formatDate(ts) {
  if (!ts) return null;
  const d = new Date(ts);
  return d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" })
       + " " + d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
}

function escHtml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function showError(msg) {
  errorMsg.textContent = msg;
  errorMsg.style.display = "block";
  tableSection.style.display = "none";
}

function clearError() { errorMsg.style.display = "none"; }

// ── File reading ──────────────────────────────────────────────────────────────
function readSample(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload  = (e) => resolve(e.target.result);
    reader.onerror = ()  => resolve("");
    reader.readAsText(file.slice(0, SAMPLE_BYTES));
  });
}

function readFullFile(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload  = (e) => resolve(e.target.result);
    reader.onerror = ()  => resolve("");
    reader.readAsText(file);
  });
}

// ── Delimiter detection ───────────────────────────────────────────────────────
function detectDelimiterChar(sample) {
  const lines = sample.split(/\r?\n/).filter(l => l.trim().length > 0).slice(0, 20);
  if (lines.length < 2) return null;

  let bestChar = null, bestScore = 0;
  for (const { char } of DELIMITERS) {
    const counts   = lines.map(l => l.split(char).length - 1);
    const avg      = counts.reduce((a, b) => a + b, 0) / counts.length;
    if (avg < 1) continue;
    const variance = counts.reduce((a, b) => a + Math.pow(b - avg, 2), 0) / counts.length;
    const score    = avg / (1 + variance);
    if (score > bestScore) { bestScore = score; bestChar = char; }
  }
  return bestChar;
}

function delimCharToLabel(char) {
  return DELIMITERS.find(d => d.char === char)?.label ?? char;
}

// ── Delimited sample parsing ──────────────────────────────────────────────────
function parseSampleRows(text, delimChar) {
  return text
    .split(/\r?\n/)
    .filter(l => l.trim().length > 0)
    .slice(0, 100)
    .map(l => l.split(delimChar));
}

// ── JSON flattening ───────────────────────────────────────────────────────────

// "projectName" -> "Project Name",  "sourcedGUID" -> "Sourced GUID",  "refAgentInstanceID" -> "Ref Agent Instance ID"
function camelToWords(str) {
  return str
    .replace(/_+/g, " ")                         // underscore_case -> spaces
    .replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2")  // keep acronyms together: "GUIDFoo" -> "GUID Foo"
    .replace(/([a-z\d])([A-Z])/g, "$1 $2")       // split camelCase: "fooBar" -> "foo Bar"
    .replace(/^[a-z]/, c => c.toUpperCase())
    .trim();
}

// Recursively flatten one plain object into an array of flat row objects.
//   parentKey  – the JSON key of the containing field, used as column prefix.
// Rules:
//   scalar          -> one column  "Parent Field"
//   nested object   -> recurse, merge columns into current row
//   primitive array -> join as "a, b, c" into one column
//   object array    -> recurse each item, expand into multiple rows (cartesian
//                     product with everything else at this level)
// qualify: true  -> always prefix with parent ("Timezone Name")
//          false -> use leaf name only; auto-qualifies only on collision using
//                  the *original* source key, not the nearest merge level
//
// Returns { rows: [{col: val}], sources: {col: originKey} }
// sources lets collisions at higher levels qualify with the true origin key
// (e.g. "timezone" not "profile") so the label is always meaningful.
function flattenObject(obj, parentKey, qualify) {
  const q       = qualify !== undefined ? qualify : (qualifyHeadersCheck?.checked ?? false);
  const scalars = {};
  const sources = {}; // col -> key that originally produced that col
  const groups  = []; // { rows, src } - src mirrors sources for that subtree

  const addScalar = (label, val, originKey) => {
    if (label in scalars) {
      // Collision: qualify using the stored origin key of the *incoming* value
      const qual = `${camelToWords(originKey)} ${label}`;
      scalars[qual] = val;
      sources[qual] = originKey;
    } else {
      scalars[label] = val;
      sources[label] = originKey;
    }
  };

  for (const [key, val] of Object.entries(obj)) {
    const wordKey = camelToWords(key);
    const label   = (q && parentKey) ? `${camelToWords(parentKey)} ${wordKey}` : wordKey;
    // origin = the key most directly responsible for this label's short name
    const origin  = parentKey || key;

    if (val === null || val === undefined || typeof val !== "object") {
      addScalar(label, val ?? null, origin);

    } else if (Array.isArray(val)) {
      if (val.length === 0) {
        addScalar(label, null, origin);
      } else if (typeof val[0] !== "object" || val[0] === null) {
        addScalar(label, val.map(String).join(", "), origin);
      } else {
        const itemResults = val
          .filter(item => item !== null && typeof item === "object" && !Array.isArray(item))
          .map(item => flattenObject(item, key, q));
        groups.push({
          rows: itemResults.flatMap(r => r.rows),
          src:  itemResults[0]?.sources ?? {},
        });
      }

    } else {
      const sub = flattenObject(val, key, q);
      if (sub.rows.length === 1) {
        for (const [subKey, subVal] of Object.entries(sub.rows[0])) {
          const srcKey = sub.sources[subKey] ?? key;
          if (subKey in scalars) {
            // Use the value's own origin key for the qualifier
            const qual = `${camelToWords(srcKey)} ${subKey}`;
            scalars[qual] = subVal;
            sources[qual] = srcKey;
          } else {
            scalars[subKey] = subVal;
            sources[subKey] = srcKey;
          }
        }
      } else {
        groups.push({ rows: sub.rows, src: sub.sources });
      }
    }
  }

  if (groups.length === 0) return { rows: [scalars], sources };

  // Cartesian product - resolve collisions using each group's source map
  // so the qualifier reflects the true origin, not the current merge level.
  let rows    = [scalars];
  let rowSrcs = [sources];

  for (const { rows: group, src: gSrc } of groups) {
    const next    = [];
    const nextSrc = [];
    for (let bi = 0; bi < rows.length; bi++) {
      const base    = rows[bi];
      const baseSrc = rowSrcs[bi];
      for (const ext of group) {
        const merged    = { ...base };
        const mergedSrc = { ...baseSrc };
        for (const [extKey, extVal] of Object.entries(ext)) {
          const srcKey = gSrc[extKey] ?? extKey;
          if (extKey in merged) {
            const qual = `${camelToWords(srcKey)} ${extKey}`;
            merged[qual]    = extVal;
            mergedSrc[qual] = srcKey;
          } else {
            merged[extKey]    = extVal;
            mergedSrc[extKey] = srcKey;
          }
        }
        next.push(merged);
        nextSrc.push(mergedSrc);
      }
    }
    rows    = next;
    rowSrcs = nextSrc;
  }

  return { rows, sources: rowSrcs[0] ?? {} };
}

// Async generator: yield one top-level JS object at a time from a JSON file
// without ever loading the whole file into memory.
// Handles root arrays ([...]) and single-level wrapper objects ({"key":[...]}).
async function* streamJsonObjects(file) {
  if (!file.stream) {
    // Fallback for browsers without the Streams API
    const text = await readFullFile(file);
    let data; try { data = JSON.parse(text); } catch { return; }
    let arr = Array.isArray(data) ? data : null;
    if (!arr && data && typeof data === "object") {
      for (const v of Object.values(data)) { if (Array.isArray(v)) { arr = v; break; } }
    }
    if (arr) for (const item of arr) {
      if (item && typeof item === "object" && !Array.isArray(item)) yield item;
    }
    return;
  }

  const decoder  = new TextDecoder();
  const reader   = file.stream().getReader();
  let inString   = false, escape = false;
  let rootType   = null;          // "array" | "object"
  let inArray    = false;         // inside the target array
  let wrapDepth  = 0;             // depth inside wrapper object
  let objDepth   = 0;             // depth inside collected object
  let acc        = "";            // character accumulator for current object

  try {
    let stop = false;
    while (!stop) {
      const { done, value } = await reader.read();
      if (done) break;

      const text    = decoder.decode(value, { stream: true });
      const pending = [];

      for (let i = 0; i < text.length && !stop; i++) {
        const ch = text[i];

        // String / escape handling (short-circuits the rest)
        if (escape)                  { escape = false; if (objDepth) acc += ch; continue; }
        if (ch === "\\" && inString) { escape = true;  if (objDepth) acc += ch; continue; }
        if (ch === '"')              { inString = !inString; if (objDepth) acc += ch; continue; }
        if (inString)                { if (objDepth) acc += ch; continue; }

        // Identify root structure
        if (rootType === null) {
          if      (ch === "[") { rootType = "array";  inArray = true; }
          else if (ch === "{") { rootType = "object"; wrapDepth = 1; }
          continue;
        }

        // Inside wrapper object - scan for the first array value
        if (rootType === "object" && !inArray) {
          if      (ch === "[" && wrapDepth === 1) { inArray = true; }
          else if (ch === "{" || ch === "[")      { wrapDepth++; }
          else if (ch === "}" || ch === "]")      { if (--wrapDepth <= 0) stop = true; }
          continue;
        }

        // Inside the target array - collect complete top-level objects
        if      (ch === "{" && objDepth === 0)               { objDepth = 1; acc = "{"; }
        else if ((ch === "{" || ch === "[") && objDepth > 0)  { objDepth++; acc += ch; }
        else if (ch === "}" && objDepth > 0) {
          acc += ch;
          if (--objDepth === 0) { try { pending.push(JSON.parse(acc)); } catch {} acc = ""; }
        }
        else if (ch === "]" && objDepth > 0)  { acc += ch; objDepth--; }
        else if (ch === "]" && objDepth === 0) { stop = true; }
        else if (objDepth > 0)                { acc += ch; }
      }

      for (const obj of pending) yield obj;
    }
  } finally {
    reader.releaseLock();
  }
}

// ── XML support ───────────────────────────────────────────────────────────────

// Scan the first 128 KB to detect the repeating record element tag name.
// Uses two strategies to handle both flat (<root><record/>) and wrapped
// (<root><container><record/></container></root>) XML structures.
async function detectXmlRecord(file) {
  const text = await new Promise(r => {
    const fr = new FileReader();
    fr.onload  = e => r(e.target.result);
    fr.onerror = ()  => r("");
    fr.readAsText(file.slice(0, 131072));
  });

  // Build depths[d] = {tagName: count} by simulating a tag stack.
  const stack  = [];
  const depths = [];
  const re     = /<(\/?[A-Za-z][A-Za-z0-9_.:-]*)[^>]*?(\/?)>/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const raw     = m[1];
    const isSelf  = m[2] === "/";
    const isClose = raw.startsWith("/");
    const tagName = isClose ? raw.slice(1) : raw;
    if (tagName.startsWith("?") || tagName.startsWith("!")) continue;
    if (!isClose) {
      const d = stack.length;
      if (!depths[d]) depths[d] = {};
      depths[d][tagName] = (depths[d][tagName] || 0) + 1;
      if (!isSelf) stack.push(tagName);
    } else if (stack.length) {
      stack.pop();
    }
  }

  function toResult(tagName) {
    const ci = tagName.indexOf(":");
    return { tagName, localName: ci >= 0 ? tagName.slice(ci + 1) : tagName };
  }

  // Strategy A: walk down from root through single-type levels.
  // The record is the last level with only ONE tag type before we hit the
  // record's own multi-type fields (id, name, position, …).
  // e.g.  company->employees(1 type)->employee(1 type)->{id,name,…}(multi) -> record=employee
  for (let d = 1; d < depths.length; d++) {
    if (Object.keys(depths[d] || {}).length <= 1) continue; // still in container chain
    const prevTypes = Object.keys(depths[d - 1] || {});
    if (prevTypes.length === 1) return toResult(prevTypes[0]);
    break;
  }

  // Strategy B: shallowest depth where any element appears 2+ times.
  for (let d = 1; d < depths.length; d++) {
    for (const [tag, count] of Object.entries(depths[d] || {})) {
      if (count >= 2) return toResult(tag);
    }
  }

  // Fallback: second tag in document order (original heuristic).
  const tags = [];
  const re2  = /<([A-Za-z][A-Za-z0-9_.:-]*)/g;
  while ((m = re2.exec(text)) !== null && tags.length < 2) {
    if (!m[1].startsWith("?") && !m[1].startsWith("!")) tags.push(m[1]);
  }
  return tags.length >= 2 ? toResult(tags[1]) : null;
}

// Convert a namespace-stripped DOM element into a plain JS object.
function xmlElementToObject(el) {
  const children = [...el.children];
  if (children.length === 0) {
    const text = el.textContent.trim();
    return text === "" ? null : text;
  }
  const groups = {};
  for (const child of children) {
    const name = child.localName;
    if (!groups[name]) groups[name] = [];
    groups[name].push(child);
  }
  const obj = {};
  for (const [name, list] of Object.entries(groups)) {
    if (list.length === 1) {
      const v = xmlElementToObject(list[0]);
      if (v !== null) obj[name] = v;
    } else {
      const vs = list.map(xmlElementToObject).filter(v => v !== null);
      if (vs.length) obj[name] = vs;
    }
  }
  return Object.keys(obj).length ? obj : null;
}

// Strip namespace prefixes from an XML string then parse it with DOMParser.
function xmlRecordToObject(xml) {
  const cleaned = xml
    .replace(/\s+xmlns(?::[A-Za-z0-9_]+)?=(?:"[^"]*"|'[^']*')/g, "")
    .replace(/(<\/?)([A-Za-z][A-Za-z0-9_.-]*):([A-Za-z][^\s>/]*)/g, "$1$3");
  try {
    const doc  = new DOMParser().parseFromString(cleaned, "application/xml");
    const root = doc.documentElement;
    if (!root || root.localName === "parsererror") return null;
    return xmlElementToObject(root);
  } catch { return null; }
}

// Extract complete record substrings from a buffer; return records + remaining fragment.
function splitXmlRecords(buf, openStr, closeStr) {
  const records = [];
  let pos = 0;
  while (true) {
    const si = buf.indexOf(openStr, pos);
    if (si === -1) {
      // No complete opening tag found. Check whether a PARTIAL opening tag
      // is straddling the end of this chunk (e.g. buffer ends with "<boo" and
      // the next chunk starts with "k>…"). Keep it so it isn't silently dropped.
      const tail     = buf.slice(pos);
      const scanFrom = Math.max(0, tail.length - (openStr.length - 1));
      const ltPos    = tail.indexOf("<", scanFrom);
      return { records, remaining: ltPos >= 0 ? tail.slice(ltPos) : "" };
    }

    // Verify the char after openStr is a tag boundary (not a longer tag name)
    const charAfter = buf[si + openStr.length];
    if (charAfter !== ">" && charAfter !== "/" && charAfter !== " " &&
        charAfter !== "\t" && charAfter !== "\n" && charAfter !== "\r") {
      pos = si + openStr.length;
      continue;
    }

    let depth = 1, cursor = si + openStr.length, found = false;
    while (depth > 0) {
      const oi = buf.indexOf(openStr, cursor);
      const ci = buf.indexOf(closeStr, cursor);
      if (ci === -1) break; // incomplete record - need more data

      if (oi !== -1 && oi < ci) {
        // Verify boundary before counting as open
        const ca = buf[oi + openStr.length];
        if (ca === ">" || ca === "/" || ca === " " || ca === "\t" || ca === "\n" || ca === "\r") {
          const gt = buf.indexOf(">", oi + openStr.length);
          if (gt === -1) break;
          if (buf[gt - 1] !== "/") depth++;
          cursor = gt + 1;
        } else {
          cursor = oi + openStr.length + 1;
        }
      } else {
        if (--depth === 0) { found = true; cursor = ci + closeStr.length; }
        else               { cursor = ci + closeStr.length; }
      }
    }
    if (!found) return { records, remaining: buf.slice(si) };
    records.push(buf.slice(si, cursor));
    pos = cursor;
  }
}

// Async generator: yield one record object at a time from a large XML file.
async function* streamXmlObjects(file) {
  const info = await detectXmlRecord(file);
  if (!info) return;

  const { tagName } = info;
  const openStr  = `<${tagName}`;
  const closeStr = `</${tagName}>`;

  if (!file.stream) {
    const text = await readFullFile(file);
    const { records } = splitXmlRecords(text, openStr, closeStr);
    for (const xml of records) { const obj = xmlRecordToObject(xml); if (obj) yield obj; }
    return;
  }

  const decoder = new TextDecoder();
  const reader  = file.stream().getReader();
  let buffer    = "";
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const { records, remaining } = splitXmlRecords(buffer, openStr, closeStr);
      buffer = remaining;
      for (const xml of records) { const obj = xmlRecordToObject(xml); if (obj) yield obj; }
    }
    buffer += decoder.decode();
    if (buffer) {
      const { records } = splitXmlRecords(buffer, openStr, closeStr);
      for (const xml of records) { const obj = xmlRecordToObject(xml); if (obj) yield obj; }
    }
  } finally {
    reader.releaseLock();
  }
}

// Dispatch to the right object streamer based on current file type.
async function* streamStructuredObjects(file) {
  if (state.fileType === "xml") yield* streamXmlObjects(file);
  else yield* streamJsonObjects(file);
}

// ── XML EAV-aware flattener (one flat row per record) ─────────────────────────
// IMS LIS / SIF XML uses Entity-Attribute-Value patterns throughout:
//   • {instanceName, instanceValue}  - key/value by name
//   • {*Type: {instanceValue:{textString}}, *Value/{formattedName}}  - typed value
//   • {*Name:"str", *Value:...}  - plain string key + value
// Instead of cartesian-product expansion (which explodes into hundreds of rows),
// we pivot each array into named columns using these patterns.

const XML_META_KEYS = new Set(["instanceIdentifier", "instanceVocabulary", "language"]);

// Walk a nested object looking for the nearest textString value.
function getDeepText(obj) {
  if (obj === null || obj === undefined) return null;
  if (typeof obj !== "object") return String(obj).trim() || null;
  if ("textString" in obj) return String(obj.textString ?? "").trim() || null;
  for (const v of Object.values(obj)) {
    if (typeof v === "object" && !Array.isArray(v)) {
      const t = getDeepText(v);
      if (t !== null) return t;
    }
  }
  return null;
}

// Shorten full-path keys to their last 2 words; resolve collisions by widening rightward.
// e.g. "Group Record Sourced GUID Sourced Id" -> "Sourced Id"
//      "Group Record Group Group Type Scheme"  -> "Type Scheme"
function _shortenXmlKeys(full) {
  const result  = {};
  const usedKeys = new Set();
  for (const [fullKey, val] of Object.entries(full)) {
    const words = fullKey.trim().split(/\s+/);
    let key = words.slice(-2).join(" ");
    let n = 2;
    while (usedKeys.has(key) && n < words.length) {
      n++;
      key = words.slice(-n).join(" ");
    }
    if (usedKeys.has(key)) key = fullKey; // last resort: full path
    result[key] = val;
    usedKeys.add(key);
  }
  return result;
}

// Produce exactly ONE flat {col: val} row from an XML-sourced JS object.
// qualify=true  -> full paths  ("Group Record Sourced GUID Sourced Id")
// qualify=false -> last 2 words ("Sourced Id")
function flattenXmlRecord(obj, qualify) {
  const out = {};
  _xmlFlatObj(obj, "", out);
  return qualify ? out : _shortenXmlKeys(out);
}

function _xmlFlatObj(val, prefix, out) {
  if (val === null || val === undefined) {
    if (prefix) out[prefix] = null;
    return;
  }
  if (typeof val !== "object") {
    if (prefix) out[prefix] = typeof val === "string" ? (val.trim() || null) : val;
    return;
  }
  if (Array.isArray(val)) { _xmlFlatArray(val, prefix, out); return; }

  // Unwrap {language?, textString} leaf wrappers
  const ks = Object.keys(val);
  if (ks.length > 0 && ks.every(k => k === "language" || k === "textString")) {
    if (prefix) out[prefix] = String(val.textString ?? "").trim() || null;
    return;
  }
  for (const [k, v] of Object.entries(val)) {
    if (XML_META_KEYS.has(k)) continue;
    const next = prefix ? `${prefix} ${camelToWords(k)}` : camelToWords(k);
    _xmlFlatObj(v, next, out);
  }
}

function _xmlFlatArray(items, prefix, out) {
  if (!items.length) { if (prefix) out[prefix] = null; return; }
  if (typeof items[0] !== "object") {
    if (prefix) out[prefix] = items.map(v => String(v ?? "")).join(", ");
    return;
  }
  // Try EAV patterns in order; fall back to first-item only
  if (_eavInstanceNameValue(items, prefix, out)) return;
  if (_eavTypeChild(items, prefix, out))         return;
  if (_eavNameValue(items, prefix, out))          return;
  _xmlFlatObj(items[0], prefix, out); // fallback
}

// Pattern 1: every item has {instanceName, instanceValue, ...}
function _eavInstanceNameValue(items, prefix, out) {
  if (!items.every(it => it && "instanceName" in it && "instanceValue" in it)) return false;
  for (const item of items) {
    const label = getDeepText(item.instanceName);
    if (!label) continue;
    const col = prefix ? `${prefix} [${label}]` : `[${label}]`;
    _xmlFlatObj(item.instanceValue, col, out);
  }
  return true;
}

// Pattern 2: every item has a *Type child whose instanceValue/textString is the discriminator
function _eavTypeChild(items, prefix, out) {
  if (!items.length) return false;
  const typeKey = Object.keys(items[0]).find(k =>
    k.endsWith("Type") &&
    items.every(it => it && k in it && it[k] && typeof it[k] === "object")
  );
  if (!typeKey) return false;
  for (const item of items) {
    const label = getDeepText(item[typeKey]?.instanceValue) ?? getDeepText(item[typeKey]);
    if (!label) continue;
    const col  = prefix ? `${prefix} [${label}]` : `[${label}]`;
    const rest = {};
    for (const [k, v] of Object.entries(item)) {
      if (k !== typeKey && !XML_META_KEYS.has(k)) rest[k] = v;
    }
    const restEntries = Object.entries(rest);
    if (restEntries.length === 1) { _xmlFlatObj(restEntries[0][1], col, out); }
    else if (restEntries.length > 1) { _xmlFlatObj(rest, col, out); }
  }
  return true;
}

// Pattern 3: every item has a plain-string *Name field (unique) + *Value field
function _eavNameValue(items, prefix, out) {
  if (!items.length) return false;
  const sample  = items[0];
  const nameKey = Object.keys(sample).find(k =>
    (k.endsWith("Name") || k.endsWith("name")) &&
    typeof sample[k] === "string" &&
    items.every(it => typeof it[k] === "string")
  );
  const valKey = Object.keys(sample).find(k => k.endsWith("Value") && k !== nameKey);
  if (!nameKey || !valKey) return false;
  const labels = items.map(it => String(it[nameKey] ?? "").trim());
  if (new Set(labels).size !== labels.length || labels.some(l => !l)) return false;
  for (const item of items) {
    const label = String(item[nameKey] ?? "").trim();
    if (!label) continue;

    // Drop the structural parent prefix - names like "parameterName/parameterValue"
    // are container boilerplate. Use the label itself as the column root.
    let colBase = camelToWords(label);
    let valObj  = item[valKey];

    // If the value is a single-key wrapper (e.g. parameterValue = {groupRecord: {...}}),
    // unwrap it so the inner key becomes the column root instead of double-prefixing.
    if (typeof valObj === "object" && !Array.isArray(valObj) && valObj !== null) {
      const innerKeys = Object.keys(valObj).filter(k => !XML_META_KEYS.has(k));
      if (innerKeys.length === 1) {
        colBase = camelToWords(innerKeys[0]);
        valObj  = valObj[innerKeys[0]];
      }
    }

    _xmlFlatObj(valObj, colBase, out);
  }
  return true;
}

// Dispatch: XML -> single EAV-pivoted row; JSON -> cartesian-product rows.
function flattenRecord(obj) {
  return state.fileType === "xml"
    ? [flattenXmlRecord(obj, qualifyHeadersCheck?.checked ?? false)]
    : flattenObject(obj, "").rows;
}

// Collect every unique key across the first N flattened rows (preserves order)
function collectJsonKeys(records, sampleSize = 50) {
  const seen = new Set();
  records.slice(0, sampleSize).forEach(r => Object.keys(r).forEach(k => seen.add(k)));
  return [...seen];
}

// ── Column type detection ─────────────────────────────────────────────────────
function detectColumnType(values) {
  const nonEmpty = values.map(v => String(v ?? "").trim()).filter(v => v !== "");
  if (nonEmpty.length === 0) return "General";

  const boolSet = new Set(["true", "false", "yes", "no", "1", "0", "y", "n"]);
  if (nonEmpty.every(v => boolSet.has(v.toLowerCase()))) return "Boolean";

  if (nonEmpty.every(v => /^-?\d+\.?\d*[eE][+\-]?\d+$/.test(v))) return "Scientific";

  if (nonEmpty.every(v => /^-?\d+\.?\d*\s*%$/.test(v))) return "Percentage";

  if (nonEmpty.every(v => /^[$€£¥]\s*-?\d[\d,]*(\.\d+)?$|^-?\d[\d,]*(\.\d+)?\s*[$€£¥]$/.test(v))) return "Currency";

  if (nonEmpty.every(v => /^-?(\d+\.?\d*|\d*\.?\d+)$/.test(v.replace(/,/g, "")))) return "Number";

  const dtRe   = /^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}/;
  if (nonEmpty.every(v => dtRe.test(v))) return "DateTime";

  const dateRe = /^(\d{4}-\d{2}-\d{2}|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})$/;
  if (nonEmpty.every(v => dateRe.test(v))) return "Date";

  const timeRe = /^\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?(\s*(AM|PM))?$/i;
  if (nonEmpty.every(v => timeRe.test(v))) return "Time";

  const fracRe = /^-?\d+\s+\d+\/[1-9]\d*$|^-?\d+\/[1-9]\d*$/;
  if (nonEmpty.every(v => fracRe.test(v))) return "Fraction";

  return "General";
}

// ── Column parsing ────────────────────────────────────────────────────────────
function parseColumnsDelimited() {
  const rows = state.rawRows;
  if (rows.length === 0) return;

  let headers, dataRows;
  if (hasHeadersCheck.checked) {
    headers  = rows[0].map(h => h.trim());
    dataRows = rows.slice(1);
  } else {
    const colCount = Math.max(...rows.map(r => r.length));
    headers  = Array.from({ length: colCount }, (_, i) => `Column ${i + 1}`);
    dataRows = rows;
  }

  state.activeColumns  = headers.map((h, i) => ({
    header:        h,
    type:          detectColumnType(dataRows.map(r => r[i] ?? "")),
    originalIndex: i,
  }));
  state.excludedColumns = [];
  state.sampleRowIndex  = 0;
}

function parseColumnsJson() {
  const records = state.jsonRecords;
  const keys    = collectJsonKeys(records);

  state.activeColumns = keys.map((key, i) => ({
    header:        key,
    type:          detectColumnType(records.slice(0, 50).map(r => String(r[key] ?? ""))),
    originalIndex: i,
    key,
  }));
  state.excludedColumns = [];
  state.sampleRowIndex  = 0;
}

// ── Sample data helpers ───────────────────────────────────────────────────────
function sampleDataRowCount() {
  if (state.fileType === "json" || state.fileType === "xml") return state.jsonRecords.length;
  return hasHeadersCheck.checked
    ? Math.max(0, state.rawRows.length - 1)
    : state.rawRows.length;
}

function getSampleValue(col) {
  if (state.fileType === "json" || state.fileType === "xml") {
    const record = state.jsonRecords[state.sampleRowIndex];
    if (!record) return "";
    return String(record[col.key] ?? "").trim();
  }
  const dataStart = hasHeadersCheck.checked ? 1 : 0;
  const row = state.rawRows[dataStart + state.sampleRowIndex];
  if (!row) return "";
  return String(row[col.originalIndex] ?? "").trim();
}

function updateSampleNav() {
  const total   = sampleDataRowCount();
  const current = total > 0 ? state.sampleRowIndex + 1 : 0;
  sampleRowLabel.innerHTML = total > 0
    ? `Sample Data [<span style="color:#3b82f6">${current}</span>]`
    : "-";
  samplePrevBtn.disabled = state.sampleRowIndex <= 0;
  sampleNextBtn.disabled = state.sampleRowIndex >= total - 1;
}

function refreshSampleCells() {
  document.querySelectorAll(".sample-cell").forEach((cell, i) => {
    const col = state.activeColumns[i];
    if (col) cell.textContent = getSampleValue(col);
  });
  updateSampleNav();
}

// ── Rendering ─────────────────────────────────────────────────────────────────
function renderFields() {
  activeList.innerHTML = state.activeColumns.map((col, i) => `
    <tr data-i="${i}">
      <td><input type="text" class="field-header-input" value="${escHtml(col.header)}" /></td>
      <td>
        <select class="field-type-select">
          ${DATA_TYPES.map(t => `<option value="${t}"${t === col.type ? " selected" : ""}>${t}</option>`).join("")}
        </select>
      </td>
      <td class="sample-cell">${escHtml(getSampleValue(col))}</td>
      <td class="order-cell">
        <button type="button" class="order-btn" data-i="${i}" data-dir="-1"${i === 0 ? " disabled" : ""}><img src="assets/images/up-arrow.png" class="order-arrow" alt="↑"></button>
        <button type="button" class="order-btn" data-i="${i}" data-dir="1"${i === state.activeColumns.length - 1 ? " disabled" : ""}><img src="assets/images/down-arrow.png" class="order-arrow" alt="↓"></button>
      </td>
      <td>
        <button type="button" class="remove-field-btn" data-i="${i}"><img src="assets/images/remove.png" class="remove-icon" alt="−"></button>
      </td>
    </tr>`).join("");

  excludedList.innerHTML = state.excludedColumns.length === 0
    ? `<li class="excluded-empty" style="font-size:0.75rem;color:#cbd5e1;padding:4px 2px;">None</li>`
    : state.excludedColumns.map((col, i) => `
        <li>
          <span class="excluded-name" title="${escHtml(col.header)}">${escHtml(col.header)}</span>
          <button type="button" class="add-back-btn" data-i="${i}">+</button>
        </li>`).join("");

  updateSampleNav();
}

// ── Field operations ──────────────────────────────────────────────────────────
function moveField(i, dir) {
  const j = i + dir;
  if (j < 0 || j >= state.activeColumns.length) return;
  [state.activeColumns[i], state.activeColumns[j]] = [state.activeColumns[j], state.activeColumns[i]];
  renderFields();
}

function removeField(i) {
  const [removed] = state.activeColumns.splice(i, 1);
  state.excludedColumns.push(removed);
  renderFields();
}

function addBackField(i) {
  const [added] = state.excludedColumns.splice(i, 1);
  state.activeColumns.push(added);
  renderFields();
}

// ── Event delegation ──────────────────────────────────────────────────────────
activeList.addEventListener("click", (e) => {
  const btn = e.target.closest("button");
  if (!btn) return;
  const i = parseInt(btn.dataset.i, 10);
  if (btn.classList.contains("order-btn")) {
    moveField(i, parseInt(btn.dataset.dir, 10));
  } else if (btn.classList.contains("remove-field-btn")) {
    removeField(i);
  }
});

activeList.addEventListener("input", (e) => {
  const input = e.target.closest(".field-header-input");
  if (!input) return;
  const i = parseInt(input.closest("tr").dataset.i, 10);
  state.activeColumns[i].header = input.value;
});

activeList.addEventListener("change", (e) => {
  const select = e.target.closest(".field-type-select");
  if (!select) return;
  const i = parseInt(select.closest("tr").dataset.i, 10);
  state.activeColumns[i].type = select.value;
});

excludedList.addEventListener("click", (e) => {
  const btn = e.target.closest(".add-back-btn");
  if (!btn) return;
  addBackField(parseInt(btn.dataset.i, 10));
});

// ── Sample navigation ─────────────────────────────────────────────────────────
samplePrevBtn.addEventListener("click", () => {
  if (state.sampleRowIndex > 0) { state.sampleRowIndex--; refreshSampleCells(); }
});

sampleNextBtn.addEventListener("click", () => {
  if (state.sampleRowIndex < sampleDataRowCount() - 1) { state.sampleRowIndex++; refreshSampleCells(); }
});

// ── Has Headers toggle (delimited only) ───────────────────────────────────────
hasHeadersCheck.addEventListener("change", () => {
  if (state.fileType !== "delimited") return;
  parseColumnsDelimited();
  renderFields();
});

// Qualify headers: re-flatten the stored raw sample without re-reading the file
qualifyHeadersCheck.addEventListener("change", () => {
  if (!["json","xml"].includes(state.fileType) || state.jsonRawSample.length === 0) return;

  const savedTypes      = new Map([...state.activeColumns, ...state.excludedColumns].map(col => [col.originalIndex, col.type]));
  const excludedIndices = new Set(state.excludedColumns.map(col => col.originalIndex));

  state.jsonRecords = state.fileType === "xml"
    ? state.jsonRawSample.map(obj => flattenXmlRecord(obj, qualifyHeadersCheck.checked))
    : state.jsonRawSample.flatMap(obj => flattenObject(obj, "").rows);
  state.excludedColumns = [];
  parseColumnsJson();

  state.activeColumns.forEach(col => {
    if (savedTypes.has(col.originalIndex)) col.type = savedTypes.get(col.originalIndex);
  });
  if (excludedIndices.size > 0) {
    const toExclude       = state.activeColumns.filter(col => excludedIndices.has(col.originalIndex));
    state.activeColumns   = state.activeColumns.filter(col => !excludedIndices.has(col.originalIndex));
    state.excludedColumns = toExclude;
  }

  renderFields();
});

// ── Bulk type setter ──────────────────────────────────────────────────────────
applyTypeBtn.addEventListener("click", () => {
  const type = bulkTypeSelect.value;
  if (!type) return;
  state.activeColumns.forEach(col => { col.type = type; });
  bulkTypeSelect.value = "";
  renderFields();
});

// ── Excel export ──────────────────────────────────────────────────────────────

function coerceValue(raw, type) {
  if (raw === "") return null;
  switch (type) {
    case "Number": {
      const n = Number(raw.replace(/,/g, ""));
      return isNaN(n) ? raw : n;
    }
    case "Currency":
    case "Accounting": {
      const n = Number(raw.replace(/[$€£¥,\s]/g, "").replace(/^\((.+)\)$/, "-$1"));
      return isNaN(n) ? raw : n;
    }
    case "Percentage": {
      const n = parseFloat(raw) / 100;
      return isNaN(n) ? raw : n;
    }
    case "Fraction": {
      const m = raw.trim().match(/^(-?\d+)\s+(\d+)\/(\d+)$|^(-?\d+)\/(\d+)$/);
      if (m) {
        if (m[1] !== undefined) {
          const whole = Number(m[1]), frac = Number(m[2]) / Number(m[3]);
          return whole < 0 ? whole - frac : whole + frac;
        }
        return Number(m[4]) / Number(m[5]);
      }
      return raw;
    }
    case "Scientific": {
      const n = Number(raw);
      return isNaN(n) ? raw : n;
    }
    case "Boolean": {
      const l = raw.toLowerCase();
      if (["true",  "yes", "1", "y"].includes(l)) return true;
      if (["false", "no",  "0", "n"].includes(l)) return false;
      return raw;
    }
    case "Date":
    case "DateTime": {
      const d = new Date(raw);
      return isNaN(d.getTime()) ? raw : d;
    }
    case "Time": {
      const m = raw.trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?(?:\s*(AM|PM))?$/i);
      if (!m) return raw;
      let h = parseInt(m[1], 10), min = parseInt(m[2], 10), sec = m[3] ? parseInt(m[3], 10) : 0;
      if (m[4] && m[4].toUpperCase() === "PM" && h < 12) h += 12;
      if (m[4] && m[4].toUpperCase() === "AM" && h === 12) h = 0;
      return (h * 3600 + min * 60 + sec) / 86400;
    }
    default: return String(raw);
  }
}

// Convert a coerced JS value to an XLSX cell object { t, v }.
function makeCellObject(val, type) {
  if (val === null || val === undefined) return null;
  if (val instanceof Date) {
    // Excel serial date: days since 1899-12-30
    const v = (val - new Date(Date.UTC(1899, 11, 30))) / 86400000;
    return { t: "n", v };
  }
  if (typeof val === "boolean") return { t: "b", v: val };
  if (typeof val === "number") return { t: "n", v: val };
  const str = String(val);
  return { t: "s", v: str.length > 32767 ? str.slice(0, 32767) : str };
}

// Async generator: yield one text line at a time from a File without loading
// the entire file into memory.  Uses the Streams API (supported in all modern
// browsers).  Falls back to readFullFile on browsers that lack file.stream().
async function* streamLines(file) {
  if (!file.stream) {
    // Fallback: load whole file (older browsers)
    const text = await readFullFile(file);
    for (const line of text.split(/\r?\n/)) yield line;
    return;
  }

  const decoder = new TextDecoder();
  const reader  = file.stream().getReader();
  let   buffer  = "";

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split(/\r?\n/);
      buffer = lines.pop(); // keep the incomplete trailing fragment
      for (const line of lines) yield line;
    }
    // Flush decoder and yield any remaining content
    buffer += decoder.decode();
    if (buffer.length > 0) yield buffer;
  } finally {
    reader.releaseLock();
  }
}

// ── Streaming XLSX writer ─────────────────────────────────────────────────────
// Generates XLSX directly via fflate's streaming ZIP - row XML is built in
// ~512 KB batches and fed to the compressor so no large string accumulates.
// This sidesteps V8's "Invalid string length" error that SheetJS.write hits
// when the worksheet XML exceeds ~500 MB for large row counts.

function _colLetter(n) {
  let s = "";
  for (n++; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(64 + (n - 1) % 26 + 1) + s;
  return s;
}

function _cellXml(addr, cell, sAttr) {
  if (cell.t === "n") return `<c r="${addr}"${sAttr}><v>${cell.v}</v></c>`;
  if (cell.t === "b") return `<c r="${addr}" t="b"${sAttr}><v>${cell.v ? 1 : 0}</v></c>`;
  const txt = String(cell.v).replace(/[&<>]/g, ch => ch === "&" ? "&amp;" : ch === "<" ? "&lt;" : "&gt;");
  return `<c r="${addr}" t="inlineStr"${sAttr}><is><t>${txt}</t></is></c>`;
}

function _rowXml(rowNum, cells, letters, styles) {
  let xml = `<row r="${rowNum}">`;
  for (let c = 0; c < letters.length; c++) {
    const cell = cells[c];
    if (cell == null) continue;
    const si = styles ? styles[c] : 0;
    xml += _cellXml(letters[c] + rowNum, cell, si ? ` s="${si}"` : "");
  }
  return xml + "</row>";
}

// Matches cellXfs order in _STYLES below: [0]General [1]Date [2]DateTime [3]Number [4]Text [5]Currency [6]Accounting [7]Percentage [8]Fraction [9]Scientific [10]Time
const _STYLE_IDX = { Date: 1, DateTime: 2, Number: 3, Text: 4, Currency: 5, Accounting: 6, Percentage: 7, Fraction: 8, Scientific: 9, Time: 10 };

const _CT       = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;
const _RELS     = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
const _WB       = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`;
const _WB_RELS  = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
const _STYLES   = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="9"><numFmt numFmtId="164" formatCode="MM/DD/YYYY"/><numFmt numFmtId="165" formatCode="MM/DD/YYYY HH:MM:SS"/><numFmt numFmtId="166" formatCode="0.##########"/><numFmt numFmtId="167" formatCode="$#,##0.00"/><numFmt numFmtId="168" formatCode="_(&quot;$&quot;* #,##0.00_);_(&quot;$&quot;* (#,##0.00);_(&quot;$&quot;* &quot;-&quot;??_);_(@_)"/><numFmt numFmtId="169" formatCode="0.00%"/><numFmt numFmtId="170" formatCode="# ?/?"/><numFmt numFmtId="171" formatCode="0.00E+00"/><numFmt numFmtId="172" formatCode="h:mm:ss AM/PM"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="11"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="168" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="169" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="170" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="171" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="172" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

async function writeStreamXLSX(cols, rowsIter, filename) {
  const enc     = new TextEncoder();
  const letters = Array.from({ length: cols.length }, (_, i) => _colLetter(i));
  const styles  = cols.map(col => _STYLE_IDX[col.type] ?? 0);

  const chunks = [];
  let zipResolve, zipReject;
  const zipDone = new Promise((res, rej) => { zipResolve = res; zipReject = rej; });

  const zip = new fflate.Zip((err, data, final) => {
    if (err) { zipReject(err); return; }
    chunks.push(data.slice());
    if (final) zipResolve();
  });

  const addFile = (name, text) => {
    const f = new fflate.ZipDeflate(name, { level: 1 });
    zip.add(f);
    f.push(enc.encode(text), true);
  };
  addFile("[Content_Types].xml",        _CT);
  addFile("_rels/.rels",                _RELS);
  addFile("xl/workbook.xml",            _WB);
  addFile("xl/_rels/workbook.xml.rels", _WB_RELS);
  addFile("xl/styles.xml",              _STYLES);

  const sheet = new fflate.ZipDeflate("xl/worksheets/sheet1.xml", { level: 1 });
  zip.add(sheet);

  const push = (xml, final = false) => sheet.push(enc.encode(xml), final);
  push(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>`);
  push(_rowXml(1, cols.map(c => ({ t: "s", v: c.header })), letters, styles));

  let buf    = "";
  let rowNum = 2;
  let count  = 0;
  for await (const row of rowsIter) {
    buf += _rowXml(rowNum++, row, letters, styles);
    count++;
    if (buf.length >= 524288) { push(buf); buf = ""; }
    if (count % 25000 === 0) {
      setExportProgress(`Exporting… ${count.toLocaleString()} rows`);
      await new Promise(r => setTimeout(r, 0));
    }
  }
  push(buf + "</sheetData></worksheet>", true);

  zip.end();
  await zipDone;

  const blob = new Blob(chunks, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url  = URL.createObjectURL(blob);
  Object.assign(document.createElement("a"), { href: url, download: filename }).click();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

// ── Export functions ───────────────────────────────────────────────────────────

async function exportDelimited() {
  const cols      = state.activeColumns;
  const skipFirst = hasHeadersCheck.checked;
  let   truncated = false;

  async function* rows() {
    let isFirst = true;
    let count   = 0;
    for await (const line of streamLines(state.currentFile)) {
      if (!line.trim()) continue;
      if (isFirst && skipFirst) { isFirst = false; continue; }
      isFirst = false;
      if (count >= MAX_EXPORT_ROWS) { truncated = true; return; }
      const cells = line.split(state.delimChar);
      const row   = [];
      cols.forEach((col, c) => {
        const cell = makeCellObject(coerceValue((cells[col.originalIndex] ?? "").trim(), col.type), col.type);
        if (cell) row[c] = cell;
      });
      yield row;
      count++;
    }
  }

  await writeStreamXLSX(cols, rows(), state.currentFile.name.replace(/\.[^.]+$/, "") + ".xlsx");
  if (truncated) alert(`File exceeds Excel's ${MAX_EXPORT_ROWS.toLocaleString()}-row limit. Export contains the first ${MAX_EXPORT_ROWS.toLocaleString()} rows.`);
}

async function exportStructured() {
  const cols      = state.activeColumns;
  let   truncated = false;

  async function* rows() {
    let count = 0;
    for await (const rawObj of streamStructuredObjects(state.currentFile)) {
      for (const flatRow of flattenRecord(rawObj)) {
        if (count >= MAX_EXPORT_ROWS) { truncated = true; return; }
        const row = [];
        cols.forEach((col, c) => {
          const cell = makeCellObject(coerceValue(String(flatRow[col.key] ?? "").trim(), col.type), col.type);
          if (cell) row[c] = cell;
        });
        yield row;
        count++;
      }
      if (truncated) return;
    }
  }

  await writeStreamXLSX(cols, rows(), state.currentFile.name.replace(/\.[^.]+$/, "") + ".xlsx");
  if (truncated) alert(`Export truncated at ${MAX_EXPORT_ROWS.toLocaleString()} rows (Excel maximum).`);
}

// ── Busy state helpers ────────────────────────────────────────────────────────
function setBusy(label) {
  viewBtn.disabled    = true;
  exportBtn.disabled  = true;
  exportBtn.innerHTML = `${escHtml(label)}`;
}

function clearBusy() {
  viewBtn.disabled    = false;
  exportBtn.disabled  = false;
  exportBtn.innerHTML = `Export`;
}

function setExportProgress(msg) {
  exportBtn.innerHTML = `${escHtml(msg)}`;
}

// ── Export dropdown toggle ────────────────────────────────────────────────────
exportBtn.addEventListener("click", () => exportMenu.classList.toggle("open"));
document.addEventListener("click", (e) => {
  if (!e.target.closest(".export-wrapper")) exportMenu.classList.remove("open");
});

// ── Excel export ──────────────────────────────────────────────────────────────
exportExcelBtn.addEventListener("click", async () => {
  exportMenu.classList.remove("open");
  if (!state.currentFile) {
    if (state.activeColumns.length > 0) alert("Use the Browse button in the File section to link your data file first.");
    return;
  }
  setBusy("Exporting…");
  try {
    if (state.fileType === "json" || state.fileType === "xml") await exportStructured();
    else                                                       await exportDelimited();
  } catch (err) {
    alert(`Export failed: ${err.message}`);
  } finally {
    clearBusy();
  }
});

// ── CSV export ────────────────────────────────────────────────────────────────
function csvCell(val) {
  const s = String(val ?? "");
  return /[,"\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

exportCsvBtn.addEventListener("click", async () => {
  exportMenu.classList.remove("open");
  if (!state.currentFile) {
    if (state.activeColumns.length > 0) alert("Use the Browse button in the File section to link your data file first.");
    return;
  }
  setBusy("Saving CSV…");

  try {
    const cols   = state.activeColumns;
    const chunks = ["﻿"]; // UTF-8 BOM (Excel opens accented chars correctly)
    chunks.push(cols.map(c => csvCell(c.header)).join(",") + "\r\n");

    let count = 0;

    if (state.fileType === "json" || state.fileType === "xml") {
      for await (const rawObj of streamStructuredObjects(state.currentFile)) {
        for (const flat of flattenRecord(rawObj)) {
          chunks.push(cols.map(c => csvCell(flat[c.key] ?? "")).join(",") + "\r\n");
          if (++count % 10000 === 0) {
            setExportProgress(`Saving CSV… ${count.toLocaleString()} rows`);
            await new Promise(r => setTimeout(r, 0));
          }
        }
      }
    } else {
      let isFirst = true;
      for await (const line of streamLines(state.currentFile)) {
        if (!line.trim()) continue;
        if (isFirst && hasHeadersCheck.checked) { isFirst = false; continue; }
        isFirst = false;
        const cells = line.split(state.delimChar);
        chunks.push(cols.map(c => csvCell((cells[c.originalIndex] ?? "").trim())).join(",") + "\r\n");
        if (++count % 25000 === 0) {
          setExportProgress(`Saving CSV… ${count.toLocaleString()} rows`);
          await new Promise(r => setTimeout(r, 0));
        }
      }
    }

    const blob = new Blob(chunks, { type: "text/csv;charset=utf-8" });
    const a    = Object.assign(document.createElement("a"), {
      href:     URL.createObjectURL(blob),
      download: state.currentFile.name.replace(/\.[^.]+$/, "") + ".csv",
    });
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 60_000);
  } catch (err) {
    alert(`CSV export failed: ${err.message}`);
  } finally {
    clearBusy();
  }
});

// ── View ──────────────────────────────────────────────────────────────────────
function buildViewPage(filename, headers, rows, truncated) {
  // Safely embed data so </script> inside values can't break the HTML parser
  const hJson = JSON.stringify(headers).replace(/<\//g, "<\\/");
  const rJson = JSON.stringify(rows).replace(/<\//g, "<\\/");
  const info  = truncated
    ? `first ${rows.length.toLocaleString()} rows (preview limit)`
    : `${rows.length.toLocaleString()} rows`;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Preview: ${escHtml(filename)}</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#f1f5f9;min-height:100vh}
header{background:#1e293b;color:#fff;padding:13px 24px;display:flex;align-items:center;justify-content:space-between;gap:16px;position:sticky;top:0;z-index:10}
.htitle{font-size:.9375rem;font-weight:700;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.hmeta{font-size:.8125rem;color:#94a3b8;white-space:nowrap;flex-shrink:0}
.wrap{padding:18px 24px}
.pager{display:flex;align-items:center;gap:10px;margin-bottom:14px}
.pbtn{padding:6px 14px;background:#fff;border:1px solid #e2e8f0;border-radius:6px;font-size:.8125rem;font-weight:500;color:#475569;cursor:pointer;transition:background .12s}
.pbtn:hover:not(:disabled){background:#f0f4ff;border-color:#a5b4fc;color:#4f46e5}
.pbtn:disabled{opacity:.35;cursor:not-allowed}
.pinfo{font-size:.8125rem;color:#64748b;flex:1}
.tscroll{background:#fff;border-radius:12px;overflow:auto;box-shadow:0 1px 3px rgba(0,0,0,.06),0 4px 16px rgba(0,0,0,.05)}
table{width:100%;border-collapse:collapse;font-size:.8125rem}
th{position:sticky;top:0;background:#f8fafc;padding:9px 14px;text-align:left;font-size:.6875rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#94a3b8;border-bottom:1px solid #e2e8f0;white-space:nowrap}
td{padding:9px 14px;border-top:1px solid #f1f5f9;color:#334155;white-space:nowrap;max-width:280px;overflow:hidden;text-overflow:ellipsis}
tr:hover td{background:#fafbff}
.notice{margin-top:12px;padding:10px 14px;background:#fef3c7;border:1px solid #fde68a;border-radius:8px;font-size:.8125rem;color:#92400e}
</style>
</head>
<body>
<header>
  <span class="htitle">${escHtml(filename)}</span>
  <span class="hmeta" id="hmeta"></span>
</header>
<div class="wrap">
  <div class="pager">
    <button class="pbtn" id="prev" disabled>&#8592; Previous</button>
    <span class="pinfo" id="pinfo"></span>
    <button class="pbtn" id="next">Next &#8594;</button>
  </div>
  <div class="tscroll">
    <table><thead id="thead"></thead><tbody id="tbody"></tbody></table>
  </div>
  ${truncated ? `<p class="notice">Preview limited to the ${info}. Export the full file to access all data.</p>` : ""}
</div>
<script>
const H=${hJson},R=${rJson},P=100;let p=0;
function e(s){return String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function render(){
  const tp=Math.max(1,Math.ceil(R.length/P));
  p=Math.max(0,Math.min(p,tp-1));
  const s=p*P,en=Math.min(s+P,R.length);
  document.getElementById("thead").innerHTML="<tr>"+H.map(h=>"<th>"+e(h)+"</th>").join("")+"</tr>";
  document.getElementById("tbody").innerHTML=R.slice(s,en).map(r=>"<tr>"+r.map(c=>{const v=e(String(c??""));return\`<td title="\${v}">\${v}</td>\`;}).join("")+"</tr>").join("");
  document.getElementById("pinfo").textContent="Page "+(p+1)+" of "+tp+" \xb7 Rows "+(s+1).toLocaleString()+"–"+en.toLocaleString()+" of "+R.length.toLocaleString();
  document.getElementById("hmeta").textContent=R.length.toLocaleString()+" rows \xb7 "+H.length+" columns";
  document.getElementById("prev").disabled=p===0;
  document.getElementById("next").disabled=en>=R.length;
  window.scrollTo(0,0);
}
document.getElementById("prev").addEventListener("click",()=>{p--;render();});
document.getElementById("next").addEventListener("click",()=>{p++;render();});
render();
<\/script>
</body></html>`;
}

viewBtn.addEventListener("click", async () => {
  if (!state.currentFile) {
    if (state.activeColumns.length > 0) alert("Use the Browse button in the File section to link your data file first.");
    return;
  }

  // Open immediately (synchronous user gesture) so popup blocker doesn't fire
  const win = window.open("", "_blank");
  if (!win) { alert("Please allow popups for this site to use View."); return; }
  win.document.write("<html><body style='font-family:sans-serif;padding:40px;color:#64748b'>Loading data…</body></html>");

  setBusy("Loading…");
  viewBtn.textContent = "Loading…";

  try {
    const cols = state.activeColumns;
    const rows = [];

    if (state.fileType === "json" || state.fileType === "xml") {
      outer: for await (const rawObj of streamStructuredObjects(state.currentFile)) {
        for (const flat of flattenRecord(rawObj)) {
          rows.push(cols.map(c => flat[c.key] ?? null));
          if (rows.length >= MAX_VIEW_ROWS) break outer;
        }
      }
    } else {
      let isFirst = true;
      for await (const line of streamLines(state.currentFile)) {
        if (!line.trim()) continue;
        if (isFirst && hasHeadersCheck.checked) { isFirst = false; continue; }
        isFirst = false;
        const cells = line.split(state.delimChar);
        rows.push(cols.map(c => (cells[c.originalIndex] ?? "").trim()));
        if (rows.length >= MAX_VIEW_ROWS) break;
      }
    }

    const html = buildViewPage(
      state.currentFile.name,
      cols.map(c => c.header),
      rows,
      rows.length >= MAX_VIEW_ROWS
    );
    win.document.open();
    win.document.write(html);
    win.document.close();
  } catch (err) {
    win.close();
    alert(`View failed: ${err.message}`);
  } finally {
    viewBtn.textContent = "Preview";
    clearBusy();
  }
});

// ── UI helpers ────────────────────────────────────────────────────────────────
function setDelimitedMode() {
  delimiterFieldEl.classList.remove("is-disabled");
  delimiterInput.disabled         = false;
  hasHeadersLabelEl.classList.remove("is-disabled");
  hasHeadersCheck.disabled        = false;
  qualifyHeadersLabelEl.classList.add("is-disabled");
  qualifyHeadersCheck.disabled    = true;
}

function setJsonMode() {
  delimiterFieldEl.classList.add("is-disabled");
  delimiterInput.disabled         = true;
  delimiterInput.value            = "";
  hasHeadersLabelEl.classList.add("is-disabled");
  hasHeadersCheck.disabled        = true;
  qualifyHeadersLabelEl.classList.remove("is-disabled");
  qualifyHeadersCheck.disabled    = false;
}

// ── File row rendering ────────────────────────────────────────────────────────
function renderFileRow(file) {
  const type = ext(file.name);
  tableBody.innerHTML = `
    <tr>
      <td class="file-name">${escHtml(file.name)}</td>
      <td class="size">${formatSize(file.size)}</td>
      <td><span class="type-badge type-${type}">${type}</span></td>
      <td class="date">${formatDate(file.lastModified) ?? "-"}</td>
      <td><button type="button" class="update-file-btn">Browse</button></td>
    </tr>`;
}

// ── Main file processing ──────────────────────────────────────────────────────
async function processFile(file) {
  clearError();
  const type = ext(file.name);

  if (!ALLOWED.includes(type)) {
    showError(`"${file.name}" is not supported. Please select a CSV, TXT, JSON, or XML file.`);
    return;
  }

  // ── SXL ──
  if (type === "sxl") {
    await loadSxl(file);
    return;
  }

  dropZone.style.display = "none";
  state.currentFile      = file;

  renderFileRow(file);
  tableSection.style.display = "block";

  // Reset config state
  state.fileType        = null;
  state.delimChar       = null;
  state.sxlFileMeta     = null;
  state.rawRows         = [];
  state.jsonRawSample   = [];
  state.jsonRecords     = [];
  state.activeColumns   = [];
  state.excludedColumns = [];
  state.sampleRowIndex  = 0;
  configSection.style.display = "none";
  sxlNotice.style.display     = "none";
  card.classList.remove("card--wide");

  // ── XML ──
  if (type === "xml") {
    setJsonMode();
    exportBtn.title = "";
    const MAX_CONFIG_RECORDS = 200;
    const rawSample  = [];
    const sampleFlat = [];
    for await (const rawObj of streamXmlObjects(file)) {
      rawSample.push(rawObj);
      sampleFlat.push(flattenXmlRecord(rawObj, qualifyHeadersCheck?.checked ?? false));
      if (sampleFlat.length >= MAX_CONFIG_RECORDS) break;
    }
    if (sampleFlat.length === 0) {
      showError(`Could not parse "${file.name}". No repeating records could be detected in this XML.`);
      dropZone.style.display = "";
      return;
    }
    state.fileType      = "xml";
    state.jsonRawSample = rawSample;
    state.jsonRecords   = sampleFlat;
    parseColumnsJson();
    renderFields();
    configSection.style.display = "block";
    card.classList.add("card--wide");
    return;
  }

  // ── JSON ──
  if (type === "json") {
    setJsonMode();
    exportBtn.title = "";

    const MAX_CONFIG_RECORDS = 200;
    const rawSample  = [];
    const sampleFlat = [];
    for await (const rawObj of streamJsonObjects(file)) {
      rawSample.push(rawObj);
      for (const row of flattenObject(rawObj, "").rows) sampleFlat.push(row);
      if (sampleFlat.length >= MAX_CONFIG_RECORDS) break;
    }

    if (sampleFlat.length === 0) {
      showError(`Could not parse "${file.name}". No tabular data could be found in this JSON structure.`);
      dropZone.style.display = "";
      return;
    }

    state.fileType      = "json";
    state.jsonRawSample = rawSample;
    state.jsonRecords   = sampleFlat;
    parseColumnsJson();
    renderFields();
    configSection.style.display = "block";
    card.classList.add("card--wide");
    return;
  }

  // ── Delimited (CSV / TXT) ──
  setDelimitedMode();

  if (type === "csv") {
    state.delimChar = ",";
  } else if (type === "txt") {
    const sample    = await readSample(file);
    state.delimChar = detectDelimiterChar(sample);
  }

  if (state.delimChar) {
    const sample  = await readSample(file);
    state.rawRows = parseSampleRows(sample, state.delimChar);
    delimiterInput.value = delimCharToLabel(state.delimChar);
    state.fileType = "delimited";
    parseColumnsDelimited();
    renderFields();
    configSection.style.display = "block";
    card.classList.add("card--wide");
  }else{
    showError(`Could not parse "${file.name}". No delimiter could be found in this Text file.`);
    dropZone.style.display = "";
    return;
  }
}

// ── Reset ─────────────────────────────────────────────────────────────────────
function reset() {
  dropZone.style.display      = "";
  tableSection.style.display  = "none";
  configSection.style.display = "none";
  sxlNotice.style.display     = "none";
  card.classList.remove("card--wide");
  clearError();
  fileInput.value             = "";
  updateFileInput.value       = "";
  hasHeadersCheck.checked     = true;
  qualifyHeadersCheck.checked = false;
  bulkTypeSelect.value        = "";
  setDelimitedMode();
  state.currentFile         = null;
  state.sxlFileMeta         = null;
  state.fileType            = null;
  state.delimChar           = null;
  state.rawRows             = [];
  state.jsonRawSample       = [];
  state.jsonRecords         = [];
  state.activeColumns       = [];
  state.excludedColumns     = [];
  state.sampleRowIndex      = 0;
}

// ── SXL save / load ───────────────────────────────────────────────────────────

function saveSxl() {
  const meta = state.currentFile
    ? { name: state.currentFile.name, size: state.currentFile.size, lastModified: state.currentFile.lastModified }
    : (state.sxlFileMeta ?? { name: "schema", size: 0, lastModified: 0 });

  const payload = {
    sxlVersion:     1,
    file:           meta,
    fileType:       state.fileType,
    delimChar:      state.delimChar,
    hasHeaders:     hasHeadersCheck.checked,
    qualifyHeaders: qualifyHeadersCheck.checked,
    activeColumns:  state.activeColumns.map(c => ({ ...c })),
    excludedColumns: state.excludedColumns.map(c => ({ ...c })),
  };

  const blob     = new Blob([JSON.stringify(payload, null, 2)], { type: "application/octet-stream" });
  const filename = meta.name.replace(/\.[^.]+$/, "") + ".sxl";
  const a        = Object.assign(document.createElement("a"), { href: URL.createObjectURL(blob), download: filename });
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 60_000);
}

async function loadSxl(file) {
  const text = await readFullFile(file);
  let data;
  try { data = JSON.parse(text); } catch {
    showError(`Could not read "${file.name}". The file may be corrupted.`);
    return;
  }
  if (!data.sxlVersion || !Array.isArray(data.activeColumns)) {
    showError(`"${file.name}" is not a valid SchemaXL file.`);
    return;
  }

  state.currentFile     = null;
  state.sxlFileMeta     = data.file ?? null;
  state.fileType        = data.fileType;
  state.delimChar       = data.delimChar ?? null;
  state.rawRows         = [];
  state.jsonRawSample   = [];
  state.jsonRecords     = [];
  state.activeColumns   = data.activeColumns;
  state.excludedColumns = data.excludedColumns ?? [];
  state.sampleRowIndex  = 0;

  hasHeadersCheck.checked     = data.hasHeaders ?? true;
  qualifyHeadersCheck.checked = data.qualifyHeaders ?? false;
  bulkTypeSelect.value        = "";

  if (data.fileType === "delimited") {
    delimiterInput.value = delimCharToLabel(data.delimChar ?? ",");
    setDelimitedMode();
  } else {
    delimiterInput.value = "";
    setJsonMode();
  }

  const meta      = data.file ?? {};
  const savedExt  = (meta.name ?? "").split(".").pop().toLowerCase();
  const badgeType = DATA_ALLOWED.includes(savedExt) ? savedExt : "csv";
  dropZone.style.display = "none";
  tableBody.innerHTML = `
    <tr>
      <td class="file-name">${escHtml(meta.name ?? "(unknown)")}</td>
      <td class="size">${meta.size ? formatSize(meta.size) : "—"}</td>
      <td><span class="type-badge type-${badgeType}">${badgeType}</span></td>
      <td class="date">${formatDate(meta.lastModified) ?? "—"}</td>
      <td><button type="button" class="update-file-btn">Browse</button></td>
    </tr>`;
  tableSection.style.display  = "block";
  sxlNotice.style.display     = "block";

  card.classList.add("card--wide");
  renderFields();
  configSection.style.display = "block";
}

saveBtn.addEventListener("click", saveSxl);

// ── Drop zone events ──────────────────────────────────────────────────────────
fileInput.addEventListener("change", (e) => {
  if (e.target.files[0]) processFile(e.target.files[0]);
});

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("drag-over");
});

dropZone.addEventListener("dragleave", (e) => {
  if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove("drag-over");
});

dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("drag-over");
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
});

resetBtn.addEventListener("click", reset);

// ── Update file ───────────────────────────────────────────────────────────────

async function handleUpdateFile(file) {
  if (!DATA_ALLOWED.includes(ext(file.name))) {
    alert(`"${file.name}" is not supported. Please select a CSV, TXT, JSON, or XML file.`);
    return false;
  }

  const fileExt  = ext(file.name);
  const expected = state.activeColumns.length + state.excludedColumns.length;

  if (state.fileType === "delimited") {
    const sample    = await readSample(file);
    const delimChar = fileExt === "csv" ? "," : detectDelimiterChar(sample);
    if (!delimChar) {
      alert(`Could not detect a delimiter in "${file.name}".`);
      return false;
    }
    const rows   = parseSampleRows(sample, delimChar);
    const actual = rows.length > 0 ? rows[0].length : 0;
    if (actual !== expected) {
      alert(`Column count mismatch: "${file.name}" has ${actual} column${actual !== 1 ? "s" : ""} but the schema expects ${expected}. Select a file with the same structure.`);
      return false;
    }
    state.currentFile    = file;
    state.rawRows        = rows;
    state.delimChar      = delimChar;
    delimiterInput.value = delimCharToLabel(delimChar);
  } else {
    const MAX_SAMPLE = 200;
    const qualify    = qualifyHeadersCheck?.checked ?? false;
    const rawSample  = [];
    const sampleFlat = [];
    if (fileExt === "json") {
      for await (const rawObj of streamJsonObjects(file)) {
        rawSample.push(rawObj);
        for (const row of flattenObject(rawObj, "").rows) sampleFlat.push(row);
        if (sampleFlat.length >= MAX_SAMPLE) break;
      }
    } else {
      for await (const rawObj of streamXmlObjects(file)) {
        rawSample.push(rawObj);
        sampleFlat.push(flattenXmlRecord(rawObj, qualify));
        if (sampleFlat.length >= MAX_SAMPLE) break;
      }
    }
    if (sampleFlat.length === 0) {
      alert(`Could not parse "${file.name}".`);
      return false;
    }
    const keys   = new Set();
    sampleFlat.forEach(row => Object.keys(row).forEach(k => keys.add(k)));
    const actual = keys.size;
    if (actual !== expected) {
      alert(`Column count mismatch: "${file.name}" has ${actual} column${actual !== 1 ? "s" : ""} but the schema expects ${expected}. Select a file with the same structure.`);
      return false;
    }
    state.currentFile   = file;
    state.jsonRawSample = rawSample;
    state.jsonRecords   = sampleFlat;
  }

  state.sampleRowIndex        = 0;
  state.sxlFileMeta           = null;
  sxlNotice.style.display     = "none";
  renderFileRow(file);
  refreshSampleCells();
  return true;
}

tableBody.addEventListener("click", (e) => {
  if (e.target.closest(".update-file-btn")) updateFileInput.click();
});

updateFileInput.addEventListener("change", async (e) => {
  const file = e.target.files[0];
  updateFileInput.value = "";
  if (!file) return;
  const btn = tableBody.querySelector(".update-file-btn");
  if (btn) { btn.disabled = true; btn.textContent = "Checking…"; }
  const ok = await handleUpdateFile(file);
  if (!ok) {
    const b = tableBody.querySelector(".update-file-btn");
    if (b) { b.disabled = false; b.textContent = "Browse"; }
  }
});
