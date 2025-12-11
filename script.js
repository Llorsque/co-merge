// Vereniging Matcher v2
// Let op: vereist internettoegang om de XLSX-bibliotheek vanaf cdnjs te laden.

const logEl = document.getElementById("log");
function log(msg) {
  if (!logEl) return;
  const time = new Date().toLocaleTimeString();
  logEl.textContent += `[${time}] ${msg}\n`;
  logEl.scrollTop = logEl.scrollHeight;
}

// Runtime check: bestaat XLSX?
if (typeof XLSX === "undefined") {
  if (logEl) {
    logEl.textContent = "";
    log("❌ De Excel-bibliotheek (XLSX) is niet geladen.");
    log("   Mogelijke oorzaken:");
    log("   - Geen internetverbinding of CDN geblokkeerd");
    log("   - Script-tag naar xlsx.full.min.js ontbreekt of is aangepast");
    log("   Oplossing:");
    log("   - Zorg dat je internet hebt en vernieuw de pagina");
    log("   - Of host de tool op GitHub Pages zodat de CDN goed laad.");
  } else {
    alert("De Excel-bibliotheek (XLSX) is niet geladen. Controleer je internetverbinding of de script-tag in index.html.");
  }
}

// ====== Alleen verder gaan als XLSX bestaat ======
if (typeof XLSX !== "undefined") {

// ====== Normalisatie helpers ======
function normalizeKvk(value) {
  if (value === undefined || value === null) return "";
  let s = String(value).trim();
  if (!s) return "";
  s = s.replace(/\D+/g, "");
  if (!s) return "";
  return s.padStart(8, "0");
}

function normalizeName(value) {
  if (value === undefined || value === null) return "";
  let s = String(value).toLowerCase().trim();
  s = s.replace(/["“”„'’`]/g, "");
  s = s.replace(/\s+/g, " ");
  return s;
}

function normalizeGemeente(value) {
  if (value === undefined || value === null) return "";
  let s = String(value).toLowerCase().trim();
  s = s.normalize("NFD").replace(/\p{Diacritic}/gu, "");
  s = s.replace(/[^a-z]/g, "");
  return s;
}

function bigrams(str) {
  const s = str;
  const res = [];
  for (let i = 0; i < s.length - 1; i++) {
    res.push(s.slice(i, i + 2));
  }
  return res;
}

function bigramSimilarity(a, b) {
  a = a || "";
  b = b || "";
  if (!a && !b) return 100;
  if (!a || !b) return 0;
  const aGrams = bigrams(a);
  const bGrams = bigrams(b);
  if (!aGrams.length || !bGrams.length) {
    return a === b ? 100 : 0;
  }
  const aMap = new Map();
  for (const g of aGrams) aMap.set(g, (aMap.get(g) || 0) + 1);
  let inter = 0;
  const bMap = new Map();
  for (const g of bGrams) bMap.set(g, (bMap.get(g) || 0) + 1);
  for (const [g, aCount] of aMap.entries()) {
    const bCount = bMap.get(g) || 0;
    inter += Math.min(aCount, bCount);
  }
  const score = (2 * inter) / (aGrams.length + bGrams.length);
  return Math.round(score * 100);
}

function naamStatus(naamCO, naamAan) {
  const e1 = naamCO == null ? "" : String(naamCO);
  const e2 = naamAan == null ? "" : String(naamAan);
  if (e1 && e1 === e2) return "Exact";
  const clean = (s) => {
    if (s == null) return "";
    let t = String(s).toLowerCase().trim();
    t = t.replace(/["“”„'’`]/g, "");
    t = t.replace(/\s+/g, " ");
    return t;
  };
  const a1 = clean(naamCO);
  const a2 = clean(naamAan);
  if (a1 && a1 === a2) return "Bijna";
  return "Anders";
}

function sheetToJson(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) throw new Error(`Kon werkblad '${sheetName}' niet vinden.`);
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

function firstSheetName(workbook) {
  return workbook.SheetNames[0];
}

function buildResult(aanRows, coRows) {
  log("Start met normaliseren van KVK-nummers...");

  const aan = aanRows.map((r) => ({
    ...r,
    KVK8: normalizeKvk(r.DOSSIERNR ?? r["DOSSIERNR"]),
    FANAAM: r.FANAAM ?? r["FANAAM"] ?? "",
    VENNAAM: r.VENNAAM ?? r["VENNAAM"] ?? "",
    RECHTSVORM: r.RECHTSVORM ?? r["RECHTSVORM"] ?? "",
    GEMEENTE: r.GEMEENTE ?? r["GEMEENTE"] ?? "",
    WOONPLAATS: r.WOONPLAATS ?? r["WOONPLAATS"] ?? "",
    POSTCODE: r.POSTCODE ?? r["POSTCODE"] ?? "",
    STRAAT: r.STRAAT ?? r["STRAAT"] ?? "",
    HUISNR: r.HUISNR ?? r["HUISNR"] ?? "",
    HUISNRTOEV: r.HUISNRTOEV ?? r["HUISNRTOEV"] ?? "",
    TELNR: r.TELNR ?? r["TELNR"] ?? "",
    EMAIL: r.EMAIL ?? r["EMAIL"] ?? "",
  }));

  const co = coRows.map((r) => ({
    ...r,
    Nr: r["Nr."] ?? r["Nr"] ?? "",
    KvKnummer: r["KvK-nummer"] ?? r["KvK nummer"] ?? r["KVK-nummer"] ?? "",
    Naam: r["Naam"] ?? "",
    Subsoort: r["Subsoort organisatie"] ?? "",
    Gemeente: r["Vestigingsgemeente"] ?? "",
    Telefoon: r["Telefoonnr."] ?? r["Telefoonnr"] ?? "",
    Email: r["E-mail"] ?? r["Email"] ?? "",
    Postadres: r["Postadres"] ?? "",
    Geblokkeerd: r["Geblokkeerd"] ?? "",
    KVK8: normalizeKvk(
      r["KvK-nummer"] ?? r["KvK nummer"] ?? r["KVK-nummer"] ?? ""
    ),
  }));

  const aanByKvk = {};
  for (const row of aan) {
    if (row.KVK8 && !aanByKvk[row.KVK8]) {
      aanByKvk[row.KVK8] = row;
    }
  }

  const tab1 = [];
  const tab2 = [];
  const tab3 = [];

  for (const r of co) {
    if (!r.KVK8) {
      tab2.push(r);
    } else if (aanByKvk[r.KVK8]) {
      const a = aanByKvk[r.KVK8];
      const row = {
        Nr: r.Nr,
        "KvK-nummer": r.KvKnummer,
        KVK8: r.KVK8,
        DOSSIERNR: a.DOSSIERNR,
        Naam_CO: r.Naam,
        Naam_AAN_officieel: a.FANAAM,
        Naam_AAN_handelsnaam: a.VENNAAM,
        Subsoort_CO: r.Subsoort,
        Rechtsvorm_AAN: a.RECHTSVORM,
        Gemeente_CO: r.Gemeente,
        Gemeente_AAN: a.GEMEENTE,
        Woonplaats_AAN: a.WOONPLAATS,
        Email_CO: r.Email,
        Email_AAN: a.EMAIL,
        Telefoon_CO: r.Telefoon,
        Telefoon_AAN: a.TELNR,
        Postadres_CO: r.Postadres,
        Geblokkeerd_CO: r.Geblokkeerd,
      };
      row.NaamStatus = naamStatus(row.Naam_CO, row.Naam_AAN_officieel);
      tab1.push(row);
    } else {
      tab3.push(r);
    }
  }

  log(`KVK-match (Tabblad 1): ${tab1.length} rijen`);
  log(`Geen KVK in CO (Tabblad 2): ${tab2.length} rijen`);
  log(`Wel KVK, geen match in Aanbieders (Tabblad 3): ${tab3.length} rijen`);

  log("Start met zoeken naar naamsuggesties voor Tabblad 3...");

  const aanCandidates = [];
  for (const a of aan) {
    const n1 = normalizeName(a.FANAAM);
    const n2 = normalizeName(a.VENNAAM);
    if (n1) aanCandidates.push({ key: n1, row: a, type: "FANAAM" });
    if (n2) aanCandidates.push({ key: n2, row: a, type: "VENNAAM" });
  }

  const tab4 = [];
  const threshold = 90;

  for (const coRow of tab3) {
    const coNameNorm = normalizeName(coRow.Naam);
    if (!coNameNorm) continue;

    let best = null;
    for (const cand of aanCandidates) {
      const score = bigramSimilarity(coNameNorm, cand.key);
      if (!best || score > best.score) {
        best = { score, cand };
      }
    }
    if (!best || best.score < threshold) continue;

    const a = best.cand.row;
    const gemCoNorm = normalizeGemeente(coRow.Gemeente);
    const gemAanNorm = normalizeGemeente(a.GEMEENTE);
    const gemeenteMatcht = gemCoNorm && gemCoNorm === gemAanNorm;

    tab4.push({
      Nr: coRow.Nr,
      "KvK-nummer": coRow.KvKnummer,
      KVK8: coRow.KVK8,
      Naam_CO: coRow.Naam,
      Gemeente_CO: coRow.Gemeente,
      Email_CO: coRow.Email,
      Subsoort_CO: coRow.Subsoort,
      Telefoon_CO: coRow.Telefoon,
      Postadres_CO: coRow.Postadres,

      DOSSIERNR: a.DOSSIERNR,
      Naam_AAN_officieel: a.FANAAM,
      Naam_AAN_handelsnaam: a.VENNAAM,
      Rechtsvorm_AAN: a.RECHTSVORM,
      Gemeente_AAN: a.GEMEENTE,
      Woonplaats_AAN: a.WOONPLAATS,
      POSTCODE_AAN: a.POSTCODE,
      STRAAT_AAN: a.STRAAT,
      HUISNR_AAN: a.HUISNR,
      HUISNRTOEV_AAN: a.HUISNRTOEV,
      Telefoon_AAN: a.TELNR,
      Email_AAN: a.EMAIL,

      NAAM_SIMILARITY_SCORE: best.score,
      Gemeente_matcht_JaNee: gemeenteMatcht ? "Ja" : "Nee",
    });
  }

  log(`Naam-suggesties (Tabblad 4): ${tab4.length} rijen`);

  const summary = [
    { Categorie: "Totaal records CO", Aantal: co.length },
    { Categorie: "Tabblad 1 - KVK match", Aantal: tab1.length },
    { Categorie: "Tabblad 2 - Geen KVK in CO", Aantal: tab2.length },
    {
      Categorie: "Tabblad 3 - Wel KVK in CO, geen match in Aanbieders",
      Aantal: tab3.length,
    },
    {
      Categorie: "Tabblad 4 - Naam-match voor Tabblad 3",
      Aantal: tab4.length,
    },
  ];

  return { summary, tab1, tab2, tab3, tab4 };
}

let latestWorkbookBlob = null;

async function handleRun() {
  const fileAan = document.getElementById("file-aanbieders").files[0];
  const fileCo = document.getElementById("file-co").files[0];
  latestWorkbookBlob = null;
  document.getElementById("download-section").classList.add("hidden");
  if (logEl) logEl.textContent = "";

  if (!fileAan || !fileCo) {
    log("⚠️ Kies zowel een Aanbieders-bestand als een CO non-profit-bestand.");
    return;
  }

  log("Bestanden lezen...");

  try {
    const [aanWb, coWb] = await Promise.all([
      readExcelFile(fileAan),
      readExcelFile(fileCo),
    ]);

    const aanSheet = firstSheetName(aanWb);
    const coSheet = firstSheetName(coWb);

    log(`Aanbieders: gebruik werkblad '${aanSheet}'`);
    log(`CO non-profit: gebruik werkblad '${coSheet}'`);

    const aanRows = sheetToJson(aanWb, aanSheet);
    const coRows = sheetToJson(coWb, coSheet);

    log(`Aanbieders-rijen: ${aanRows.length}`);
    log(`CO-rijen: ${coRows.length}`);

    const { summary, tab1, tab2, tab3, tab4 } = buildResult(aanRows, coRows);

    const outWb = XLSX.utils.book_new();
    const wsSummary = XLSX.utils.json_to_sheet(summary);
    const ws1 = XLSX.utils.json_to_sheet(tab1);
    const ws2 = XLSX.utils.json_to_sheet(tab2);
    const ws3 = XLSX.utils.json_to_sheet(tab3);
    const ws4 = XLSX.utils.json_to_sheet(tab4);

    XLSX.utils.book_append_sheet(outWb, wsSummary, "Samenvatting");
    XLSX.utils.book_append_sheet(outWb, ws1, "Tabblad1_KVK_match");
    XLSX.utils.book_append_sheet(outWb, ws2, "Tabblad2_Geen_KVK");
    XLSX.utils.book_append_sheet(outWb, ws3, "Tabblad3_Wel_KVK_geen_match");
    XLSX.utils.book_append_sheet(outWb, ws4, "Tabblad4_Naam_match");

    const wbout = XLSX.write(outWb, { bookType: "xlsx", type: "array" });
    latestWorkbookBlob = new Blob([wbout], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    document.getElementById("download-section").classList.remove("hidden");
    log("✅ Vergelijken voltooid. Je kunt nu het resultaat downloaden.");

  } catch (err) {
    console.error(err);
    log("❌ Er ging iets mis: " + err.message);
  }
}

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Kon bestand niet lezen."));
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function handleDownload() {
  if (!latestWorkbookBlob) return;
  const url = URL.createObjectURL(latestWorkbookBlob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "resultaat_matcher.xlsx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

document.getElementById("btn-run").addEventListener("click", handleRun);
document.getElementById("btn-download").addEventListener("click", handleDownload);

} // einde if XLSX bestaat
