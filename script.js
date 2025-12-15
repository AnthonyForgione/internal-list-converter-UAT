/* ========================= 
   DOM READY
========================= */
document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("fileInput");
  const convertBtn = document.getElementById("convertBtn");
  const output = document.getElementById("output");
  const downloadLink = document.getElementById("downloadLink");

  convertBtn.addEventListener("click", () => {
    if (!fileInput.files.length) {
      output.textContent = "❌ Please select an Excel file first.";
      return;
    }
    processExcel(fileInput.files[0]);
  });

  /* =========================
     Utilities
  ========================== */
  function isEmpty(value) {
    if (value === null || value === undefined) return true;
    if (typeof value === "number" && isNaN(value)) return true;
    if (typeof value === "string") {
      const v = value.trim().toLowerCase();
      return v === "" || v === "nan";
    }
    if (Array.isArray(value)) return value.length === 0;
    return false;
  }

  function addIfNotEmpty(obj, key, value) {
    if (!isEmpty(value)) {
      obj[key] = value;
    }
  }

  function parsePartialDate(value) {
    if (isEmpty(value)) return null;

    if (typeof value === "string") {
      const t = value.trim();
      if (/^\d{4}$/.test(t)) return t;
      if (/^\d{4}-\d{2}$/.test(t)) return t;
      if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
      return t;
    }

    if (value instanceof Date && !isNaN(value)) {
      return value.toISOString().slice(0, 10);
    }

    return String(value);
  }

  function normalizeKey(k) {
    return String(k || "")
      .normalize("NFKD")
      .toLowerCase()
      .replace(/\s+/g, "")
      .replace(/[^\w]/g, "");
  }

  function cleanAndSplit(value) {
    if (isEmpty(value)) return [];
    return String(value)
      .split(/[,;]/)
      .map(v => v.trim())
      .filter(Boolean);
  }

  /* =========================
     Row transformation
  ========================== */
  function transformRow(row) {
    const n = {};
    Object.entries(row).forEach(([k, v]) => {
      n[normalizeKey(k)] = v;
    });

    const o = {};

    /* -------- Basic fields -------- */
    addIfNotEmpty(o, "type", n.type);
    addIfNotEmpty(o, "profileId", n.profileid ? String(n.profileid) : null);
    addIfNotEmpty(o, "action", n.action);
    addIfNotEmpty(o, "activeStatus", n.activestatus);
    addIfNotEmpty(o, "name", n.name);
    addIfNotEmpty(o, "suffix", n.suffix);
    addIfNotEmpty(o, "profileNotes", n.profilenotes);

    /* -------- Array fields -------- */
    const arrayMap = {
      countryOfRegistrationCode: "countryofregistrationcode",
      countryOfAffiliationCode: "countryofaffiliationcode",
      formerlySanctionedRegionCode: "formerlysanctionedregioncode",
      sanctionedRegionCode: "sanctionedregioncode",
      enhancedRiskCountryCode: "enhancedriskcountrycode",
      dateOfRegistrationArray: "dateofregistrationarray",
      dateOfBirthArray: "dateofbirtharray",
      residentOfCode: "residentofcode",
      citizenshipCode: "citizenshipcode",
      sources: "sources",
      companyUrls: "companyurls"
    };

    Object.entries(arrayMap).forEach(([outKey, inKey]) => {
      const arr = cleanAndSplit(n[inKey]);
      if (arr.length) o[outKey] = arr;
    });

    /* -------- Identity numbers -------- */
    const identities = [];
    Object.entries(row).forEach(([col, val]) => {
      if (isEmpty(val)) return;

      const key = normalizeKey(col);

      if (key.includes("passportno")) identities.push({ type: "passport_no", value: String(val) });
      else if (key.includes("duns")) identities.push({ type: "duns", value: String(val) });
      else if (key.includes("nationaltax")) identities.push({ type: "tax_no", value: String(val) });
      else if (key.includes("lei")) identities.push({ type: "lei", value: String(val) });
      else if (key.includes("nationalid")) identities.push({ type: "national_id", value: String(val) });
      else if (key.includes("drivinglicence")) identities.push({ type: "driving_licence", value: String(val) });
      else if (key.includes("socialsecurity")) identities.push({ type: "ssn", value: String(val) });
    });

    if (identities.length) o.identityNumbers = identities;

    /* -------- Addresses -------- */
    const addr = {};
    addIfNotEmpty(addr, "line", n.addressline);
    addIfNotEmpty(addr, "city", n.city);
    addIfNotEmpty(addr, "province", n.province);
    addIfNotEmpty(addr, "postCode", n.postcode ? String(n.postcode).replace(/\.0$/, "") : null);
    addIfNotEmpty(addr, "countryCode", n.countrycode ? String(n.countrycode).toUpperCase().slice(0, 2) : null);

    if (Object.keys(addr).length) o.addresses = [addr];

    /* -------- Aliases -------- */
    const aliases = [];
    Object.keys(row).forEach(col => {
      if (normalizeKey(col).startsWith("aliases") && !isEmpty(row[col])) {
        aliases.push({ name: String(row[col]), type: "Also Known As" });
      }
    });

    if (aliases.length) o.aliases = aliases;

    /* -------- Lists -------- */
    const lists = [];
    for (let i = 1; i <= 4; i++) {
      const name = row[`List ${i}`];
      if (isEmpty(name)) continue;

      const e = {
        id: name,
        name: name,
        hierarchy: [{ id: name, name }]
      };

      const active = String(row[`Active List ${i}`]).toLowerCase() === "true";
      e.active = active;
      e.listActive = active;

      const since = parsePartialDate(row[`Since List ${i}`]);
      const to = parsePartialDate(row[`To List ${i}`]);
      if (since) e.since = since;
      if (to) e.to = to;

      lists.push(e);
    }

    if (lists.length) o.lists = lists;

    return o;
  }

  /* =========================
     File processing
  ========================== */
  async function processExcel(file) {
    try {
      output.textContent = "⏳ Processing file...";
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: false });

      const records = rows.map(transformRow).filter(r => Object.keys(r).length);

      if (!records.length) {
        output.textContent = "❌ No valid rows found.";
        downloadLink.style.display = "none";
        return;
      }

      const jsonl = records.map(r => JSON.stringify(r)).join("\n");

      output.textContent =
        jsonl.slice(0, 4000) +
        (jsonl.length > 4000 ? "\n\n...preview truncated..." : "");

      const blob = new Blob([jsonl], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      downloadLink.href = url;
      downloadLink.download = "output.jsonl";
      downloadLink.style.display = "block";
      downloadLink.textContent = "Download JSONL file";

    } catch (err) {
      output.textContent = "❌ Error: " + err.message;
      console.error(err);
    }
  }
});
