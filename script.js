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
  function isEmpty(v) {
    if (v === undefined || v === null) return true;
    if (typeof v === "string" && v.trim() === "") return true;
    if (typeof v === "number" && isNaN(v)) return true;
    if (Array.isArray(v) && v.length === 0) return true;
    return false;
  }

  function add(obj, key, value) {
    if (!isEmpty(value)) obj[key] = value;
  }

  function normalizeKey(k) {
    return String(k || "")
      .normalize("NFKD")
      .toLowerCase()
      .replace(/\s+/g, "")
      .replace(/[^\w]/g, "");
  }

  function split(value) {
    if (isEmpty(value)) return [];
    return String(value)
      .split(/[,;]/)
      .map(v => v.trim())
      .filter(Boolean);
  }

  function parsePartialDate(value) {
    if (isEmpty(value)) return undefined;
    const t = String(value).trim();
    if (/^\d{4}$/.test(t)) return t;
    if (/^\d{4}-\d{2}$/.test(t)) return t;
    if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
    return t;
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

    // Basic fields
    if (!isEmpty(n.type)) o.type = n.type;
    if (!isEmpty(n.profileid)) o.profileId = String(n.profileid);
    if (!isEmpty(n.action)) o.action = n.action;
    if (!isEmpty(n.activestatus)) o.activeStatus = n.activestatus;
    if (!isEmpty(n.name)) o.name = n.name;
    if (!isEmpty(n.suffix)) o.suffix = n.suffix;
    if (!isEmpty(n.profilenotes)) o.profileNotes = n.profilenotes;

    // Array fields (ONLY if non-empty)
    const arrays = {
      countryOfRegistrationCode: split(n.countryofregistrationcode),
      countryOfAffiliationCode: split(n.countryofaffiliationcode),
      formerlySanctionedRegionCode: split(n.formerlysanctionedregioncode),
      sanctionedRegionCode: split(n.sanctionedregioncode),
      enhancedRiskCountryCode: split(n.enhancedriskcountrycode),
      dateOfRegistrationArray: split(n.dateofregistrationarray),
      dateOfBirthArray: split(n.dateofbirtharray),
      residentOfCode: split(n.residentofcode),
      citizenshipCode: split(n.citizenshipcode),
      sources: split(n.sources),
      companyUrls: split(n.companyurls)
    };

    Object.entries(arrays).forEach(([k, v]) => {
      if (v.length) o[k] = v;
    });

    // Identity numbers
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

    // Address
    const addr = {};
    if (!isEmpty(n.addressline)) addr.line = n.addressline;
    if (!isEmpty(n.city)) addr.city = n.city;
    if (!isEmpty(n.province)) addr.province = n.province;
    if (!isEmpty(n.postcode)) addr.postCode = String(n.postcode).replace(/\.0$/, "");
    if (!isEmpty(n.countrycode)) addr.countryCode = String(n.countrycode).toUpperCase().slice(0, 2);

    if (Object.keys(addr).length) o.addresses = [addr];

    // Lists
    const lists = [];
    for (let i = 1; i <= 4; i++) {
      const name = row[`List ${i}`];
      if (isEmpty(name)) continue;

      const e = {
        id: name,
        name,
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
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: undefined, raw: false });

    const records = rows.map(transformRow).filter(r => Object.keys(r).length);

    const jsonl = records.map(r => JSON.stringify(r)).join("\n");

    output.textContent = jsonl.slice(0, 4000);
    const blob = new Blob([jsonl], { type: "application/json" });
    const url = URL.createObjectURL(blob);

    downloadLink.href = url;
    downloadLink.download = "output.jsonl";
    downloadLink.style.display = "block";
  }
});
