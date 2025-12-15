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
    return v === null || v === undefined || String(v).trim() === "";
  }

  function normalizeKey(k) {
    return String(k)
      .normalize("NFKD")
      .toLowerCase()
      .replace(/\s+/g, "")
      .replace(/[^\w]/g, "");
  }

  function parsePartialDate(v) {
    if (isEmpty(v)) return null;
    const s = String(v).trim();
    if (/^\d{4}(-\d{2})?(-\d{2})?$/.test(s)) return s;
    return s;
  }

  function splitToArray(v) {
    if (isEmpty(v)) return null;
    return String(v).split(/[,;]/).map(x => x.trim()).filter(Boolean);
  }

  /* =========================
     Canonical ASAM mapping
  ========================== */
  const ARRAY_FIELDS = {
    countryofregistrationcode: "countryOfRegistrationCode",
    countryofaffiliationcode: "countryOfAffiliationCode",
    formerlysanctionedregioncode: "formerlySanctionedRegionCode",
    sanctionedregioncode: "sanctionedRegionCode",
    enhancedriskcountrycode: "enhancedRiskCountryCode",
    dateofregistrationarray: "dateOfRegistrationArray",
    dateofbirtharray: "dateOfBirthArray",
    residentofcode: "residentOfCode",
    citizenshipcode: "citizenshipCode",
    sources: "sources",
    companyurls: "companyUrls"
  };

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
    o.type = n.type;
    o.profileId = n.profileid ? String(n.profileid) : undefined;
    o.action = n.action;
    o.activeStatus = n.activestatus;
    o.name = n.name;
    o.suffix = n.suffix;
    o.profileNotes = n.profilenotes;

    // Array fields (CORRECT camelCase output)
    Object.entries(ARRAY_FIELDS).forEach(([nk, outKey]) => {
      const arr = splitToArray(n[nk]);
      if (arr) o[outKey] = arr;
    });

    // Identity numbers (Passport FIXED)
    const ids = [];
    if (!isEmpty(n.nationaltaxno)) ids.push({ type: "tax_no", value: String(n.nationaltaxno) });
    if (!isEmpty(n.dunsnumber)) ids.push({ type: "duns", value: String(n.dunsnumber) });
    if (!isEmpty(n.legalentityidentifierlei)) ids.push({ type: "lei", value: String(n.legalentityidentifierlei) });
    if (!isEmpty(n.nationalid)) ids.push({ type: "national_id", value: String(n.nationalid) });
    if (!isEmpty(n.drivinglicenceno)) ids.push({ type: "driving_licence", value: String(n.drivinglicenceno) });
    if (!isEmpty(n.socialsecurityno)) ids.push({ type: "ssn", value: String(n.socialsecurityno) });
    if (!isEmpty(n.passportno)) ids.push({ type: "passport_no", value: String(n.passportno) });

    if (ids.length) o.identityNumbers = ids;

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
      const v = row[`List ${i}`];
      if (isEmpty(v)) continue;
      const e = {
        id: v,
        name: v,
        active: String(row[`Active List ${i}`]).toLowerCase() === "true",
        listActive: String(row[`Active List ${i}`]).toLowerCase() === "true",
        hierarchy: [{ id: v, name: v }]
      };
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
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

    const records = rows.map(transformRow);
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
  }
});
