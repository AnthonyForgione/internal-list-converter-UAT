/* =========================
   DOM READY
========================= */

document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("fileInput");
  const convertBtn = document.getElementById("convertBtn");
  const output = document.getElementById("output");
  const downloadLink = document.getElementById("downloadLink");

  // Columns that should NEVER be treated as dates or numbers
  const NEVER_DATE_COLUMNS = new Set([
    "profileId",
    "type",
    "Duns Number",
    "Legal Entity Identifier (LEI)",
    "National Tax No.",
    "National ID",
    "Driving Licence No.",
    "Social Security No.",
    "Passport No."
  ]);

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
    return false;
  }

  function parseDateToYMD(value) {
    if (value instanceof Date) {
      return value.toISOString().slice(0, 10);
    }
    const d = new Date(value);
    return isNaN(d) ? null : d.toISOString().slice(0, 10);
  }

  function cleanAndSplit(value) {
    if (isEmpty(value)) return [];

    const parsedDate = parseDateToYMD(value);
    if (parsedDate) return [parsedDate];

    if (typeof value === "string") {
      if (value.includes(",")) return value.split(",").map(v => v.trim()).filter(Boolean);
      if (value.includes(";")) return value.split(";").map(v => v.trim()).filter(Boolean);
      return [value.trim()];
    }
    return [value];
  }

  function addIfNotEmpty(obj, key, val) {
    if (val !== null && val !== undefined && val !== "" && !(Array.isArray(val) && !val.length)) {
      obj[key] = val;
    }
  }

  /* =========================
     Row transformation
  ========================== */

  function transformRow(row, aliasCols) {
    const o = {};

    const getVal = c => {
      if (isEmpty(row[c])) return null;

      // Force NEVER_DATE_COLUMNS to string
      if (NEVER_DATE_COLUMNS.has(c)) return String(row[c]).trim();

      // Otherwise parse dates if possible
      return parseDateToYMD(row[c]) || row[c];
    };

    const getArr = c => cleanAndSplit(row[c]);

    // Core fields
    [
      "type","profileId","action","activeStatus","name","suffix","gender",
      "profileNotes","lastModifiedDate"
    ].forEach(f => addIfNotEmpty(o, f, getVal(f)));

    // Array fields
    [
      "countryOfRegistrationCode","countryOfAffiliationCode",
      "formerlySanctionedRegionCode","sanctionedRegionCode",
      "enhancedRiskCountryCode","dateOfRegistrationArray",
      "dateOfBirthArray","residentOfCode","citizenshipCode",
      "sources","companyUrls"
    ].forEach(f => addIfNotEmpty(o, f, getArr(f)));

    /* Identity numbers */
    const ids = [];
    const type = String(o.type || "").toUpperCase();

    const tax = getVal("National Tax No.");
    if (!isEmpty(tax)) ids.push({ type: "tax_no", value: String(tax) });

    if (type === "COMPANY") {
      const duns = getVal("Duns Number");
      const lei = getVal("Legal Entity Identifier (LEI)");
      if (!isEmpty(duns)) ids.push({ type: "duns", value: String(duns) });
      if (!isEmpty(lei)) ids.push({ type: "lei", value: String(lei) });
    }

    if (type === "PERSON") {
      [["National ID","national_id"],
       ["Driving Licence No.","driving_licence"],
       ["Social Security No.","ssn"],
       ["Passport No.","passport_no"]
      ].forEach(([c,t]) => {
        const v = getVal(c);
        if (!isEmpty(v)) ids.push({ type: t, value: String(v) });
      });
    }

    addIfNotEmpty(o, "identityNumbers", ids);

    /* Address */
    const addr = {};
    if (!isEmpty(row["Address Line"])) addr.line = String(row["Address Line"]);
    if (!isEmpty(row.city)) addr.city = String(row.city);
    if (!isEmpty(row.province)) addr.province = String(row.province);
    if (!isEmpty(row.postCode)) addr.postCode = String(row.postCode).replace(/\.0$/, "");
    if (!isEmpty(row.countryCode)) addr.countryCode = String(row.countryCode).toUpperCase().slice(0,2);
    if (Object.keys(addr).length) o.addresses = [addr];

    /* Aliases */
    const aliases = [];
    aliasCols.forEach(c => {
      if (!isEmpty(row[c])) aliases.push({ name: String(row[c]), type: "Also Known As" });
    });
    addIfNotEmpty(o, "aliases", aliases);

    /* Lists */
    const lists = [];
    for (let i=1;i<=4;i++) {
      if (isEmpty(row[`List ${i}`])) continue;
      const e = {};
      const v = getVal(`List ${i}`);
      addIfNotEmpty(e,"id",v);
      addIfNotEmpty(e,"name",v);
      const active = String(row[`Active List ${i}`]).toLowerCase()==="true";
      e.active = active;
      e.listActive = active;
      if (!isEmpty(v)) e.hierarchy=[{id:v,name:v}];
      addIfNotEmpty(e,"since",getVal(`Since List ${i}`));
      addIfNotEmpty(e,"to",getVal(`To List ${i}`));
      lists.push(e);
    }
    addIfNotEmpty(o,"lists",lists);

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

      const rows = XLSX.utils.sheet_to_json(sheet, {
        defval: null,
        raw: true  // keep Excel values exactly as-is
      });

      if (!rows.length) {
        output.textContent = "❌ No rows found in the sheet.";
        downloadLink.style.display = "none";
        return;
      }

      const aliasCols = Object.keys(rows[0] || {})
        .filter(c => c.toLowerCase().startsWith("aliases") && /^\d+$/.test(c.slice(7)));

      const records = rows.map(r => transformRow(r, aliasCols));
      const jsonl = records.map(r => JSON.stringify(r)).join("\n");

      const blob = new Blob([jsonl], { type: "application/jsonl" });
      const url = URL.createObjectURL(blob);

      downloadLink.href = url;
      downloadLink.download = "output.jsonl";
      downloadLink.style.display = "block";
      downloadLink.textContent = "Download JSONL";

      // Preview first 4000 chars
      output.textContent = jsonl.slice(0, 4000) + (jsonl.length > 4000 ? "\n\n...preview truncated..." : "");
    } catch (err) {
      output.textContent = "❌ Error: " + err.message;
      console.error(err);
    }
  }
});
