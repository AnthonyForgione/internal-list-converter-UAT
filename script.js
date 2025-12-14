document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("fileInput");
  const convertBtn = document.getElementById("convertBtn");
  const output = document.getElementById("output");
  const downloadLink = document.getElementById("downloadLink");

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
  ========================= */

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
  ========================= */

  function transformRow(row, aliasCols) {
    const o = {};

    const getVal = c => {
      if (isEmpty(row[c])) return null;
      return NEVER_DATE_COLUMNS.has(c) ? String(row[c]) : parseDateToYMD(row[c]) || row[c];
    };

    const getArr = c => cleanAndSplit(row[c]);

    [
      "type","profileId","action","activeStatus","name","suffix","gender",
      "profileNotes","lastModifiedDate"
    ].forEach(f => addIfNotEmpty(o, f, getVal(f)));

    [
      "countryOfRegistrationCode","countryOfAffiliationCode",
      "formerlySanctionedRegionCode","sanctionedRegionCode",
      "enhancedRiskCountryCode","dateOfRegistrationArray",
      "dateOfBirthArray","residentOfCode","citizenshipCode",
      "sources","companyUrls"
    ].forEach(f => addIfNotEmpty(o, f, getArr(f)));

    /* Identity numbers */
    const ids = [];
    const typeUpper = String(o.type || "").toUpperCase();

    const tax = getVal("National Tax No.");
    if (!isEmpty(tax)) ids.push({ type: "tax_no", value: String(tax) });

    if (typeUpper === "COMPANY") {
      const duns = getVal("Duns Number");
      const lei = getVal("Legal Entity Identifier (LEI)");
      if (!isEmpty(duns)) ids.push({ type: "duns", value: String(duns) });
      if (!isEmpty(lei)) ids.push({ type: "lei", value: String(lei) });
    }

    if (typeUpper === "PERSON") {
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
  ========================= */

  async function processExcel(file) {
    try {
      output.textContent = "⏳ Processing file...";
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      let rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });

      // Coerce NEVER_DATE_COLUMNS to string to prevent Excel auto-date conversion
      rows.forEach(row => {
        NEVER_DATE_COLUMNS.forEach(col => {
          if (col in row && row[col] !== null && row[col] !== undefined) {
            row[col] = String(row[col]);
          }
        });
      });

      const aliasCols = Object.keys(rows[0] || {})
        .filter(c => c.toLowerCase().startsWith("aliases") && /^\d*$/.test(c.slice(7)));

      const records = rows.map(r => transformRow(r, aliasCols));
      if (!records.length) {
        output.textContent = "❌ No valid rows found for conversion.";
        downloadLink.style.display = "none";
        return;
      }

      const jsonl = records.map(r => JSON.stringify(r)).join("\n");

      // Show preview
      output.textContent = jsonl.slice(0, 4000) + (jsonl.length > 4000 ? "\n\n...preview truncated..." : "");

      // Prepare download
      const blob = new Blob([jsonl], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      downloadLink.href = url;
      downloadLink.download = file.name.replace(/\.[^/.]+$/, "") + ".jsonl";
      downloadLink.style.display = "block";
      downloadLink.textContent = "Download JSONL";

    } catch (err) {
      output.textContent = "❌ Error: " + err.message;
      console.error(err);
    }
  }
});
