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
  ========================= */

  function isEmpty(value) {
    if (value === null || value === undefined) return true;
    if (typeof value === "number" && isNaN(value)) return true;
    if (typeof value === "string") {
      const v = value.trim().toLowerCase();
      return v === "" || v === "nan";
    }
    if (Array.isArray(value) && value.length === 0) return true;
    return false;
  }

  function parsePartialDate(value) {
    if (isEmpty(value)) return null;
    let str = String(value).trim();
    // Handle Excel dates stored as numbers
    if (!isNaN(str) && Number(str) > 30) {
      const d = new Date(Math.round((Number(str) - 25569)*86400*1000));
      str = d.toISOString().slice(0,10);
    }
    // Match YYYY or YYYY-MM or YYYY-MM-DD
    const match = str.match(/^(\d{4})(?:-(\d{2}))?(?:-(\d{2}))?/);
    if (!match) return str;
    const [_, y, m, d] = match;
    if (y && m && d) return `${y}-${m}-${d}`;
    if (y && m) return `${y}-${m}`;
    return y;
  }

  function cleanAndSplit(value) {
    if (isEmpty(value)) return [];
    const parsedDate = parsePartialDate(value);
    if (parsedDate) return [parsedDate];
    if (typeof value === "string") {
      if (value.includes(",")) return value.split(",").map(v => v.trim()).filter(Boolean);
      if (value.includes(";")) return value.split(";").map(v => v.trim()).filter(Boolean);
      return [value.trim()];
    }
    return [value];
  }

  function addIfNotEmpty(obj, key, val) {
    if (!isEmpty(val)) obj[key] = val;
  }

  function normalizeColName(col) {
    if (!col) return "";
    return String(col).trim().replace(/\t/g,'').toLowerCase();
  }

  /* =========================
     Row transformation
  ========================= */

  function transformRow(row, aliasCols, dateCols) {
    const o = {};

    // Normalize keys once
    const normalizedRow = {};
    Object.entries(row).forEach(([k,v]) => {
      normalizedRow[normalizeColName(k)] = v;
    });

    const getVal = c => {
      if (isEmpty(normalizedRow[c])) return null;
      if (c === "profileid") return String(normalizedRow[c]).trim();
      if (dateCols.has(c)) return parsePartialDate(normalizedRow[c]);
      return String(normalizedRow[c]).trim();
    };

    const getArr = c => cleanAndSplit(normalizedRow[c]);

    // Basic fields
    ["type","profileid","action","activeStatus","name","suffix","gender","profileNotes","lastModifiedDate"]
      .forEach(f => addIfNotEmpty(o,f,getVal(f)));

    ["countryofregistrationcode","countryofaffiliationcode",
     "formerlysanctionedregioncode","sanctionedregioncode","enhancedriskcountrycode",
     "dateofregistrationarray","dateofbirtharray","residentofcode","citizenshipcode",
     "sources","companyurls"
    ].forEach(f => addIfNotEmpty(o,f,getArr(f)));

    // Identity numbers
    const ids = [];
    const type = String(getVal("type") || "").toUpperCase();

    const tax = getVal("nationaltaxno");
    if (!isEmpty(tax)) ids.push({ type: "tax_no", value: String(tax) });

    if (type === "COMPANY") {
      const duns = getVal("dunsnumber");
      const lei = getVal("legalentityidentifier(lei)");
      if (!isEmpty(duns)) ids.push({ type: "duns", value: String(duns) });
      if (!isEmpty(lei)) ids.push({ type: "lei", value: String(lei) });
    }

    if (type === "PERSON") {
      [
        ["nationalid","national_id"],
        ["drivinglicenceno","driving_licence"],
        ["socialsecurityno","ssn"],
        ["passportno","passport_no"]
      ].forEach(([col,t]) => {
        const v = getVal(col);
        if (!isEmpty(v)) ids.push({ type: t, value: String(v) });
      });
    }
    addIfNotEmpty(o,"identityNumbers",ids);

    // Address
    const addr = {};
    if (!isEmpty(normalizedRow["addressline"])) addr.line = String(normalizedRow["addressline"]);
    if (!isEmpty(normalizedRow["city"])) addr.city = String(normalizedRow["city"]);
    if (!isEmpty(normalizedRow["province"])) addr.province = String(normalizedRow["province"]);
    if (!isEmpty(normalizedRow["postcode"])) addr.postCode = String(normalizedRow["postcode"]).replace(/\.0$/,'');
    if (!isEmpty(normalizedRow["countrycode"])) addr.countryCode = String(normalizedRow["countrycode"]).toUpperCase().slice(0,2);
    if (Object.keys(addr).length) addIfNotEmpty(o,"addresses",[addr]);

    // Aliases
    const aliases = [];
    aliasCols.forEach(c => {
      if (!isEmpty(normalizedRow[c])) aliases.push({ name: String(normalizedRow[c]), type: "Also Known As" });
    });
    addIfNotEmpty(o,"aliases",aliases);

    // Lists
    const lists = [];
    for (let i=1;i<=4;i++){
      if (isEmpty(normalizedRow[`list ${i}`])) continue;
      const e = {};
      const v = getVal(`list ${i}`);
      addIfNotEmpty(e,"id",v);
      addIfNotEmpty(e,"name",v);
      const active = String(normalizedRow[`activelist ${i}`]).toLowerCase()==="true";
      e.active = active;
      e.listActive = active;
      if (!isEmpty(v)) e.hierarchy=[{id:v,name:v}];
      addIfNotEmpty(e,"since",getVal(`since list ${i}`));
      addIfNotEmpty(e,"to",getVal(`to list ${i}`));
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
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

      if (!rows.length) {
        output.textContent = "❌ No rows found in the file.";
        downloadLink.style.display = "none";
        return;
      }

      // Detect alias columns and date columns
      const aliasCols = Object.keys(rows[0] || {}).filter(c => normalizeColName(c).startsWith("aliases"));
      const dateCols = new Set(Object.keys(rows[0] || {}).map(c => normalizeColName(c)).filter(c => c.includes("date")));

      const records = rows.map(r => transformRow(r, aliasCols, dateCols));
      const jsonl = records.map(r => JSON.stringify(r)).join("\n");

      const blob = new Blob([jsonl], { type: "application/json" });
      const url = URL.createObjectURL(blob);

      downloadLink.href = url;
      downloadLink.download = "output.jsonl";
      downloadLink.style.display = "block";
      downloadLink.textContent = "Download JSONL";

      // Preview first 4000 chars
      output.textContent = jsonl.slice(0,4000) + (jsonl.length>4000 ? "\n\n...preview truncated..." : "");
    } catch(err) {
      output.textContent = "❌ Error: " + err.message;
      console.error(err);
    }
  }

});
