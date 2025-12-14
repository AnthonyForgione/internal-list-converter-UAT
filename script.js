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

  function parsePartialDate(value) {
    if (isEmpty(value)) return null;
    if (typeof value === "string") {
      const trimmed = value.trim();
      if (/^\d{4}$/.test(trimmed)) return trimmed; // YYYY
      if (/^\d{4}-\d{2}$/.test(trimmed)) return trimmed; // YYYY-MM
      if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed; // YYYY-MM-DD
      const d = new Date(trimmed);
      if (!isNaN(d)) return d.toISOString().slice(0,10);
      return trimmed;
    }
    if (value instanceof Date && !isNaN(value)) {
      return value.toISOString().slice(0,10);
    }
    return String(value);
  }

  function normalizeKey(k) {
    if (!k) return "";
    return String(k).trim().toLowerCase().replace(/[^a-z0-9]/g,"");
  }

  function cleanAndSplit(value) {
    if (isEmpty(value)) return [];
    if (typeof value === "string") {
      return value.split(/[,;]/).map(v => v.trim()).filter(Boolean);
    }
    if (Array.isArray(value)) return value.map(v => String(v).trim()).filter(Boolean);
    return [String(value)];
  }

  function addIfNotEmpty(obj, key, val) {
    if (val !== null && val !== undefined && val !== "" && !(Array.isArray(val) && !val.length)) {
      obj[key] = val;
    }
  }

  /* =========================
     Row transformation
  ========================== */
  function transformRow(row) {
    const normalizedRow = {};
    Object.entries(row).forEach(([k,v]) => {
      normalizedRow[normalizeKey(k)] = v;
    });

    const o = {};

    // Keep profileId as string to avoid Excel date conversion
    addIfNotEmpty(o, "type", normalizedRow["type"]);
    addIfNotEmpty(o, "profileId", normalizedRow["profileid"] ? String(normalizedRow["profileid"]) : null);
    addIfNotEmpty(o, "action", normalizedRow["action"]);
    addIfNotEmpty(o, "activeStatus", normalizedRow["activestatus"]);
    addIfNotEmpty(o, "name", normalizedRow["name"]);
    addIfNotEmpty(o, "suffix", normalizedRow["suffix"]);
    addIfNotEmpty(o, "profileNotes", normalizedRow["profilenotes"]);

    // Array fields
    [
      "countryofregistrationcode","countryofaffiliationcode","formerlysanctionedregioncode",
      "sanctionedregioncode","enhancedriskcountrycode","dateofregistrationarray",
      "dateofbirtharray","residentofcode","citizenshipcode","sources","companyurls"
    ].forEach(f => {
      addIfNotEmpty(o, f, cleanAndSplit(normalizedRow[f]));
    });

    // Identity numbers (robust normalized detection)
    const ids = [];
    Object.entries(row).forEach(([col, val]) => {
      if (isEmpty(val)) return;

      const normCol = normalizeKey(col);

      if (normCol.includes("duns")) {
        ids.push({ type: "duns", value: String(val) });
      } else if (normCol.includes("passport")) {
        ids.push({ type: "passport_no", value: String(val) });
      } else if (normCol.includes("nationaltax")) {
        ids.push({ type: "tax_no", value: String(val) });
      } else if (normCol.includes("lei")) {
        ids.push({ type: "lei", value: String(val) });
      } else if (normCol.includes("nationalid")) {
        ids.push({ type: "national_id", value: String(val) });
      } else if (normCol.includes("drivinglicence")) {
        ids.push({ type: "driving_licence", value: String(val) });
      } else if (normCol.includes("socialsecurity")) {
        ids.push({ type: "ssn", value: String(val) });
      }
    });
    addIfNotEmpty(o, "identityNumbers", ids);

    // Addresses
    const addr = {};
    if (!isEmpty(normalizedRow["addressline"])) addr.line = String(normalizedRow["addressline"]);
    if (!isEmpty(normalizedRow["city"])) addr.city = String(normalizedRow["city"]);
    if (!isEmpty(normalizedRow["province"])) addr.province = String(normalizedRow["province"]);
    if (!isEmpty(normalizedRow["postcode"])) addr.postCode = String(normalizedRow["postcode"]).replace(/\.0$/,"");
    if (!isEmpty(normalizedRow["countrycode"])) addr.countryCode = String(normalizedRow["countrycode"]).toUpperCase().slice(0,2);
    if (Object.keys(addr).length) addIfNotEmpty(o,"addresses",[addr]);

    // Aliases
    const aliases = [];
    Object.keys(row).forEach(col=>{
      if (normalizeKey(col).startsWith("aliases")) {
        if (!isEmpty(row[col])) aliases.push({name:String(row[col]), type:"Also Known As"});
      }
    });
    addIfNotEmpty(o,"aliases",aliases);

    // Lists
    const lists = [];
    for (let i=1;i<=4;i++){
      if (isEmpty(row[`List ${i}`])) continue;
      const e = {};
      const v = row[`List ${i}`];
      addIfNotEmpty(e,"id",v);
      addIfNotEmpty(e,"name",v);
      const active = String(row[`Active List ${i}`]).toLowerCase()==="true";
      e.active = active;
      e.listActive = active;
      if (!isEmpty(v)) e.hierarchy=[{id:v,name:v}];
      addIfNotEmpty(e,"since",parsePartialDate(row[`Since List ${i}`]));
      addIfNotEmpty(e,"to",parsePartialDate(row[`To List ${i}`]));
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
      const rows = XLSX.utils.sheet_to_json(sheet, { defval:null, raw:false });

      const records = rows.map(transformRow);
      if (!records.length) {
        output.textContent = "❌ No valid rows found for conversion.";
        downloadLink.style.display = "none";
        return;
      }

      const jsonl = records.map(r=>JSON.stringify(r)).join("\n");
      output.textContent = jsonl.slice(0,4000)+(jsonl.length>4000?"\n\n...preview truncated...":"");

      const blob = new Blob([jsonl],{type:"application/json"});
      const url = URL.createObjectURL(blob);
      downloadLink.href = url;
      downloadLink.download = "output.jsonl";
      downloadLink.style.display = "block";
      downloadLink.textContent = "Download JSONL file";

    } catch(err){
      output.textContent = "❌ Error: "+err.message;
      console.error(err);
    }
  }
});
