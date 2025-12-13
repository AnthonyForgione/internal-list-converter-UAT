/* =========================
   Utility helpers
========================= */

function isEmpty(value) {
  if (value === null || value === undefined) return true;

  if (typeof value === "string") {
    const v = value.trim().toLowerCase();
    return v === "" || v === "nan";
  }

  if (typeof value === "number" && isNaN(value)) return true;

  return false;
}

function parseDateToYMD(value) {
  if (value instanceof Date && !isNaN(value)) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "string") {
    const parsed = new Date(value);
    if (!isNaN(parsed)) {
      return parsed.toISOString().slice(0, 10);
    }
  }

  return null;
}

function cleanAndSplit(value) {
  if (isEmpty(value)) return [];

  // Try datetime conversion
  const parsedDate = parseDateToYMD(value);
  if (parsedDate) return [parsedDate];

  if (typeof value === "string") {
    if (value.includes(",")) {
      return value.split(",").map(v => v.trim()).filter(Boolean);
    }
    if (value.includes(";")) {
      return value.split(";").map(v => v.trim()).filter(Boolean);
    }
    return [value.trim()];
  }

  return [value];
}

function addFieldIfNotEmpty(target, key, value) {
  if (
    value !== null &&
    value !== undefined &&
    value !== "" &&
    !(Array.isArray(value) && value.length === 0)
  ) {
    target[key] = value;
  }
}

/* =========================
   Core transformation
========================= */

function transformRowToJson(row, dynamicAliasColumns) {
  const jsonObj = {};

  function getVal(col) {
    let val = row[col];

    if (isEmpty(val)) return null;

    const parsedDate = parseDateToYMD(val);
    if (parsedDate) return parsedDate;

    return val;
  }

  function getArrayVal(col) {
    return cleanAndSplit(row[col]);
  }

  // Direct mappings
  [
    "type",
    "profileId",
    "action",
    "activeStatus",
    "name",
    "suffix",
    "gender",
    "profileNotes",
    "lastModifiedDate"
  ].forEach(f => addFieldIfNotEmpty(jsonObj, f, getVal(f)));

  [
    "countryOfRegistrationCode",
    "countryOfAffiliationCode",
    "formerlySanctionedRegionCode",
    "sanctionedRegionCode",
    "enhancedRiskCountryCode",
    "dateOfRegistrationArray",
    "dateOfBirthArray",
    "residentOfCode",
    "citizenshipCode",
    "sources",
    "companyUrls"
  ].forEach(f => addFieldIfNotEmpty(jsonObj, f, getArrayVal(f)));

  /* ---------- Identity Numbers ---------- */

  const identityNumbers = [];
  const entityType = String(jsonObj.type || "").toUpperCase();

  const taxNo = getVal("National Tax No.");
  if (!isEmpty(taxNo)) {
    identityNumbers.push({ type: "tax_no", value: String(taxNo) });
  }

  if (entityType === "COMPANY") {
    const duns = getVal("Duns Number");
    if (!isEmpty(duns)) identityNumbers.push({ type: "duns", value: String(duns) });

    const lei = getVal("Legal Entity Identifier (LEI)");
    if (!isEmpty(lei)) identityNumbers.push({ type: "lei", value: String(lei) });
  }

  if (entityType === "PERSON") {
    const nid = getVal("National ID");
    if (!isEmpty(nid)) identityNumbers.push({ type: "national_id", value: String(nid) });

    const dl = getVal("Driving Licence No.\t");
    if (!isEmpty(dl)) identityNumbers.push({ type: "driving_licence", value: String(dl) });

    const ssn = getVal("Social Security No.");
    if (!isEmpty(ssn)) identityNumbers.push({ type: "ssn", value: String(ssn) });

    const passport = getVal("Passport No.\t");
    if (!isEmpty(passport)) identityNumbers.push({ type: "passport_no", value: String(passport) });
  }

  addFieldIfNotEmpty(jsonObj, "identityNumbers", identityNumbers);

  /* ---------- Address ---------- */

  const address = {};

  if (!isEmpty(row["Address Line"])) address.line = String(row["Address Line"]);
  if (!isEmpty(row.city)) address.city = String(row.city);
  if (!isEmpty(row.province)) address.province = String(row.province);

  if (!isEmpty(row.postCode)) {
    address.postCode = String(row.postCode).replace(/\.0$/, "");
  }

  if (!isEmpty(row.countryCode)) {
    address.countryCode = String(row.countryCode).toUpperCase().slice(0, 2);
  }

  if (Object.keys(address).length > 0) {
    jsonObj.a
