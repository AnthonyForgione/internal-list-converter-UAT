// script.js
// Browser-side XLS/XLSX -> JSONL converter for ASAM client import
// Preserves full field structure including optional fields and correct address keys
// Requires SheetJS (xlsx.full.min.js)

(function () {

  // Utility helpers
  function isEmpty(value) {
    if (value === null || value === undefined) return true;
    if (typeof value === 'number' && isNaN(value)) return true;
    if (typeof value === 'string') return value.trim() === '';
    if (Array.isArray(value)) return value.length === 0;
    if (value instanceof Date) return isNaN(value.getTime());
    if (typeof value === 'object') return Object.keys(value).length === 0;
    return false;
  }

  function _to_string_id(value) {
    if (typeof value === 'number' && Number.isInteger(value)) return String(value);
    if (value instanceof Date) return String(value.valueOf());
    return String(value);
  }

  function _to_string_list(value) {
    if (isEmpty(value)) return null;
    if (Array.isArray(value)) return value.map(v => String(v).trim()).filter(Boolean);
    if (typeof value === 'string') return value.split(',').map(s => s.trim()).filter(Boolean);
    return [String(value)];
  }

  function _to_unix_timestamp_ms(value) {
    if (isEmpty(value)) return null;
    if (value instanceof Date) {
      if (isNaN(value.getTime())) return null;
      return value.getTime();
    }
    if (typeof value === 'number') {
      if (value > 1e12) return Math.floor(value);
      if (value > 1e9) return Math.floor(value * 1000);
    }
    const parsed = Date.parse(String(value));
    if (!isNaN(parsed)) return parsed;
    return null;
  }

  function normalizeKey(k) {
    if (k === undefined || k === null) return '';
    return String(k).replace(/^"+|"+$/g, '').trim();
  }

  function normalizeRowKeys(row) {
    const out = {};
    for (const [k, v] of Object.entries(row)) {
      out[normalizeKey(k)] = v;
    }
    return out;
  }

  function transformRowToClientJson(rowRaw) {
    const row = normalizeRowKeys(rowRaw);
    const clientData = { objectType: 'client' };

    function addFieldIfNotEmpty(key, value) {
      if (!isEmpty(value)) clientData[key] = value;
    }

    // Primary fields
    addFieldIfNotEmpty('clientId', row['clientId']);
    addFieldIfNotEmpty('entityType', row['entityType']);
    addFieldIfNotEmpty('status', row['status']);

    const entityTypeUpper = row['entityType'] ? String(row['entityType']).toUpperCase() : null;

    // Name fields
    if (entityTypeUpper === 'ORGANISATION' || entityTypeUpper === 'ORGANIZATION') {
      addFieldIfNotEmpty('companyName', row['name']);
    } else {
      addFieldIfNotEmpty('name', row['name']);
      addFieldIfNotEmpty('forename', row['forename']);
      addFieldIfNotEmpty('middlename', row['middlename']);
      addFieldIfNotEmpty('surname', row['surname']);
    }

    // Common fields as strings
    addFieldIfNotEmpty('titles', row['titles']);
    addFieldIfNotEmpty('suffixes', row['suffixes']);

    // Person-specific fields
    if (entityTypeUpper === 'PERSON') {
      let genderValue = row['gender'];
      if (typeof genderValue === 'string' && !isEmpty(genderValue)) genderValue = genderValue.toUpperCase();
      addFieldIfNotEmpty('gender', genderValue);
      addFieldIfNotEmpty('dateOfBirth', row['dateOfBirth'] ? String(row['dateOfBirth']) : null);
      addFieldIfNotEmpty('birthPlaceCountryCode', row['birthPlaceCountryCode']);
      addFieldIfNotEmpty('deceasedOn', row['deceasedOn'] ? String(row['deceasedOn']) : null);
      addFieldIfNotEmpty('occupation', row['occupation']);
      addFieldIfNotEmpty('domicileCodes', _to_string_list(row['domicileCodes']));
      addFieldIfNotEmpty('nationalityCodes', _to_string_list(row['nationalityCodes']));
    }

    // Organisation-specific
    if (entityTypeUpper === 'ORGANISATION' || entityTypeUpper === 'ORGANIZATION') {
      addFieldIfNotEmpty('incorporationCountryCode', row['incorporationCountryCode']);
      addFieldIfNotEmpty('dateOfIncorporation', row['dateOfIncorporation'] ? String(row['dateOfIncorporation']) : null);
    }

    // Assessment required boolean
    const assessmentRequiredRawValue = row['assessmentRequired'];
    let assessmentRequiredBoolean = false;
    if (!isEmpty(assessmentRequiredRawValue)) {
      const rawStr = String(assessmentRequiredRawValue).toLowerCase();
      assessmentRequiredBoolean = ['true','1','1.0','t','yes','y'].includes(rawStr);
    }
    if (!isEmpty(assessmentRequiredRawValue)) {
      addFieldIfNotEmpty('assessmentRequired', assessmentRequiredBoolean);
    }

    // lastReviewed only if assessmentRequired
    if (assessmentRequiredBoolean) {
      addFieldIfNotEmpty('lastReviewed', _to_unix_timestamp_ms(row['lastReviewed']));
    }

    // Optional review fields
    addFieldIfNotEmpty('periodicReviewStartDate', _to_unix_timestamp_ms(row['periodicReviewStartDate']));
    addFieldIfNotEmpty('periodicReviewPeriod', row['periodicReviewPeriod'] ? String(row['periodicReviewPeriod']) : null);
    addFieldIfNotEmpty('segment', row['segment'] ? String(row['segment']) : null);

    // Addresses (keep keys intact)
    const currentAddress = {};
    const addressKeys = ['line1','line2','line3','line4','poBox','city','state','province','postcode','country','countryCode'];
    addressKeys.forEach(k => {
      if (!isEmpty(row[k])) currentAddress[k] = String(row[k]);
    });
    if (Object.keys(currentAddress).length > 0) addFieldIfNotEmpty('addresses', [currentAddress]);

    // Aliases
    const aliasColumns = ['aliases1','aliases2','aliases3','aliases4'];
    const aliasNameTypes = { 'aliases1':'AKA1','aliases2':'AKA2','aliases3':'AKA3','aliases4':'AKA4' };
    const aliasesList = [];
    for (const col of aliasColumns) {
      const val = row[col];
      if (!isEmpty(val)) {
        const nameType = aliasNameTypes[col] || col.toUpperCase();
        if (entityTypeUpper === 'PERSON') {
          aliasesList.push({name:String(val), nameType});
        } else {
          aliasesList.push({companyName:String(val), nameType});
        }
      }
    }
    if (aliasesList.length > 0) addFieldIfNotEmpty('aliases', aliasesList);

    return clientData;
  }

  // UI handling
  function init() {
    const fileInput = document.getElementById('fileInput');
    const convertBtn = document.getElementById('convertBtn');
    const outputEl = document.getElementById('output');
    const downloadLink = document.getElementById('downloadLink');

    convertBtn.addEventListener('click', () => {
      if (!fileInput.files || fileInput.files.length === 0) {
        alert('Please choose an XLS/XLSX file first.');
        return;
      }

      const file = fileInput.files[0];
      const reader = new FileReader();

      reader.onload = function (e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, {type: 'array', cellDates: true});
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, {defval: null, raw: false});

        const transformed = rows.map(transformRowToClientJson);

        if (transformed.length === 0) {
          outputEl.textContent = 'No valid rows found for conversion.';
          downloadLink.style.display = 'none';
          return;
        }

        const jsonlLines = transformed.map(rec => JSON.stringify(rec));
        const jsonlContent = jsonlLines.join('\n');

        outputEl.textContent = jsonlContent.slice(0, 4000) + (jsonlContent.length>4000?'\n\n...preview truncated...':'');
        const blob = new Blob([jsonlContent], {type:'application/json'});
        const url = URL.createObjectURL(blob);
        downloadLink.href = url;
        downloadLink.download = file.name.replace(/\.[^/.]+$/, '') + '.jsonl';
        downloadLink.style.display = 'inline-block';
        downloadLink.textContent = 'Download JSONL file';
      };

      reader.readAsArrayBuffer(file);
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
