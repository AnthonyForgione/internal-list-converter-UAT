// script.js
// Browser-side XLS/XLSX -> JSONL converter adapted from the user's Python logic.
// Requires SheetJS (xlsx.full.min.js) loaded in index.html.

(function () {
  // Utility helpers
  function isEmpty(value) {
    if (value === null || value === undefined) return true;
    if (typeof value === 'number' && isNaN(value)) return true;
    if (typeof value === 'string') return value.trim() === '';
    if (Array.isArray(value)) return value.length === 0;
    if (typeof value === 'object') {
      if (value instanceof Date) return isNaN(value.getTime());
      return Object.keys(value).length === 0;
    }
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
    if (typeof value === 'string') {
      return value.split(',').map(s => s.trim()).filter(Boolean);
    }
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

    const entityTypeUpper = row['entityType'] ? String(row['entityType']).toUpperCase() : null;

    // Primary fields
    addFieldIfNotEmpty('clientId', row['clientId']);
    addFieldIfNotEmpty('entityType', row['entityType']);
    addFieldIfNotEmpty('status', row['status']);

    // Name fields
    if (entityTypeUpper === 'ORGANISATION' || entityTypeUpper === 'ORGANIZATION') {
      addFieldIfNotEmpty('companyName', row['name']);
    } else {
      addFieldIfNotEmpty('name', row['name']);
      addFieldIfNotEmpty('forename', row['forename']);
      addFieldIfNotEmpty('middlename', row['middlename']);
      addFieldIfNotEmpty('surname', row['surname']);
    }

    // Common fields
    addFieldIfNotEmpty('titles', _to_string_list(row['titles']));
    addFieldIfNotEmpty('suffixes', _to_string_list(row['suffixes']));

    // Person-specific
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

    // Addresses
    const currentAddress = {};
    ['Address line1','Address line2','Address line3','Address line4','poBox','city','state','province','postcode','country','countryCode']
      .forEach(k => { if (!isEmpty(row[k])) currentAddress[k==='countryCode'?k:(k==='poBox'?'poBox':k.replace(' ','').toLowerCase())] = String(row[k]); });
    if (Object.keys(currentAddress).length > 0) addFieldIfNotEmpty('addresses', [currentAddress]);

    return clientData;
  }

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
        console.log('Workbook loaded:', workbook.SheetNames);

        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];

        const rows = XLSX.utils.sheet_to_json(sheet, {defval: null, raw: false});
        console.log('First 2 rows:', rows.slice(0,2));

        const transformedRaw = rows.map(transformRowToClientJson);
        console.log('Transformed first row:', transformedRaw[0]);

        // TEMP: bypass filter for debug
        const transformed = transformedRaw; 

        if (transformed.length === 0) {
          outputEl.textContent = 'No valid rows found for conversion.';
          downloadLink.style.display = 'none';
          return;
        }

        const jsonlLines = transformed.map(rec => JSON.stringify(rec));
        const jsonlContent = jsonlLines.join('\n');

        outputEl.textContent = jsonlContent.slice(0,4000) + (jsonlContent.length>4000?'\n\n...preview truncated...':'');
        const blob = new Blob([jsonlContent], {type: 'application/json'});
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
