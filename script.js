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
      // if it's a Date and invalid
      if (value instanceof Date) return isNaN(value.getTime());
      // treat empty object as empty
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
    // If it's already a Date object
    if (value instanceof Date) {
      if (isNaN(value.getTime())) return null;
      return value.getTime();
    }
    // If it's an Excel serial number (SheetJS might give numbers for dates)
    if (typeof value === 'number') {
      // Heuristic: treat as Excel serial if it's reasonably large/small? Safer to attempt Date from value if not.
      // But do not try to convert number blindly. We'll convert numeric -> Date via JS Date if it looks like epoch seconds/millis.
      // If > 1e12 treat as ms epoch; if between 1e9 and 1e12 treat as seconds -> ms
      if (value > 1e12) return Math.floor(value); // already ms
      if (value > 1e9) return Math.floor(value * 1000); // seconds -> ms
      // Otherwise treat as Excel serial: SheetJS often returns JS Date if cellDates:true; so rarely here.
    }
    // Try parsing string with Date
    const parsed = Date.parse(String(value));
    if (!isNaN(parsed)) return parsed;
    return null;
  }

  // normalize header keys: remove quotes, trim whitespace
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

  // Main row transformer (mirrors the Python transform_row_to_client_json)
  function transformRowToClientJson(rowRaw) {
    const row = normalizeRowKeys(rowRaw);
    const clientData = {};
    clientData.objectType = 'client';

    function addFieldIfNotEmpty(key, value) {
      if (!isEmpty(value)) clientData[key] = value;
    }

    function maybeString(v) {
      if (v === null || v === undefined) return v;
      return String(v);
    }

    // read primary fields
    addFieldIfNotEmpty('clientId', row['clientId']);
    addFieldIfNotEmpty('entityType', row['entityType']);
    addFieldIfNotEmpty('status', row['status']);

    let entityTypeUpper = null;
    if (!isEmpty(row['entityType'])) {
      entityTypeUpper = String(row['entityType']).toUpperCase();
    }

    // 2. Name fields
    if (entityTypeUpper === 'ORGANISATION' || entityTypeUpper === 'ORGANIZATION') {
      addFieldIfNotEmpty('companyName', row['name']);
    } else if (entityTypeUpper === 'PERSON') {
      addFieldIfNotEmpty('name', row['name']);
      addFieldIfNotEmpty('forename', row['forename']);
      addFieldIfNotEmpty('middlename', row['middlename']);
      addFieldIfNotEmpty('surname', row['surname']);
    } else {
      // fallback
      addFieldIfNotEmpty('name', row['name']);
      addFieldIfNotEmpty('forename', row['forename']);
      addFieldIfNotEmpty('middlename', row['middlename']);
      addFieldIfNotEmpty('surname', row['surname']);
    }

    // Common fields
    addFieldIfNotEmpty('titles', row['titles']);
    addFieldIfNotEmpty('suffixes', row['suffixes']);

    // Person-specific
    if (entityTypeUpper === 'PERSON') {
      let genderValue = row['gender'];
      if (typeof genderValue === 'string' && !isEmpty(genderValue)) genderValue = genderValue.toUpperCase();
      addFieldIfNotEmpty('gender', genderValue);

      const dob = row['dateOfBirth'];
      addFieldIfNotEmpty('dateOfBirth', isEmpty(dob) ? dob : String(dob));

      addFieldIfNotEmpty('birthPlaceCountryCode', row['birthPlaceCountryCode']);

      const deceasedOn = row['deceasedOn'];
      addFieldIfNotEmpty('deceasedOn', isEmpty(deceasedOn) ? deceasedOn : String(deceasedOn));

      addFieldIfNotEmpty('occupation', row['occupation']);
      addFieldIfNotEmpty('domicileCodes', _to_string_list(row['domicileCodes']));
      addFieldIfNotEmpty('nationalityCodes', _to_string_list(row['nationalityCodes']));
    }

    // Organisation-specific
    if (entityTypeUpper === 'ORGANISATION' || entityTypeUpper === 'ORGANIZATION') {
      addFieldIfNotEmpty('incorporationCountryCode', row['incorporationCountryCode']);
      const doi = row['dateOfIncorporation'];
      addFieldIfNotEmpty('dateOfIncorporation', isEmpty(doi) ? doi : String(doi));
    }

    // assessmentRequired boolean parsing
    const assessmentRequiredRawValue = row['assessmentRequired'];
    let assessmentRequiredBoolean = false;
    if (!isEmpty(assessmentRequiredRawValue)) {
      const rawStr = String(assessmentRequiredRawValue).toLowerCase();
      assessmentRequiredBoolean = ['true', '1', '1.0', 't', 'yes', 'y'].includes(rawStr);
    }

    // lastReviewed (only if assessmentRequired true) -> unix ms
    if (assessmentRequiredBoolean) {
      addFieldIfNotEmpty('lastReviewed', _to_unix_timestamp_ms(row['lastReviewed']));
    }

    addFieldIfNotEmpty('periodicReviewStartDate', _to_unix_timestamp_ms(row['periodicReviewStartDate']));

    const periodic_review_period_value = row['periodicReviewPeriod'];
    addFieldIfNotEmpty('periodicReviewPeriod', isEmpty(periodic_review_period_value) ? periodic_review_period_value : String(periodic_review_period_value));

    // Addresses
    const currentAddress = {};
    if (!isEmpty(row['Address line1'])) currentAddress.line1 = String(row['Address line1']);
    if (!isEmpty(row['Address line2'])) currentAddress.line2 = String(row['Address line2']);
    if (!isEmpty(row['Address line3'])) currentAddress.line3 = String(row['Address line3']);
    if (!isEmpty(row['Address line4'])) currentAddress.line4 = String(row['Address line4']);
    if (!isEmpty(row['poBox'])) currentAddress.poBox = String(row['poBox']);
    if (!isEmpty(row['city'])) currentAddress.city = String(row['city']);
    if (!isEmpty(row['state'])) currentAddress.state = String(row['state']);
    if (!isEmpty(row['province'])) currentAddress.province = String(row['province']);
    if (!isEmpty(row['postcode'])) currentAddress.postcode = String(row['postcode']);
    if (!isEmpty(row['country'])) currentAddress.country = String(row['country']);
    if (!isEmpty(row['countryCode'])) currentAddress.countryCode = String(row['countryCode']).toUpperCase().substring(0,2);

    if (Object.keys(currentAddress).length > 0) addFieldIfNotEmpty('addresses', [currentAddress]);

    addFieldIfNotEmpty('segment', isEmpty(row['segment']) ? row['segment'] : String(row['segment']));

    // identityNumbers
    const identityNumbersList = [];
    if (entityTypeUpper === 'ORGANISATION' || entityTypeUpper === 'ORGANIZATION') {
      if (!isEmpty(row['Duns Number'])) identityNumbersList.push({type:'duns', value: _to_string_id(row['Duns Number'])});
      if (!isEmpty(row['National Tax No.'])) identityNumbersList.push({type:'tax_no', value: _to_string_id(row['National Tax No.'])});
      if (!isEmpty(row['Legal Entity Identifier (LEI)'])) identityNumbersList.push({type:'lei', value: _to_string_id(row['Legal Entity Identifier (LEI)'])});
    } else if (entityTypeUpper === 'PERSON') {
      if (!isEmpty(row['National ID'])) identityNumbersList.push({type:'national_id', value: _to_string_id(row['National ID'])});
      if (!isEmpty(row['Driving Licence No.'])) identityNumbersList.push({type:'driving_licence', value: _to_string_id(row['Driving Licence No.'])});
      if (!isEmpty(row['Social Security No.'])) identityNumbersList.push({type:'ssn', value: _to_string_id(row['Social Security No.'])});
      if (!isEmpty(row['Passport No.'])) identityNumbersList.push({type:'passport_no', value: _to_string_id(row['Passport No.'])});
    }
    if (identityNumbersList.length > 0) clientData.identityNumbers = identityNumbersList;

    // aliases (aliases1..4)
    const aliasColumns = ['aliases1','aliases2','aliases3','aliases4'];
    const aliasNameTypes = {
      'aliases1':'AKA1','aliases2':'AKA2','aliases3':'AKA3','aliases4':'AKA4'
    };
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
    if (aliasesList.length > 0) clientData.aliases = aliasesList;

    // security object
    const securityEnabled = row['Security Enabled'];
    if (!isEmpty(securityEnabled) && ['true','t','1','yes','y'].includes(String(securityEnabled).toLowerCase())) {
      const securityTags = {};
      if (!isEmpty(row['Tag 1'])) securityTags.orTags1 = row['Tag 1'];
      if (!isEmpty(row['Tag 2'])) securityTags.orTags2 = row['Tag 2'];
      if (!isEmpty(row['Tag 3'])) securityTags.orTags3 = row['Tag 3'];
      clientData.security = securityTags;
    }

    if (!isEmpty(assessmentRequiredRawValue)) {
      addFieldIfNotEmpty('assessmentRequired', assessmentRequiredBoolean);
    }

    return clientData;
  }

  // Main UI handling
  function init() {
    const fileInput = document.getElementById('fileInput');
    const convertBtn = document.getElementById('convertBtn');
    const outputEl = document.getElementById('output');
    const downloadLink = document.getElementById('downloadLink');

    if (!fileInput || !convertBtn || !outputEl || !downloadLink) {
      console.warn('Expected elements: #fileInput, #convertBtn, #output, #downloadLink');
    }

    convertBtn.addEventListener('click', () => {
      if (!fileInput.files || fileInput.files.length === 0) {
        alert('Please choose an XLS/XLSX file first.');
        return;
      }

      const file = fileInput.files[0];
      const reader = new FileReader();

      reader.onload = function (e) {
        const data = e.target.result;
        // prefer array buffer read and type 'array'
        const workbook = XLSX.read(data, {type: 'array', cellDates: true});
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];

        // Convert sheet rows to JS objects (header row as keys)
        const rows = XLSX.utils.sheet_to_json(sheet, {defval: null, raw: false});

        // Apply transformation
        const transformedRaw = rows.map(transformRowToClientJson);

        // Filter out sparse records (similar to your Python logic)
        const transformed = transformedRaw.filter(record => {
          const hasEntityType = record.entityType && !isEmpty(record.entityType);
          const hasNameInfo = ['name','forename','surname','companyName'].some(k => record[k] && !isEmpty(record[k]));
          return hasEntityType || hasNameInfo;
        });

        if (transformed.length === 0) {
          outputEl.textContent = 'No valid rows found for conversion.';
          downloadLink.style.display = 'none';
          return;
        }

        // Build JSONL content
        const jsonlLines = transformed.map(rec => JSON.stringify(rec));
        const jsonlContent = jsonlLines.join('\n');

        // Show preview (first 4000 chars)
        outputEl.textContent = jsonlContent.slice(0, 4000) + (jsonlContent.length > 4000 ? '\n\n...preview truncated...' : '');

        // Prepare download
        const blob = new Blob([jsonlContent], {type: 'application/json'});
        const url = URL.createObjectURL(blob);
        downloadLink.href = url;
        const filename = file.name.replace(/\.[^/.]+$/, '') + '.jsonl';
        downloadLink.download = filename;
        downloadLink.style.display = 'inline-block';
        downloadLink.textContent = 'Download JSONL file';
      };

      // Read as array buffer for SheetJS
      reader.readAsArrayBuffer(file);
    });
  }

  // run when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
