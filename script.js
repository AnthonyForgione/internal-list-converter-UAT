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
    Utilities (Python-equivalent)
    ========================== */

    function isEmpty(value) {
        if (value === null || value === undefined) return true;
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
            const t = value.trim();
            if (/^\d{4}$/.test(t)) return t;
            if (/^\d{4}-\d{2}$/.test(t)) return t;
            if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
            const d = new Date(t);
            if (!isNaN(d)) return d.toISOString().slice(0, 10);
            return t;
        }
        if (value instanceof Date && !isNaN(value)) {
            return value.toISOString().slice(0, 10);
        }
        return String(value);
    }

    function cleanAndSplit(value) {
        if (isEmpty(value)) return [];
        if (typeof value === "string") {
            return value
                .split(/[,;]/)
                .map(v => v.trim())
                .filter(Boolean);
        }
        return [String(value)];
    }

    function normalizeKey(k) {
        return String(k || "")
            .normalize("NFKD")
            .toLowerCase()
            .replace(/\s+/g, "")
            .replace(/[^\w]/g, "");
    }

    function addFieldIfNotEmpty(obj, key, value) {
        if (
            value !== null &&
            value !== undefined &&
            value !== "" &&
            !(Array.isArray(value) && value.length === 0)
        ) {
            obj[key] = value;
        }
    }

    /* =========================
    Row Transformation
    ========================== */
    function transformRow(row) {
        const r = {};
        Object.entries(row).forEach(([k, v]) => {
            r[normalizeKey(k)] = v;
        });

        const o = {};

        /* Basic fields */
        addFieldIfNotEmpty(o, "type", r.type);
        addFieldIfNotEmpty(o, "profileId", r.profileid ? String(r.profileid) : null);
        addFieldIfNotEmpty(o, "action", r.action);
        addFieldIfNotEmpty(o, "activeStatus", r.activestatus);
        addFieldIfNotEmpty(o, "name", r.name);
        addFieldIfNotEmpty(o, "suffix", r.suffix);
        addFieldIfNotEmpty(o, "profileNotes", r.profilenotes);
        addFieldIfNotEmpty(o, "lastModifiedDate", parsePartialDate(r.lastmodifieddate));

        /* Array fields */
        [
            ["countryOfRegistrationCode", "countryofregistrationcode"],
            ["countryOfAffiliationCode", "countryofaffiliationcode"],
            ["formerlySanctionedRegionCode", "formerlysanctionedregioncode"],
            ["sanctionedRegionCode", "sanctionedregioncode"],
            ["enhancedRiskCountryCode", "enhancedriskcountrycode"],
            ["dateOfRegistrationArray", "dateofregistrationarray"],
            ["dateOfBirthArray", "dateofbirtharray"],
            ["residentOfCode", "residentofcode"],
            ["citizenshipCode", "citizenshipcode"],
            ["sources", "sources"],
            ["companyUrls", "companyurls"]
        ].forEach(([outKey, inKey]) => {
            addFieldIfNotEmpty(o, outKey, cleanAndSplit(r[inKey]));
        });

        /* =========================
        Identity Numbers (Python-parity)
        ========================== */
        const ids = [];
        const typeUpper = String(o.type || "").toUpperCase();

        if (!isEmpty(r.nationaltaxno) && r.nationaltaxno.toString().trim().toLowerCase() !== "nan") {
            ids.push({ type: "tax_no", value: String(r.nationaltaxno) });
        }

        if (typeUpper === "COMPANY") {
            if (!isEmpty(r.dunsnumber) && r.dunsnumber.toString().trim().toLowerCase() !== "nan") {
                ids.push({ type: "duns", value: String(r.dunsnumber) });
            }
            if (!isEmpty(r.legalentityidentifierlei) && r.legalentityidentifierlei.toString().trim().toLowerCase() !== "nan") {
                ids.push({ type: "lei", value: String(r.legalentityidentifierlei) });
            }
        } else if (typeUpper === "PERSON") {
            if (!isEmpty(r.nationalid) && r.nationalid.toString().trim().toLowerCase() !== "nan") {
                ids.push({ type: "national_id", value: String(r.nationalid) });
            }
            if (!isEmpty(r.drivinglicenceno) && r.drivinglicenceno.toString().trim().toLowerCase() !== "nan") {
                ids.push({ type: "driving_licence", value: String(r.drivinglicenceno) });
            }
            if (!isEmpty(r.socialsecurityno) && r.socialsecurityno.toString().trim().toLowerCase() !== "nan") {
                ids.push({ type: "ssn", value: String(r.socialsecurityno) });
            }
            if (!isEmpty(r.passportno) && r.passportno.toString().trim().toLowerCase() !== "nan") {
                ids.push({ type: "passport_no", value: String(r.passportno) });
            }
        }
        addFieldIfNotEmpty(o, "identityNumbers", ids);

        /* =========================
        Address
        ========================== */
        const addr = {};
        addFieldIfNotEmpty(addr, "line", r.addressline);
        addFieldIfNotEmpty(addr, "city", r.city);
        addFieldIfNotEmpty(addr, "province", r.province);
        addFieldIfNotEmpty(
            addr,
            "postCode",
            r.postcode ? String(r.postcode).replace(/\.0$/, "") : null
        );
        addFieldIfNotEmpty(
            addr,
            "countryCode",
            r.countrycode ? String(r.countrycode).toUpperCase().slice(0, 2) : null
        );
        addFieldIfNotEmpty(o, "addresses", Object.keys(addr).length ? [addr] : []);

        /* =========================
        Aliases
        ========================== */
        const aliases = [];
        Object.keys(row).forEach(col => {
            if (normalizeKey(col).startsWith("aliases") && !isEmpty(row[col])) {
                aliases.push({ name: String(row[col]), type: "Also Known As" });
            }
        });
        addFieldIfNotEmpty(o, "aliases", aliases);

        /* =========================
        Lists
        ========================== */
        const lists = [];
        for (let i = 1; i <= 4; i++) {
            const listVal = row[`List ${i}`];
            if (isEmpty(listVal)) continue;

            const e = {};
            addFieldIfNotEmpty(e, "id", listVal);
            addFieldIfNotEmpty(e, "name", listVal);

            const active = String(row[`Active List ${i}`]).toLowerCase() === "true";
            addFieldIfNotEmpty(e, "active", active);
            addFieldIfNotEmpty(e, "listActive", active);

            addFieldIfNotEmpty(e, "hierarchy", [{ id: listVal, name: listVal }]);
            addFieldIfNotEmpty(e, "since", parsePartialDate(row[`Since List ${i}`]));
            addFieldIfNotEmpty(e, "to", parsePartialDate(row[`To List ${i}`]));

            lists.push(e);
        }
        addFieldIfNotEmpty(o, "lists", lists);

        /* =========================
        Python-parity cleanup: remove empty fields
        ========================== */
        Object.keys(o).forEach(k => {
            if (isEmpty(o[k])) delete o[k];
        });

        return o;
    }

    /* =========================
    File Processing
    ========================== */
    async function processExcel(file) {
        output.textContent = "⏳ Processing file...";
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: false });

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
