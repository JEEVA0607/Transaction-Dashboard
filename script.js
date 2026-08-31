let originalData = [];
let filteredData = [];

let statusChart = null;
let typeChart = null;
let creditDebitChart = null;
let hourChart = null;


// =============================================
// FILE UPLOAD
// =============================================

const excelFile = document.getElementById("excelFile");

if (excelFile) {
    excelFile.addEventListener("change", handleFile);
}


function handleFile(event) {

    const file = event.target.files[0];

    if (!file) return;

    const reader = new FileReader();

    reader.onload = function (e) {

        try {

            const workbook = XLSX.read(
                e.target.result,
                { type: "array", cellDates: true }
            );

            if (!workbook.SheetNames.length) {
                alert("Excel file-il sheet kandethan pattiyilla.");
                return;
            }

            const sheet =
                workbook.Sheets[workbook.SheetNames[0]];

            const data =
                XLSX.utils.sheet_to_json(
                    sheet,
                    { defval: "" }
                );

            originalData = data.map(cleanRow);

            filteredData = [...originalData];

            initializeFilters();

            renderDashboard();

        } catch (error) {

            console.error(error);

            alert(
                "Excel file read cheyyan pattiyilla."
            );
        }
    };

    reader.readAsArrayBuffer(file);
}


// =============================================
// CLEAN DATA
// =============================================

function cleanRow(row) {

    const clean = {};

    Object.keys(row).forEach(key => {

        const cleanKey =
            String(key)
                .trim()
                .replace(/\s+/g, " ");

        clean[cleanKey] = row[key];

    });

    return clean;
}


// =============================================
// COLUMN FINDER
// =============================================

function getColumn(row, possibleNames) {

    if (!row) return "";

    const keys = Object.keys(row);

    for (const wanted of possibleNames) {

        const wantedClean =
            String(wanted)
                .toLowerCase()
                .trim()
                .replace(/\s+/g, " ");

        const exact =
            keys.find(key =>
                String(key)
                    .toLowerCase()
                    .trim()
                    .replace(/\s+/g, " ")
                    === wantedClean
            );

        if (exact !== undefined) {
            return row[exact];
        }
    }

    return "";
}


// =============================================
// UTR / TRANSACTION REFERENCE
// =============================================

function getUTR(row) {

    const explicit = String(
        getColumn(row, [
            "UTR",
            "UTR No",
            "UTR Number",
            "UTRNO",
            "Transaction UTR",
            "Transaction ID",
            "Txn ID",
            "Reference",
            "Reference No"
        ]) || ""
    ).trim();

    if (explicit) return explicit;

    return extractTransactionReference(row);
}

// DPB UTR/reference is embedded in the remark, e.g.
// DPB-1821045324-W444-31458593
function extractTransactionReference(row) {

    const text = `${getEntryType(row)} ${getRemark(row)}`;

    const match = text.match(/\bDPB-[^-\s]+-[^-\s]+-[^-\s]+/i);

    return match ? match[0].toUpperCase() : "";
}

// =============================================
// REVD / REV DETECTION
// =============================================

function isREVD(row) {

    const text = `${getEntryType(row)} ${getRemark(row)}`.toUpperCase();

    return (
        /\bREVD\b/.test(text) ||
        /\bREV\b/.test(text) && text.includes("DPB-") ||
        text.includes("REVERSE DEBIT") ||
        text.includes("REVERSED DEBIT")
    );
}

// =============================================
// REMARK
// =============================================

function getRemark(row) {

    return String(
        getColumn(row, [
            "Remark",
            "Remarks",
            "Description",
            "Narration",
            "Transaction Remark"
        ]) || ""
    ).trim();
}


// =============================================
// USER / FROM TO
// =============================================

function getUser(row) {

    return String(
        getColumn(row, [
            "Fromto",
            "From / To",
            "From To",
            "From/To",
            "User",
            "Username",
            "User Name",
            "Name"
        ]) || "Unknown"
    ).trim();
}


// =============================================
// CREDIT
// =============================================

function getCredit(row) {

    return Math.abs(
        num(
            getColumn(row, [
                "Credit"
            ])
        )
    );
}


// =============================================
// DEBIT
// =============================================

function getDebit(row) {

    return Math.abs(
        num(
            getColumn(row, [
                "Debit"
            ])
        )
    );
}


// =============================================
// AMOUNT
// CREDIT / DEBIT ONLY
// PTS IS COMPLETELY IGNORED
// =============================================

function getAmount(row) {

    const credit = getCredit(row);
    const debit = getDebit(row);

    // If Credit exists, use Credit
    if (credit > 0) {
        return credit;
    }

    // Otherwise use Debit
    if (debit > 0) {
        return debit;
    }

    return 0;
}


// =============================================
// NUMBER
// =============================================

function num(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return 0;
    }

    if (typeof value === "number") {
        return isFinite(value) ? value : 0;
    }

    const cleaned =
        String(value)
            .replace(/,/g, "")
            .replace(/[₹$€£]/g, "")
            .trim();

    return Number(cleaned) || 0;
}


// =============================================
// DATE
// =============================================

function parseDate(value) {

    if (value instanceof Date) {

        if (!isNaN(value.getTime())) {
            return value;
        }

        return null;
    }


    if (typeof value === "number") {

        if (value > 20000) {

            const excelEpoch =
                new Date(
                    Date.UTC(1899, 11, 30)
                );

            const date =
                new Date(
                    excelEpoch.getTime()
                    + value * 86400000
                );

            return isNaN(date.getTime())
                ? null
                : date;
        }
    }


    if (!value) return null;


    const stringValue =
        String(value).trim();


    // DD/MM/YYYY or DD-MM-YYYY
    const dmy =
        stringValue.match(
            /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/
        );

    if (dmy) {

        const day = Number(dmy[1]);
        const month = Number(dmy[2]) - 1;
        const year = Number(dmy[3]);

        const date =
            new Date(
                year,
                month,
                day
            );

        return isNaN(date.getTime())
            ? null
            : date;
    }


    const date =
        new Date(stringValue);

    if (!isNaN(date.getTime())) {
        return date;
    }


    return null;
}


function getRowDate(row) {

    return getColumn(row, [
        "Date",
        "Transaction Date",
        "Txn Date",
        "Created Date",
        "Created At",
        "Date Time",
        "Timestamp",
        "Time"
    ]);
}


function formatDateInput(date) {

    if (!date) return "";

    const y =
        date.getFullYear();

    const m =
        String(
            date.getMonth() + 1
        ).padStart(2, "0");

    const d =
        String(
            date.getDate()
        ).padStart(2, "0");

    return `${y}-${m}-${d}`;
}


function formatDate(value) {

    const date =
        parseDate(value);

    if (!date) {
        return value || "";
    }

    return date.toLocaleString(
        "en-IN",
        {
            day: "2-digit",
            month: "2-digit",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit"
        }
    );
}


// =============================================
// ENTRY TYPE
// =============================================

function getEntryType(row) {

    return String(
        getColumn(row, [
            "Type",
            "Entry Type",
            "Transaction Type",
            "Category"
        ]) || ""
    )
        .trim()
        .toUpperCase();
}


// =============================================
// STATUS
// =============================================

function getStatus(row) {

    const text =
        (
            getRemark(row)
            + " "
            + getEntryType(row)
        )
            .toUpperCase();


    if (
        text.includes("REJECTED") ||
        text.includes("REJECT")
    ) {
        return "Rejected";
    }


    if (
        text.includes("HOLD")
    ) {
        return "Hold";
    }


    if (
        text.includes("SUCCESSFUL") ||
        text.includes("SUCCESS") ||
        text.includes("COMPLETED")
    ) {
        return "Successful";
    }


    return "Other";
}


// =============================================
// NORMAL TYPE
// =============================================

function getType(row) {

    const text =
        (
            getRemark(row)
            + " "
            + getEntryType(row)
        )
            .toUpperCase();


    if (/\bESD\b/.test(text))
        return "ESD";


    if (/\bDPB\b/.test(text))
        return "DPB";


    if (/\bUPD\b/.test(text))
        return "UPD";


    if (/\bDTD\b/.test(text))
        return "DTD";


    if (/\bGP\b/.test(text))
        return "GP";


    return "Other";
}


// =============================================
// BONUS TYPE
// =============================================

function getBonusType(row) {

    const type =
        getEntryType(row)
            .replace(/\s+/g, " ")
            .trim();


    const remark =
        getRemark(row)
            .toUpperCase()
            .replace(/\s+/g, " ")
            .trim();


    const text =
        `${type} ${remark}`;


    // REVERSE BONUS
    if (
        text.includes("REVB") ||
        text.includes("REVERSE BONUS") ||
        text.includes("REVERSE-BONUS") ||
        text.includes("REVERSE BONUS")
    ) {
        return "REVB - REVERSE BONUS";
    }


    // REFER BONUS
    if (
        text.includes("REFER - BONUS") ||
        text.includes("REFER BONUS") ||
        text.includes("REFER-BONUS") ||
        text.includes("REFFRAL BONUS") ||
        text.includes("REFERRAL BONUS")
    ) {
        return "REFER - BONUS";
    }


    // WELCOME BONUS
    if (
        text.includes("WELCOME BONUS") ||
        text.includes("WELCOME-BONUS")
    ) {
        return "WELCOME BONUS";
    }


    // WEEKLY CASHBACK
    if (
        text.includes("WEEKLY CASHBACK BONUS") ||
        text.includes("WEEKLY CASHBACK")
    ) {
        return "WEEKLY CASHBACK BONUS";
    }


    // NORMAL BONUS
    if (
        text.includes("BONUS:-") ||
        text.includes("BONUS -") ||
        text.includes("BONUS:") ||
        /\bBONUS\b/.test(text)
    ) {
        return "BONUS:-";
    }


    return null;
}


// =============================================
// BONUS DETECTION
// =============================================

function isBonus(row) {

    return getBonusType(row) !== null;
}


// =============================================
// REVERSAL DETECTION
// =============================================

function isReversal(row) {

    const type = getEntryType(row).toUpperCase();
    const remark = getRemark(row).toUpperCase();
    const text = `${type} ${remark}`;

    return (
        isREVD(row) ||
        /\bREVB\b/.test(text) ||
        /\bREV\b/.test(text) ||
        text.includes("REVERSE BONUS") ||
        text.includes("REVERSAL") ||
        text.includes("REVERSED")
    );
}

// =============================================
// ₹1 DETECTION
// =============================================

function isOneRupee(row) {

    const amount =
        getAmount(row);

    return Math.abs(amount - 1) < 0.000001;
}


// =============================================
// NORMAL CATEGORY
// =============================================

function getCategory(row) {

    const type =
        getType(row);


    if (
        [
            "ESD",
            "DPB",
            "UPD",
            "DTD",
            "GP"
        ].includes(type)
    ) {
        return type;
    }


    return "Other";
}


function classifyRows(rows) {

    const result = {
        normal: [],
        bonus: [],
        reversal: [],
        oneRupee: [],
        revd: []
    };

    // ---------------------------------------------------------
    // REVERSED / DUPLICATE UTR HANDLING
    // ---------------------------------------------------------
    // Some statement files contain a REVD row. Some files contain
    // the reversal credit but the REVD text is missing. In both cases
    // the business rule is the same:
    //
    //   1st debit for that UTR  -> EXCLUDE from main totals
    //   reversal/credit row    -> EXCLUDE from main totals
    //   later debit            -> VALID and COUNT + AMOUNT
    //
    // IMPORTANT: never remove every row sharing the UTR.
    const excludedRows = new Set();
    const rowIndex = new Map();
    originalData.forEach((row, index) => rowIndex.set(row, index));

    const utrGroups = new Map();
    originalData.forEach(row => {
        const utr = getUTR(row);
        if (!utr) return;
        if (!utrGroups.has(utr)) utrGroups.set(utr, []);
        utrGroups.get(utr).push(row);
    });

    utrGroups.forEach(group => {
        const debitRows = group.filter(row => getDebit(row) > 0);
        const creditRows = group.filter(row => getCredit(row) > 0);
        const reversalRows = group.filter(row => isREVD(row) || isReversal(row));

        // A reversal can be explicitly labelled REVD/REV OR can appear
        // only as a credit against the same UTR.
        const hasExplicitReversal = reversalRows.some(row => isReversal(row));

        if (hasExplicitReversal || (debitRows.length && creditRows.length)) {
            // Always exclude only the FIRST debit for this UTR.
            if (debitRows.length) {
                excludedRows.add(debitRows[0]);
            }

            // Explicit reversal: put the reversal row(s) in reversal section.
            // Missing-REVD case: the credit row is also excluded from main
            // totals and shown in reversal section for visibility.
            creditRows.forEach(row => excludedRows.add(row));
        }
    });

    rows.forEach(row => {
        // Bonus is completely separate from main totals.
        if (isBonus(row)) {
            result.bonus.push(row);
            return;
        }

        // ₹1 is completely separate from main totals.
        if (isOneRupee(row)) {
            result.oneRupee.push(row);
            return;
        }

        const utr = getUTR(row);
        const group = utr ? utrGroups.get(utr) || [] : [];
        const hasCreditDebitPair = group.some(r => getCredit(r) > 0) && group.some(r => getDebit(r) > 0);
        const explicitReversal = isReversal(row);

        // Explicit REVD/REV rows always belong to Reversal.
        if (explicitReversal) {
            result.reversal.push(row);
            if (isREVD(row)) result.revd.push(row);
            return;
        }

        // Missing-REVD case: a credit row sharing a UTR with a debit is
        // the reversal credit. Keep it out of main totals and show it in
        // the Reversal section.
        if (hasCreditDebitPair && getCredit(row) > 0) {
            result.reversal.push(row);
            return;
        }

        // First debit of a UTR that has a credit/reversal is excluded.
        if (excludedRows.has(row)) {
            return;
        }

        result.normal.push(row);
    });

    return result;
}


// =============================================
// COPY SUMMARY VALUE
// =============================================

function copySummaryValue(id, mode) {
    const el = document.getElementById(id);
    if (!el) return;

    let value = el.textContent.trim();

    if (mode === "count") {
        // Copy only the number, e.g. "1,234 entries" -> "1234".
        const match = value.replace(/,/g, "").match(/-?\d+(?:\.\d+)?/);
        value = match ? match[0] : "0";
    } else {
        // Copy only the numeric amount, without ₹/$/€ and without commas.
        value = value
            .replace(/[₹$€£]/g, "")
            .replace(/,/g, "")
            .trim();
        const match = value.match(/-?\d+(?:\.\d+)?/);
        value = match ? match[0] : "0";
    }

    navigator.clipboard.writeText(value).then(() => {
        const buttons = document.querySelectorAll(`[data-copy-id="${id}"]`);
        buttons.forEach(btn => {
            const old = btn.textContent;
            btn.textContent = "✓";
            setTimeout(() => btn.textContent = old, 900);
        });
    }).catch(() => {
        const input = document.createElement("textarea");
        input.value = value;
        document.body.appendChild(input);
        input.select();
        document.execCommand("copy");
        input.remove();
    });
}


// =============================================
// FILTER INITIALIZATION
// =============================================

function initializeFilters() {

    const statusSelect =
        document.getElementById("statusFilter");

    const typeSelect =
        document.getElementById("typeFilter");


    if (!statusSelect || !typeSelect)
        return;


    statusSelect.innerHTML =
        `<option value="">All</option>`;


    typeSelect.innerHTML =
        `<option value="">All</option>`;


    const statuses =
        [
            ...new Set(
                originalData.map(
                    row => getStatus(row)
                )
            )
        ];


    const types =
        [
            ...new Set(
                originalData.map(
                    row => getType(row)
                )
            )
        ];


    statuses.forEach(status => {

        statusSelect.innerHTML +=
            `
            <option value="${escapeHTML(status)}">
                ${escapeHTML(status)}
            </option>
            `;

    });


    types.forEach(type => {

        typeSelect.innerHTML +=
            `
            <option value="${escapeHTML(type)}">
                ${escapeHTML(type)}
            </option>
            `;

    });


    const dates =
        originalData
            .map(row =>
                parseDate(
                    getRowDate(row)
                )
            )
            .filter(Boolean);


    if (dates.length) {

        dates.sort(
            (a, b) => a - b
        );


        const fromDate =
            document.getElementById("fromDate");

        const toDate =
            document.getElementById("toDate");


        if (fromDate) {

            fromDate.value =
                formatDateInput(dates[0]);

        }


        if (toDate) {

            toDate.value =
                formatDateInput(
                    dates[dates.length - 1]
                );

        }

    }
}


// =============================================
// FILTER EVENTS
// =============================================

const statusFilter =
    document.getElementById("statusFilter");

if (statusFilter) {

    statusFilter.addEventListener(
        "change",
        applyFilters
    );

}


const typeFilter =
    document.getElementById("typeFilter");

if (typeFilter) {

    typeFilter.addEventListener(
        "change",
        applyFilters
    );

}


const searchFilter =
    document.getElementById("searchFilter");

if (searchFilter) {

    searchFilter.addEventListener(
        "input",
        applyFilters
    );

}


const fromDate =
    document.getElementById("fromDate");

if (fromDate) {

    fromDate.addEventListener(
        "change",
        applyFilters
    );

}


const toDate =
    document.getElementById("toDate");

if (toDate) {

    toDate.addEventListener(
        "change",
        applyFilters
    );

}


// =============================================
// APPLY FILTERS
// =============================================

function applyFilters() {

    const status =
        document.getElementById("statusFilter")?.value || "";


    const type =
        document.getElementById("typeFilter")?.value || "";


    const search =
        (
            document.getElementById("searchFilter")?.value || ""
        )
            .toLowerCase()
            .trim();


    const from =
        document.getElementById("fromDate")?.value || "";


    const to =
        document.getElementById("toDate")?.value || "";


    filteredData =
        originalData.filter(row => {

            const rowDate =
                parseDate(
                    getRowDate(row)
                );


            const rowDateString =
                rowDate
                    ? formatDateInput(rowDate)
                    : "";


            if (
                from &&
                (
                    !rowDateString ||
                    rowDateString < from
                )
            ) {
                return false;
            }


            if (
                to &&
                (
                    !rowDateString ||
                    rowDateString > to
                )
            ) {
                return false;
            }


            if (
                status &&
                getStatus(row) !== status
            ) {
                return false;
            }


            if (
                type &&
                getType(row) !== type
            ) {
                return false;
            }


            if (search) {

                const searchable =
                    Object.values(row)
                        .join(" ")
                        .toLowerCase();


                if (
                    !searchable.includes(search)
                ) {
                    return false;
                }

            }


            return true;

        });


    renderDashboard();
}


// =============================================
// RESET FILTERS
// =============================================

const resetFilters =
    document.getElementById("resetFilters");

if (resetFilters) {

    resetFilters.addEventListener(
        "click",
        () => {

            if (statusFilter)
                statusFilter.value = "";


            if (typeFilter)
                typeFilter.value = "";


            if (searchFilter)
                searchFilter.value = "";


            initializeFilters();


            filteredData =
                [...originalData];


            renderDashboard();

        }
    );

}


// =============================================
// DASHBOARD
// =============================================

function renderDashboard() {

    if (!originalData.length) {
        clearDashboard();
        return;
    }


    updateKPIs();

    renderStatusChart();

    renderTypeChart();

    renderCreditDebitChart();

    renderHourChart();

    renderCreditTable();

    renderDebitTable();

    renderProblemTable();

    renderCategorySection();

    renderWithdrawalSection();

    renderBonusSection();

    renderReversalTable();

    renderOneRupeeTable();

    renderMainTable();
}


// =============================================
// CLEAR DASHBOARD
// =============================================

function clearDashboard() {

    setText("totalCredit", "0.00");
    setText("totalDebit", "0.00");
    setText("netMovement", "0.00");
    setText("transactionCount", "0");
    setText("reversalCount", "0");
    setText("oneRupeeCount", "0");

    setText(
        "mainCategoryCount",
        "0 entries"
    );

    setText(
        "withdrawalCount",
        "0 entries"
    );

    setText("holdCount", "0");
    setText("holdAmount", "₹0.00");
    setText("successfulCount", "0");
    setText("successfulAmount", "₹0.00");

    setText(
        "bonusCount",
        "0 entries"
    );

    setText(
        "reversalCountText",
        "0 entries"
    );

    setText(
        "oneRupeeCountText",
        "0 entries"
    );

    setText(
        "rowCount",
        "0 rows"
    );

    [
        "categoryTable",
        "bonusTable",
        "reversalTable",
        "oneRupeeTable",
        "creditTable",
        "debitTable",
        "problemTable",
        "mainTable"
    ].forEach(id => {

        const element =
            document.getElementById(id);

        if (element) {
            element.innerHTML = "";
        }

    });


    destroyChart(statusChart);
    destroyChart(typeChart);
    destroyChart(creditDebitChart);
    destroyChart(hourChart);


    statusChart = null;
    typeChart = null;
    creditDebitChart = null;
    hourChart = null;
}


// =============================================
// KPI
// =============================================

function updateKPIs() {

    const classified = classifyRows(filteredData);
    const normal = classified.normal;

    const totalCredit = normal.reduce((sum, row) => sum + getCredit(row), 0);
    const totalDebit = normal.reduce((sum, row) => sum + getDebit(row), 0);

    // Net movement is the absolute movement of valid/main transactions.
    const netMovement = totalCredit + totalDebit;

    setText("totalCredit", money(totalCredit));
    setText("totalDebit", money(totalDebit));
    setText("netMovement", money(netMovement));
    setText("transactionCount", normal.length.toLocaleString());
    setText("reversalCount", classified.reversal.length.toLocaleString());
    setText("oneRupeeCount", classified.oneRupee.length.toLocaleString());
}


// =============================================
// CATEGORY SECTION
// =============================================

function renderCategorySection() {

    const normal =
        classifyRows(filteredData).normal;


    const categories = {

        ESD: [],
        DPB: [],
        UPD: [],
        DTD: [],
        GP: []

    };


    normal.forEach(row => {

        const category =
            getCategory(row);


        if (categories[category]) {

            categories[category].push(row);

        }

    });


    Object.keys(categories)
        .forEach(category => {

            const rows =
                categories[category];


            const count =
                rows.length;


            const amount =
                rows.reduce(
                    (sum, row) =>
                        sum + getAmount(row),
                    0
                );


            const id =
                category.toLowerCase();


            setText(
                id + "Count",
                count.toLocaleString()
            );


            setText(
                id + "Amount",
                "₹" + money(amount)
            );

        });


    const allCategoryRows =
        Object.values(categories).flat();


    setText(
        "mainCategoryCount",
        `${allCategoryRows.length.toLocaleString()} entries`
    );


    const table =
        document.getElementById("categoryTable");


    if (!table) return;


    table.innerHTML =
        allCategoryRows.length
            ? allCategoryRows
                .map(categoryRowHTML)
                .join("")
            : emptyRow(6);
}


// =============================================
// CATEGORY ROW
// =============================================

function categoryRowHTML(row) {

    return `
        <tr>

            <td>
                ${escapeHTML(
                    formatDate(
                        getRowDate(row)
                    )
                )}
            </td>

            <td>
                ${escapeHTML(
                    getCategory(row)
                )}
            </td>

            <td class="credit">
                ₹${money(
                    getAmount(row)
                )}
            </td>

            <td>
                ${escapeHTML(
                    getRemark(row)
                )}
            </td>

            <td>
                ${escapeHTML(
                    getUser(row)
                )}
            </td>

            <td>
                ${escapeHTML(
                    getUTR(row)
                )}
            </td>

        </tr>
    `;
}


// =============================================
// WITHDRAWAL SECTION
// =============================================

function renderWithdrawalSection() {

    // Withdrawal cards intentionally use the raw filtered rows.
    // Hold and Successful are separate status buckets and do not enter
    // ESD/DPB/UPD/DTD/GP totals.
    const holdRows = filteredData.filter(row => getStatus(row) === "Hold");
    const successfulRows = filteredData.filter(row => getStatus(row) === "Successful");

    const holdAmount = holdRows.reduce(
        (sum, row) => sum + getAmount(row),
        0
    );

    const successfulAmount = successfulRows.reduce(
        (sum, row) => sum + getAmount(row),
        0
    );

    setText("holdCount", holdRows.length.toLocaleString());
    setText("holdAmount", "₹" + money(holdAmount));
    setText("successfulCount", successfulRows.length.toLocaleString());
    setText("successfulAmount", "₹" + money(successfulAmount));
    setText(
        "withdrawalCount",
        `${(holdRows.length + successfulRows.length).toLocaleString()} entries`
    );
}


// =============================================
// BONUS SECTION
// =============================================

function renderBonusSection() {

    const rows =
        classifyRows(filteredData).bonus;


    const bonusTypes = {

        "REFER - BONUS": [],
        "WELCOME BONUS": [],
        "BONUS:-": [],
        "WEEKLY CASHBACK BONUS": [],
        "REVB - REVERSE BONUS": []

    };


    rows.forEach(row => {

        const type =
            getBonusType(row);


        if (bonusTypes[type]) {

            bonusTypes[type].push(row);

        }

    });


    // ---------------------------------------------
    // FIRST BONUS DESIGN IDs
    // ---------------------------------------------

    updateBonusBox(
        bonusTypes["REFER - BONUS"],
        "referBonusCount",
        "referBonusAmount"
    );


    updateBonusBox(
        bonusTypes["WELCOME BONUS"],
        "welcomeBonusCount",
        "welcomeBonusAmount"
    );


    updateBonusBox(
        bonusTypes["BONUS:-"],
        "normalBonusCount",
        "normalBonusAmount"
    );


    updateBonusBox(
        bonusTypes["WEEKLY CASHBACK BONUS"],
        "weeklyCashbackBonusCount",
        "weeklyCashbackBonusAmount"
    );


    // ---------------------------------------------
    // TOTAL
    // ---------------------------------------------

    const totalBonusAmount =
        rows.reduce(
            (sum, row) =>
                sum + getAmount(row),
            0
        );


    setText(
        "bonusCount",
        `${rows.length.toLocaleString()} entries`
    );


    setText(
        "bonusTotalAmount",
        "₹" + money(totalBonusAmount)
    );


    // ---------------------------------------------
    // BONUS TABLE
    // ---------------------------------------------

    const table =
        document.getElementById("bonusTable");


    if (!table) return;


    table.innerHTML =
        rows.length
            ? rows
                .map(row => `

                    <tr>

                        <td>
                            ${escapeHTML(
                                formatDate(
                                    getRowDate(row)
                                )
                            )}
                        </td>

                        <td>
                            ${escapeHTML(
                                getBonusType(row)
                            )}
                        </td>

                        <td class="credit">
                            ₹${money(
                                getAmount(row)
                            )}
                        </td>

                        <td>
                            ${escapeHTML(
                                getRemark(row)
                            )}
                        </td>

                        <td>
                            ${escapeHTML(
                                getUser(row)
                            )}
                        </td>

                        <td>
                            ${escapeHTML(
                                getUTR(row)
                            )}
                        </td>

                    </tr>

                `)
                .join("")
            : emptyRow(6);
}


// =============================================
// BONUS BOX UPDATE
// =============================================

function updateBonusBox(
    rows,
    countId,
    amountId
) {

    const count =
        rows.length;


    const amount =
        rows.reduce(
            (sum, row) =>
                sum + getAmount(row),
            0
        );


    setText(
        countId,
        count.toLocaleString()
    );


    setText(
        amountId,
        "₹" + money(amount)
    );
}


// =============================================
// REVERSAL TABLE
// =============================================

function renderReversalTable() {

    const rows =
        classifyRows(filteredData).reversal;


    setText(
        "reversalCountText",
        `${rows.length.toLocaleString()} entries`
    );

    const totalReversalAmount = rows.reduce(
        (sum, row) => sum + getAmount(row),
        0
    );

    setText(
        "reversalTotalAmount",
        "₹" + money(totalReversalAmount)
    );


    const table =
        document.getElementById("reversalTable");


    if (!table) return;


    table.innerHTML =
        rows.length
            ? rows.map(row => `

                <tr>

                    <td>
                        ${escapeHTML(
                            formatDate(
                                getRowDate(row)
                            )
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getEntryType(row)
                        )}
                    </td>

                    <td class="debit">
                        ₹${money(
                            getAmount(row)
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getRemark(row)
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getUser(row)
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getUTR(row)
                        )}
                    </td>

                </tr>

            `).join("")
            : emptyRow(6);
}


// =============================================
// ₹1 TABLE
// =============================================

function renderOneRupeeTable() {

    const rows =
        classifyRows(filteredData).oneRupee;


    setText(
        "oneRupeeCountText",
        `${rows.length.toLocaleString()} entries`
    );

    const totalOneRupeeAmount = rows.reduce(
        (sum, row) => sum + getAmount(row),
        0
    );

    setText(
        "oneRupeeTotalAmount",
        "₹" + money(totalOneRupeeAmount)
    );


    const table =
        document.getElementById("oneRupeeTable");


    if (!table) return;


    table.innerHTML =
        rows.length
            ? rows.map(row => `

                <tr>

                    <td>
                        ${escapeHTML(
                            formatDate(
                                getRowDate(row)
                            )
                        )}
                    </td>

                    <td class="credit">
                        ₹${money(
                            getAmount(row)
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getRemark(row)
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getUser(row)
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getUTR(row)
                        )}
                    </td>

                </tr>

            `).join("")
            : emptyRow(5);
}


// =============================================
// STATUS CHART
// =============================================

function renderStatusChart() {

    const canvas =
        document.getElementById("statusChart");


    if (!canvas) return;


    const counts = {};


    filteredData.forEach(row => {

        const status =
            getStatus(row);


        counts[status] =
            (counts[status] || 0) + 1;

    });


    destroyChart(statusChart);


    statusChart =
        new Chart(canvas, {

            type: "doughnut",

            data: {

                labels:
                    Object.keys(counts),

                datasets: [{

                    data:
                        Object.values(counts),

                    backgroundColor: [
                        "#22c55e",
                        "#f59e0b",
                        "#ef4444",
                        "#64748b"
                    ]

                }]

            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                plugins: {

                    legend: {
                        position: "bottom"
                    }

                }

            }

        });
}


// =============================================
// TYPE CHART
// =============================================

function renderTypeChart() {

    const canvas =
        document.getElementById("typeChart");


    if (!canvas) return;


    const counts = {};


    const normal =
        classifyRows(filteredData).normal;


    normal.forEach(row => {

        const type =
            getType(row);


        counts[type] =
            (counts[type] || 0) + 1;

    });


    destroyChart(typeChart);


    typeChart =
        new Chart(canvas, {

            type: "bar",

            data: {

                labels:
                    Object.keys(counts),

                datasets: [{

                    label:
                        "Valid Entries",

                    data:
                        Object.values(counts),

                    backgroundColor:
                        "#2563eb"

                }]

            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                scales: {

                    y: {
                        beginAtZero: true
                    }

                }

            }

        });
}


// =============================================
// CREDIT / DEBIT CHART
// =============================================

function renderCreditDebitChart() {

    const canvas =
        document.getElementById(
            "creditDebitChart"
        );


    if (!canvas) return;


    const normal =
        classifyRows(filteredData).normal;


    const credit =
        normal.reduce(
            (sum, row) =>
                sum + getCredit(row),
            0
        );


    const debit =
        normal.reduce(
            (sum, row) =>
                sum + getDebit(row),
            0
        );


    destroyChart(
        creditDebitChart
    );


    creditDebitChart =
        new Chart(canvas, {

            type: "bar",

            data: {

                labels: [
                    "Credit",
                    "Debit"
                ],

                datasets: [{

                    data: [
                        credit,
                        debit
                    ],

                    backgroundColor: [
                        "#16a34a",
                        "#dc2626"
                    ]

                }]

            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                plugins: {

                    legend: {
                        display: false
                    }

                },

                scales: {

                    y: {
                        beginAtZero: true
                    }

                }

            }

        });
}


// =============================================
// HOURLY CHART
// =============================================

function renderHourChart() {

    const canvas =
        document.getElementById("hourChart");


    if (!canvas) return;


    const hours =
        Array(24).fill(0);


    const normal =
        classifyRows(filteredData).normal;


    normal.forEach(row => {

        const date =
            parseDate(
                getRowDate(row)
            );


        if (date) {

            hours[
                date.getHours()
            ]++;

        }

    });


    destroyChart(hourChart);


    hourChart =
        new Chart(canvas, {

            type: "line",

            data: {

                labels:
                    hours.map(
                        (_, i) =>
                            `${String(i).padStart(2, "0")}:00`
                    ),

                datasets: [{

                    label:
                        "Valid Entries",

                    data:
                        hours,

                    borderColor:
                        "#7c3aed",

                    backgroundColor:
                        "rgba(124,58,237,.1)",

                    fill: true,

                    tension: 0.35

                }]

            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                scales: {

                    y: {
                        beginAtZero: true
                    }

                }

            }

        });
}


// =============================================
// CREDIT TABLE
// =============================================

function renderCreditTable() {

    const table =
        document.getElementById("creditTable");


    if (!table) return;


    const normal =
        classifyRows(filteredData).normal;


    const rows =
        [...normal]
            .filter(
                row =>
                    getCredit(row) > 0
            )
            .sort(
                (a, b) =>
                    getCredit(b) -
                    getCredit(a)
            )
            .slice(0, 10);


    table.innerHTML =
        rows.length
            ? rows.map(row => `

                <tr>

                    <td>
                        ${escapeHTML(
                            formatDate(
                                getRowDate(row)
                            )
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getSerial(row)
                        )}
                    </td>

                    <td class="credit">
                        ₹${money(
                            getCredit(row)
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getRemark(row)
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getUser(row)
                        )}
                    </td>

                </tr>

            `).join("")
            : emptyRow(5);
}


// =============================================
// DEBIT TABLE
// =============================================

function renderDebitTable() {

    const table =
        document.getElementById("debitTable");


    if (!table) return;


    const normal =
        classifyRows(filteredData).normal;


    const rows =
        [...normal]
            .filter(
                row =>
                    getDebit(row) > 0
            )
            .sort(
                (a, b) =>
                    getDebit(b) -
                    getDebit(a)
            )
            .slice(0, 10);


    table.innerHTML =
        rows.length
            ? rows.map(row => `

                <tr>

                    <td>
                        ${escapeHTML(
                            formatDate(
                                getRowDate(row)
                            )
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getSerial(row)
                        )}
                    </td>

                    <td class="debit">
                        ₹${money(
                            getDebit(row)
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getRemark(row)
                        )}
                    </td>

                    <td>
                        ${escapeHTML(
                            getUser(row)
                        )}
                    </td>

                </tr>

            `).join("")
            : emptyRow(5);
}


// =============================================
// SERIAL NUMBER
// =============================================

function getSerial(row) {

    return String(
        getColumn(row, [
            "Sr No",
            "Serial No",
            "S No",
            "S.No",
            "Sr.No"
        ]) || ""
    );
}


// =============================================
// HOLD / REJECTED
// =============================================

function renderProblemTable() {

    const table =
        document.getElementById("problemTable");


    if (!table) return;


    const rows =
        filteredData.filter(row => {

            const status =
                getStatus(row);


            return (
                status === "Hold" ||
                status === "Rejected"
            );

        });


    table.innerHTML =
        rows.length
            ? rows.map(row => {

                const status =
                    getStatus(row);


                const cls =
                    status === "Rejected"
                        ? "status-rejected"
                        : "status-hold";


                return `
                    <tr>

                        <td>
                            ${escapeHTML(
                                formatDate(
                                    getRowDate(row)
                                )
                            )}
                        </td>

                        <td>
                            ${escapeHTML(
                                getSerial(row)
                            )}
                        </td>

                        <td>
                            ₹${money(
                                getAmount(row)
                            )}
                        </td>

                        <td>

                            <span
                                class="status ${cls}"
                            >
                                ${escapeHTML(status)}
                            </span>

                        </td>

                        <td>
                            ${escapeHTML(
                                getRemark(row)
                            )}
                        </td>

                        <td>
                            ${escapeHTML(
                                getUser(row)
                            )}
                        </td>

                    </tr>
                `;

            }).join("")
            : emptyRow(6);
}


// =============================================
// MAIN TABLE
// =============================================

function renderMainTable() {

    const table =
        document.getElementById("mainTable");


    const head =
        document.getElementById("mainHead");


    const count =
        document.getElementById("rowCount");


    if (!table || !head) return;


    const normal =
        classifyRows(filteredData).normal;


    if (!normal.length) {

        head.innerHTML = "";


        table.innerHTML =
            emptyRow(1);


        if (count) {

            count.textContent =
                "0 valid rows";

        }

        return;
    }


    const columns =
        Object.keys(normal[0]);


    head.innerHTML =
        `
        <tr>
            ${columns.map(
                col =>
                    `<th>
                        ${escapeHTML(col)}
                    </th>`
            ).join("")}
        </tr>
        `;


    table.innerHTML =
        normal.map(row => {

            return `
                <tr>

                    ${columns.map(
                        col =>
                            `<td>
                                ${escapeHTML(
                                    row[col]
                                )}
                            </td>`
                    ).join("")}

                </tr>
            `;

        }).join("");


    if (count) {

        count.textContent =
            `${normal.length.toLocaleString()} valid rows`;

    }
}


// =============================================
// EMPTY ROW
// =============================================

function emptyRow(colspan) {

    return `
        <tr>
            <td
                colspan="${colspan}"
                class="empty-message"
            >
                No entries found
            </td>
        </tr>
    `;
}


// =============================================
// CHART DESTROY
// =============================================

function destroyChart(chart) {

    if (chart) {
        chart.destroy();
    }
}


// =============================================
// TEXT
// =============================================

function setText(id, value) {

    const element =
        document.getElementById(id);


    if (element) {
        element.textContent = value;
    }
}


// =============================================
// MONEY
// =============================================

function money(value) {

    return Number(value || 0)
        .toLocaleString(
            "en-IN",
            {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            }
        );
}


// =============================================
// HTML SECURITY
// =============================================

function escapeHTML(value) {

    return String(value ?? "")
        .replace(
            /[&<>"']/g,
            char => ({
                "&": "&amp;",
                "<": "&lt;",
                ">": "&gt;",
                '"': "&quot;",
                "'": "&#039;"
            })[char]
        );
}


// =============================================
// DOWNLOAD CSV
// =============================================

const downloadBtn =
    document.getElementById("downloadBtn");


if (downloadBtn) {

    downloadBtn.addEventListener(
        "click",
        downloadCSV
    );

}


function downloadCSV() {

    const normal =
        classifyRows(filteredData).normal;


    if (!normal.length) {

        alert(
            "No valid data to download."
        );

        return;
    }


    const worksheet =
        XLSX.utils.json_to_sheet(
            normal
        );


    const csv =
        XLSX.utils.sheet_to_csv(
            worksheet
        );


    const blob =
        new Blob(
            [csv],
            {
                type:
                    "text/csv;charset=utf-8;"
            }
        );


    const url =
        URL.createObjectURL(blob);


    const link =
        document.createElement("a");


    link.href = url;


    link.download =
        "valid_transactions.csv";


    document.body.appendChild(link);


    link.click();


    document.body.removeChild(link);


    URL.revokeObjectURL(url);
}
