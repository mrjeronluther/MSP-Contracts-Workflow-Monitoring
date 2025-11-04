// ='================================================================
// SCRIPT CONFIGURATION
// =================================================================
const SHEET_ID = "1Aa2O2XVinhL7x2zaBlCgpAcv7IcLAcqeUgGVfRNidzw";
const RECIPIENT_SHEET_ID = "1Aa2O2XVinhL7x2zaBlCgpAcv7IcLAcqeUgGVfRNidzw";
const RECIPIENT_SHEET_ID1 = "1qheN_KURc-sOKSngpzVxLvfkkc8StzGv-1gMvGJZdsc";
const EXTERNAL_DATA_SHEET_ID = "1Aa2O2XVinhL7x2zaBlCgpAcv7IcLAcqeUgGVfRNidzw";

const TAB_NAME = "LE UNIFICADO";
const RECIPIENT_TAB_NAME = "PropEmailAdd";
const RECIPIENT_TAB_NAME1 = "dvSupplierPayee";
const EXTERNAL_DATA_TAB_NAME = "RefNoSeries";
const LOG_TAB_NAME = "EmailLogs"; // New tab for logging



const isDebugging = true;
const FALLBACK_RECIPIENTS = "jecastro@megaworld-lifestyle.com";

// --- CACHE VARIABLES ---
let propertyRecipientCache = null;
let supplierRecipientCache = null;
let externalDataCache = null; // Cache for all external data from RefNoSeries

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Send Email')
    .addItem('Run', 'checkSheetAndSendEmails')
    .addToUi();
 
}

// =================================================================
// EMAIL LOGGING FUNCTION
// =================================================================

/**
 * Logs the status of an email sent to a dedicated "EmailLogs" tab.
 * Creates the tab and header row if they don't exist.
 * @param {string} recipient The email address(es) the email was sent to.
 * @param {string} subject The subject line of the email.
 * @param {string} status The status of the email ('SUCCESS' or 'FAILURE').
 * @param {string} [errorMessage] Optional error message if the status is 'FAILURE'.
 */
function logEmailStatus(recipient, subject, status, errorMessage) {
    try {
        const ss = SpreadsheetApp.openById(SHEET_ID);
        let logSheet = ss.getSheetByName(LOG_TAB_NAME);

        // If the log sheet doesn't exist, create it and add headers.
        if (!logSheet) {
            logSheet = ss.insertSheet(LOG_TAB_NAME);
            logSheet.appendRow(["Timestamp", "Recipient", "Subject", "Status", "Error Message"]);
            logSheet.getRange("A1:E1").setFontWeight("bold"); // Make header bold
        }

        const timestamp = new Date();
        logSheet.appendRow([timestamp, recipient, subject, status, errorMessage || ""]);
    } catch (e) {
        Logger.log(`Failed to log email status. Error: ${e.message}`);
    }
}


// =================================================================
// REUSABLE DATA & RECIPIENT HELPER FUNCTIONS
// =================================================================

/**
 * REFACTORED: Fetches all data from the "RefNoSeries" sheet and caches it.
 * This function runs ONLY ONCE per execution to maximize efficiency.
 * @param {string} lookupValue The value (MSP Ref#) to look up in the cache.
 * @returns {object} An object containing all data for the given key, or a default object if not found.
 */
function getExternalData(lookupValue) {
    if (externalDataCache === null) {
        if (isDebugging) Logger.log('Initializing EXTERNAL DATA cache from RefNoSeries...');
        externalDataCache = new Map();
        try {
            const sheet = SpreadsheetApp.openById(EXTERNAL_DATA_SHEET_ID).getSheetByName(EXTERNAL_DATA_TAB_NAME);
            if (sheet) {
                const data = sheet.getDataRange().getValues();
                // Skip header row (index 0)
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const refKey = String(row[0]).trim(); // Key is in Column A
                    if (refKey) {
                        externalDataCache.set(refKey, {
                            series: row[1] || "",
                            supplier: row[7] || "Not Found",
                            sector: row[12] || "Not Found"
                        });
                    }
                }
                if (isDebugging) Logger.log('External data cache populated successfully.');
            } else {
                if (isDebugging) Logger.log(`Warning: External data sheet "${EXTERNAL_DATA_TAB_NAME}" not found.`);
            }
        } catch (e) {
            Logger.log(`Error reading external data sheet: ${e.message}`);
        }
    }

    const matchedData = externalDataCache.get(String(lookupValue || '').trim());

    if (matchedData) {
        return matchedData;
    } else {
        if (isDebugging) Logger.log(`  [ExternalData] Data for "${lookupValue}" not found in cache.`);
        return {
            series: "NOT FOUND",
            supplier: "NOT FOUND",
            sector: "NOT FOUND"
        };
    }
}

/**
 * Gets recipients from the Property sheet ("Sheet3").
 * Caches data in 'propertyRecipientCache' for efficiency.
 * @param {string} lookupValue The property/township name.
 * @returns {string} A comma-separated string of email addresses.
 */
function getPropertyRecipients(lookupValue) {
    if (propertyRecipientCache === null) {
        if (isDebugging) Logger.log('Initializing PROPERTY recipient cache...');
        propertyRecipientCache = new Map();
        try {
            const sheet = SpreadsheetApp.openById(RECIPIENT_SHEET_ID).getSheetByName(RECIPIENT_TAB_NAME);
            if (sheet) {
                const data = sheet.getDataRange().getValues();
                const headers = data[0];
                headers.forEach((header, colIndex) => {
                    if (header) {
                        const emails = [];
                        for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
                            const email = data[rowIndex][colIndex];
                            if (email && String(email).includes('@')) {
                                emails.push(String(email).trim());
                            }
                        }
                        propertyRecipientCache.set(String(header).trim(), emails);
                    }
                });
                if (isDebugging) Logger.log('Property recipient cache populated successfully.');
            } else {
                if (isDebugging) Logger.log(`Warning: Property sheet "${RECIPIENT_TAB_NAME}" not found.`);
            }
        } catch (e) {
            Logger.log(`Error reading property recipient sheet: ${e.message}`);
        }
    }
    const matchedEmails = propertyRecipientCache.get(String(lookupValue || '').trim());
    if (matchedEmails && matchedEmails.length > 0) {
        const recipientString = matchedEmails.join(',');
        if (isDebugging) Logger.log(`  [PropertyRecipients] Found for "${lookupValue}": ${recipientString}`);
        return recipientString;
    } else {
        if (isDebugging) Logger.log(`  [PropertyRecipients] Not found for "${lookupValue}". Using fallback.`);
        return FALLBACK_RECIPIENTS;
    }
}

/**
 * REVISED: Gets recipients from the Supplier sheet ("dvSuppliers").
 * This version reads a vertical layout: looks up a name in Column C and gets the email(s) from Column J.
 * Caches data in the 'supplierRecipientCache' variable for efficiency.
 * @param {string} lookupValue The supplier name to search for in Column C.
 * @returns {string} A comma-separated string of email addresses.
 */
function getSupplierRecipients(lookupValue) {
    // Step 1: Populate the cache if it's empty. This runs only once per execution.
    if (supplierRecipientCache === null) {
        if (isDebugging) Logger.log('Initializing SUPPLIER recipient cache from vertical list...');
        supplierRecipientCache = new Map(); // Use a Map for efficient 'key -> value' storage.
        try {
            const sheet = SpreadsheetApp.openById(RECIPIENT_SHEET_ID1).getSheetByName(RECIPIENT_TAB_NAME1);
            if (sheet) {
                const data = sheet.getDataRange().getValues();

                // Loop through each row of the sheet, starting from the first row (index 0).
                for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
                    const row = data[rowIndex];
                    const supplierName = row[2]; // Column C (index 2) is the key.
                    const emailCell = row[9]; // Column J (index 9) contains the emails.

                    // Proceed only if we have both a supplier name and something in the email cell.
                    if (supplierName && emailCell) {

                        // --- Handle multiple emails in a single cell ---
                        // 1. Convert cell to string and split by comma, ANY whitespace, or semicolon.
                        // 2. Map over the results to trim whitespace from each potential email.
                        // 3. Filter the array to keep only non-blank strings that contain "@".
                        const validEmails = String(emailCell)
                            .split(/[,\s;]+/) // <-- THE FIX IS HERE
                            .map(email => email.trim())
                            .filter(email => email && email.includes('@'));

                        // If we found any valid emails, add them to the cache.
                        if (validEmails.length > 0) {
                            // Join the clean emails into a single comma-separated string.
                            const emailString = validEmails.join(',');
                            // Set the cache: "Supplier Name" -> "email1@a.com,email2@b.com"
                            supplierRecipientCache.set(String(supplierName).trim(), emailString);
                        }
                    }
                }
                if (isDebugging) Logger.log('Supplier recipient cache populated successfully.');

            } else {
                if (isDebugging) Logger.log(`Warning: Supplier sheet "${RECIPIENT_TAB_NAME1}" not found.`);
            }
        } catch (e) {
            Logger.log(`Error reading supplier recipient sheet: ${e.message}`);
        }
    }

    // Step 2: Look up the value (e.g., the specific Supplier) in the now-populated cache.
    const matchedEmails = supplierRecipientCache.get(String(lookupValue || "").trim());

    // Step 3: Return the result.
    if (matchedEmails) { // The cache now stores the final, clean string.
        if (isDebugging) Logger.log(`  [SupplierRecipients] Found for "${lookupValue}": ${matchedEmails}`);
        return matchedEmails;
    } else {
        if (isDebugging) Logger.log(`  [SupplierRecipients] Not found for "${lookupValue}". Using fallback.`);
        return FALLBACK_RECIPIENTS;
    }
}


/**
 * Main function to be triggered to check the sheet and send emails.
 */
function checkSheetAndSendEmails() {
    try {
        const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_NAME);
        if (!sheet) {
            Logger.log(`Execution failed: Could not find the sheet. Check Sheet ID and Tab Name.`);
            return;
        }
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        const formulas = dataRange.getFormulas();
        const headers = values[7]; // Assume header is on row 8

        if (isDebugging) {
            Logger.log(`Starting check on "${TAB_NAME}". Total data rows to process: ${values.length - 8}`);
        }

        for (let i = 8; i < values.length; i++) {
            const rowData = values[i];
            const rowFormulas = formulas[i];
            const rowNumber = i + 1;

            if (rowData.every((cell) => cell === "")) continue;

            if (isDebugging) Logger.log(`\n--- Processing Row #${rowNumber} ---`);

            const b_val = rowData[1];
            const f_val = rowData[5];
            const f_formula = rowFormulas[5];
            const h_val = rowData[7];
            const i_val = rowData[8];
            const j_val = rowData[9];
            const l_val = rowData[11];
            const m_val = rowData[12];
            const p_val = rowData[15];
            const r_val = rowData[17];

            const statusIsDone = String(p_val).trim().toLowerCase() === "done";
            if (statusIsDone) {
                const m_date = new Date(m_val);
                if (m_val && !isNaN(m_date.getTime())) {
                    const today = new Date();
                    today.setHours(0, 0, 0, 0);
                    const twoDaysAgo = new Date(today.getTime() - (2 * 24 * 60 * 60 * 1000));
                    if (m_date < twoDaysAgo) {
                        if (isDebugging) Logger.log(`ROW SKIPPED: Status is "DONE" and date in Col M is more than 2 days old.`);
                        continue;
                    }
                }
            }

            if (!isConsideredNumber(b_val) && isConsideredNotBlank(b_val) && isDateAndToday(h_val)) {
                if (isDebugging) Logger.log(`Condition 1 MET.`);
                sendEmailForCondition1(rowNumber, rowData, headers);
            }
            if (isConsideredNumber(b_val) && isDateAndToday(h_val)) {
                if (isDebugging) Logger.log(`Condition 2 MET.`);
                sendEmailForCondition2(rowNumber, rowData, headers);
            }
            if (isConsideredNumber(b_val) && isValidLink(f_val, f_formula) && isDateAndToday(i_val)) {
                if (isDebugging) Logger.log(`Condition 3 MET.`);
                sendEmailForCondition3(rowNumber, rowData, headers);
            }
            if (r_val !== true && isConsideredNumber(b_val) && isValidLink(f_val, f_formula) && isDateAndAfterToday(i_val) && isNotValidDateOrBlank(j_val)) {
                if (isDebugging) Logger.log(`Condition 4 MET.`);
                sendEmailForCondition4(rowNumber, rowData, headers);
            }
            if (isConsideredNumber(b_val) && isValidLink(f_val, f_formula) && isDateAndAfterToday(i_val) && isDateAndToday(j_val)) {
                if (isDebugging) Logger.log(`Condition 5 MET.`);
                sendEmailForCondition5(rowNumber, rowData, headers);
            }
            if (isConsideredNumber(b_val) && isValidLink(f_val, f_formula) && isDateAndToday(l_val)) {
                if (isDebugging) Logger.log(`Condition 6 MET.`);
                sendEmailForCondition6(rowNumber, rowData, headers);
            }
            if (isConsideredNumber(b_val) && isValidLink(f_val, f_formula) && isDateAndAfterToday(l_val) && isNotValidDateOrBlank(m_val)) {
                if (isDebugging) Logger.log(`Condition 7 MET.`);
                sendEmailForCondition7(rowNumber, rowData, headers);
            }
            if (isConsideredNumber(b_val) && isValidLink(f_val, f_formula) && isDateAndAfterToday(l_val) && isDateAndToday(m_val)) {
                if (isDebugging) Logger.log(`Condition 8 MET.`);
                sendEmailForCondition8(rowNumber, rowData, headers);
            }
        }
        if (isDebugging) Logger.log(`\n--- Check Complete ---`);
    } catch (e) {
        Logger.log(`An error occurred: ${e.message}\n${e.stack}`);
    }
}


// =================================================================
// ROBUST DATA CHECKING FUNCTIONS
// =================================================================

function parseAndNormalizeDate(value) {
    if (!value) return null;
    const date = new Date(value);
    if (isNaN(date.getTime())) return null;
    date.setHours(0, 0, 0, 0);
    return date;
}

function isDateAndToday(value) {
    const date = parseAndNormalizeDate(value);
    if (!date) return false;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return date.getTime() === today.getTime();
}

function isDateAndAfterToday(value) {
    const date = parseAndNormalizeDate(value);
    if (!date) return false;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return date.getTime() > today.getTime();
}

function isNotValidDateOrBlank(value) {
    const trimmedValue = String(value || '').trim();
    if (trimmedValue === '' || trimmedValue === '-') return true;
    return isNaN(new Date(value).getTime());
}

function isConsideredNumber(value) {
    return value !== null && value !== '' && !isNaN(Number(value)) && isFinite(value);
}

function isValidLink(cellValue, formulaValue) {
    const isFormulaLink = formulaValue && formulaValue.toLowerCase().startsWith("=hyperlink(");
    const isTextLink = typeof cellValue === 'string' && (cellValue.toLowerCase().includes("http://") || cellValue.toLowerCase().includes("https://"));
    return isFormulaLink || isTextLink;
}

function isConsideredNotBlank(value) {
    return value !== null && value !== undefined && String(value).trim() !== '';
}

// =================================================================
// EMAIL SENDING FUNCTIONS (Conditions 1-8)
// =================================================================

function createSubjectForCondition1(rowData) {
    const u_val = rowData[20] || 'N/A';
    const v_val = rowData[21] || 'N/A';
    const w_val = rowData[22] || 'N/A';
    const x_val = rowData[23] || 'N/A';
    const extracted_u = u_val.split(' - ')[0].trim();
    const extracted_x = x_val.split(' - ')[0].trim();
    return `(${extracted_u}): ${v_val} & ${w_val} (${extracted_x})`;
}

// --- CONDITIONS 1-5: Send ONLY to Property ---
function sendEmailForCondition1(rowNumber, rowData, headers) { const recipients = getPropertyRecipients(rowData[20]); const subject = "CANNOT PROCESS MSP CONTRACT " + createSubjectForCondition1(rowData); if (!recipients || (recipients ===
FALLBACK_RECIPIENTS && !rowData[20])) { logEmailStatus(FALLBACK_RECIPIENTS, subject, "FAILURE", "Recipient lookup was blank."); return; } const c_val = rowData[2] || 'No Remarks'; const extracted_c = c_val.split(' - ')[0].trim(); const body
= `
<p>Hi Team,</p>
<p>Pertaining to the request for a formal agreement for a MSP that you have filed (see details below), we regret to inform you that the REQUEST CANNOT BE PROCESSED because: <b> ${extracted_c} </b></p>
<p><b>Details:</b></p>
<ul>
    <li><b>${headers[18]}:</b> ${rowData[18]}</li>
    <li><b>${headers[19]}:</b> ${rowData[19]}</li>
    <li><b>${headers[20]}:</b> ${rowData[20]}</li>
    <li><b>Service Provider:</b> ${externalData.supplier}</li>
    <li><b>${headers[22]}:</b> ${rowData[22]}</li>
    <li><b>${headers[23]}:</b> ${rowData[23]}</li>
    <li><b>${headers[24]}:</b> ${rowData[24]}</li>
    <li><b>${headers[3]}:</b> ${rowData[3]}</li>
    <li><b>${headers[4]}:</b> ${rowData[4]}</li>
</ul>
<p>Please re-file your request when the reason for concern or issue is addressed. Thank you!</p>
`; try { MailApp.sendEmail(recipients, subject, "", { htmlBody: body }); logEmailStatus(recipients, subject, "SUCCESS"); } catch (e) { logEmailStatus(recipients, subject, "FAILURE", e.message); } }


function sendEmailForCondition2(rowNumber, rowData, headers) { const recipients = getPropertyRecipients(rowData[20]); const b_val = rowData[1]; const externalData = getExternalData(b_val); const subject = `ON PROCESS: MSP ref# ${b_val} |
${externalData.series}`; if (!recipients || (recipients === FALLBACK_RECIPIENTS && !rowData[20])) { logEmailStatus(FALLBACK_RECIPIENTS, subject, "FAILURE", "Recipient lookup was blank."); return; } const u_val = rowData[20] || 'N/A'; const
extracted_u = u_val.split(' - ')[0].trim(); const body = `
<p>Hi Team,</p>
<p>I wanted to inform you that your request is now in process and currently under review by the approver. Please find the details below for your reference:</p>
<p><b>Item Details:</b></p>
<ul>
    <li><b>${headers[20]}:</b> ${extracted_u}</li>
    <li><b>Service Provider:</b> ${externalData.supplier}</li>
    <li><b>${headers[22]}:</b> ${rowData[22]}</li>
    <li><b>Sector:</b> ${externalData.sector}</li>
    <li><b>${headers[1]}:</b> ${b_val}</li>
    <li><b>${headers[2]}:</b> ${rowData[2]}</li>
</ul>
<p>I will keep you updated on any further progress. Thank you for your patience and cooperation.</p>
`; try { MailApp.sendEmail(recipients, subject, "", { htmlBody: body }); logEmailStatus(recipients, subject, "SUCCESS"); } catch (e) { logEmailStatus(recipients, subject, "FAILURE", e.message); } }


function sendEmailForCondition3(rowNumber, rowData, headers) { const recipients = getPropertyRecipients(rowData[20]); const b_val = rowData[1]; const externalData = getExternalData(b_val); const subject = `ON PROCESS: MSP ref# ${b_val} |
${externalData.series}`; if (!recipients || (recipients === FALLBACK_RECIPIENTS && !rowData[20])) { logEmailStatus(FALLBACK_RECIPIENTS, subject, "FAILURE", "Recipient lookup was blank."); return; } let pdfDraft = "Not Found"; try { const
refSheet1 = SpreadsheetApp.openById(EXTERNAL_DATA_SHEET_ID).getSheetByName("MLC ONLY MONITORING"); if (refSheet1) { const refData = refSheet1.getDataRange().getValues(); for (let i = 0; i < refData.length; i++) { if
(String(refData[i][3]).trim() == String(b_val).trim()) { pdfDraft = refData[i][5]; break; } } } } catch (e) { Logger.log("Error fetching PDF Draft: " + e.message); pdfDraft = "ERROR"; } const body = `
<p>Hi Team,</p>
<p>Pertaining to the request for a formal agreement for a MSP that you have filed with details below, we are pleased to send to you the draft agreement/appointment for your perusal.</p>
<p><b>Details:</b></p>
<ul>
    <li><b>${headers[18]}:</b> ${rowData[18]}</li>
    <li><b>${headers[19]}:</b> ${rowData[19]}</li>
    <li><b>Service Provider:</b> ${externalData.supplier}</li>
    <li><b>${headers[22]}:</b> ${rowData[22]}</li>
    <li><b>${headers[23]}:</b> ${rowData[23]}</li>
    <li><b>${headers[24]}:</b> ${rowData[24]}</li>
    <li><b>${headers[3]}:</b> ${rowData[3]}</li>
    <li><b>${headers[4]}:</b> ${rowData[4]}</li>
    <li><b>${headers[1]}:</b> ${b_val}</li>
    <li><b>Sector:</b> ${externalData.sector}</li>
    <li><b>${headers[2]}:</b> ${rowData[2]}</li>
    <li><b>Link to the PDF draft : SFC:</b> ${rowData[5]}</li>
    <li><b>Link to the PDF draft : MLC:</b> ${pdfDraft}</li>
</ul>
<p>Please return to us the contract, 4 copies only for LOA/RA and 2 copies for MLC For scanned signed contract send in this same email trail Kindly comply as soon as possible to avoid delays.</p>
`; try { MailApp.sendEmail(recipients, subject, "", { htmlBody: body }); logEmailStatus(recipients, subject, "SUCCESS"); } catch (e) { logEmailStatus(recipients, subject, "FAILURE", e.message); } }



function sendEmailForCondition4(rowNumber, rowData, headers) {
    const recipients = getPropertyRecipients(rowData[20]);
    const b_val = rowData[1];
    const externalData = getExternalData(b_val);
    const subject = `ON PROCESS: MSP ref# ${b_val} | ${externalData.series}`;
    if (!recipients || (recipients === FALLBACK_RECIPIENTS && !rowData[20])) {
        logEmailStatus(FALLBACK_RECIPIENTS, subject, "FAILURE", "Recipient lookup was blank.");
        return;
    }
    const contractFinderUrl = "https://script.google.com/a/macros/megaworld-lifestyle.com/s/AKfycbzcFIF5i_mHXNw3PKrPAvKpLr7EjVv2L3CFFlwMK5IFNu7nv6F0vwYa_4PqlGUCdPBd/exec";
    const body = `
<p>Hello, ${rowData[20]} Team.</p>
<p>This is a reminder that the document previously sent is still pending the GM's signature, as per initial instructions.</p>
<ul>
    <li><b>${headers[16]}:</b> ${rowData[16]}</li>
</ul>
<p>Please return:</p>
<ul>
    <li>4 copies for LOA/RA</li>
    <li>2 copies for MLC</li>
</ul>
<p>Signed scanned copies can be sent via this email thread or through the <a href="${contractFinderUrl}">Contract Finder</a>.</p>
<p>You may also send the hard copies via our messenger.</p>
<p>Thank you for your attention and cooperation.</p>
`;
    try {
        MailApp.sendEmail(recipients, subject, "", {
            htmlBody: body
        });
        logEmailStatus(recipients, subject, "SUCCESS");
    } catch (e) {
        logEmailStatus(recipients, subject, "FAILURE", e.message);
    }
}

function sendEmailForCondition5(rowNumber, rowData, headers) {
    const recipients = getPropertyRecipients(rowData[20]);
    const b_val = rowData[1];
    const externalData = getExternalData(b_val);
    const subject = `ON PROCESS: MSP ref# ${b_val} | ${externalData.series}`;
    if (!recipients || (recipients === FALLBACK_RECIPIENTS && !rowData[20])) {
        logEmailStatus(FALLBACK_RECIPIENTS, subject, "FAILURE", "Recipient lookup was blank.");
        return;
    }
    const body = `<p>Dear, ${rowData[20]} Team.</p><p>The GM/CGM-signed contract has been received. It is now routing for signature with the COG and MOSD teams. We will update you once routing is complete or if further action is needed.</p><p>Thank you.</p>`;
    try {
        MailApp.sendEmail(recipients, subject, "", {
            htmlBody: body
        });
        logEmailStatus(recipients, subject, "SUCCESS");
    } catch (e) {
        logEmailStatus(recipients, subject, "FAILURE", e.message);
    }
}

// --- CONDITIONS 6-8: Send to MULTIPLE groups ---
function sendEmailForCondition6(rowNumber, rowData, headers) {
    const b_val = rowData[1];
    const externalData = getExternalData(b_val);
    const subject = `For Signing and Notary: MSP ref# ${b_val} | ${externalData.series}`;
    const propertyRecipients = getPropertyRecipients(rowData[20]);
    const supplierRecipients = getSupplierRecipients(externalData.supplier);
    const pcuRecipients = getPropertyRecipients("PCU");
    const allRecipients = [propertyRecipients, supplierRecipients, pcuRecipients];
    const finalRecipients = [...new Set(allRecipients.join(',').split(',').filter(e => e && e.includes('@')))].join(',');

    if (!finalRecipients) {
        logEmailStatus("N/A", subject, "FAILURE", "No valid recipients found after combining lists.");
        return;
    }

    const u_val = rowData[20] || 'N/A';
    const extracted_u = u_val.split(" - ")[0].trim();
    const body = `
<p>Hi Team,</p>
<p>We would like to inform you that the following contracts are now ready for pickup at your earliest convenience:</p>
<p><b>Details:</b></p>
<ul>
    <li><b>${headers[20]}:</b> ${extracted_u}</li>
    <li><b>Service Provider:</b> ${externalData.supplier}</li>
    <li><b>${headers[22]}:</b> ${rowData[22]}</li>
    <li><b>${headers[23]}:</b> ${rowData[23]}</li>
    <li><b>Sector:</b> ${externalData.sector}</li>
    <li><b>${headers[1]}:</b> ${b_val}</li>
    <li><b>${headers[2]}:</b> ${rowData[2]}</li>
</ul>
<p>To ensure a smooth and timely billing process, we kindly request that these documents be collected within the next three (3) working days. Timely pickup will help avoid any delays in contract processing.</p>
<p><b>Important Reminders:</b></p>
<ul>
    <li>Return 3 copies of notarized contracts for LOA, RA and Extension. 1 copy for Main Legal Contract.</li>
    <li>Please notify me at least one (1) day before pickup in case I am on leave.</li>
    <li>
        <b>Pickup Schedule:</b> Monday to Friday, 9:00 AM to 5:00 PM<br />
        <i>Strictly no pickup during lunch break (12:00 NN – 1:00 PM)</i>
    </li>
    <li>
        <b>Pickup Location:</b><br />
        2nd Floor, IBM Plaza Bldg<br />
        8 Eastwood Avenue, Quezon City, 1110 Metro Manila
    </li>
</ul>
<p>Thank you for your cooperation and prompt action.</p>
`;
    try {
        MailApp.sendEmail(finalRecipients, subject, "", {
            htmlBody: body
        });
        logEmailStatus(finalRecipients, subject, "SUCCESS");
    } catch (e) {
        logEmailStatus(finalRecipients, subject, "FAILURE", e.message);
    }
}





function sendEmailForCondition7(rowNumber, rowData, headers) {
    const b_val = rowData[1];
    const externalData = getExternalData(b_val);
    const subject = `For Signing and Notary: MSP ref# ${b_val} | ${externalData.series}`;
    const propertyRecipients = getPropertyRecipients(rowData[20]);
    const supplierRecipients = getSupplierRecipients(externalData.supplier);
    const pcuRecipients = getPropertyRecipients("PCU");
    const allRecipients = [propertyRecipients, supplierRecipients, pcuRecipients];
    const finalRecipients = [...new Set(allRecipients.join(',').split(',').filter(e => e && e.includes('@')))].join(',');

    if (!finalRecipients) {
        logEmailStatus("N/A", subject, "FAILURE", "No valid recipients found after combining lists.");
        return;
    }

    const body = `
<p>Hi ${externalData.supplier},</p>
<p>This is a follow-up on the notarized contracts to help avoid delays in billing and processing.</p>



<p><b>Reminders:</b></p>
<ul>
    <li>- ${headers[16]}: <b>${rowData[16]}</b></li>
    <li>- Return 3 copies for LOA, RA, and Extension; 1 copy for the Main Legal Contract</li>
    <li>- Notify me at least 1 day before pickup in case of leave</li>
    <li>- Pickup Schedule: Monday to Friday, 9:00 AM - 5:00 PM (No pickups during lunch break: 12:00 NN - 1:00 PM)</li>
</ul>
<p><b>Pickup Location:</b></p>
<ul>
    <li>- 2nd Floor, IBM Plaza Bldg 8 Eastwood Avenue, Quezon City, 1110 Metro Manila</li>
</ul>
<p>Thank you for your prompt attention.</p>
`;
    try {
        MailApp.sendEmail(finalRecipients, subject, "", {
            htmlBody: body
        });
        logEmailStatus(finalRecipients, subject, "SUCCESS");
    } catch (e) {
        logEmailStatus(finalRecipients, subject, "FAILURE", e.message);
    }
}


function sendEmailForCondition8(rowNumber, rowData, headers) {
    const b_val = rowData[1];
    const externalData = getExternalData(b_val);
    const subject = `Receipt of notarize contract: MSP ref# ${b_val} | ${externalData.series}`;
    const propertyRecipients = getPropertyRecipients(rowData[20]);
    const supplierRecipients = getSupplierRecipients(externalData.supplier);
    const pcuRecipients = getPropertyRecipients("PCU");
    const tkRecipients = getPropertyRecipients("TK");
    const allRecipients = [propertyRecipients, supplierRecipients, pcuRecipients, tkRecipients];
    const finalRecipients = [...new Set(allRecipients.join(',').split(',').filter(e => e && e.includes('@')))].join(',');

    if (!finalRecipients) {
        logEmailStatus("N/A", subject, "FAILURE", "No valid recipients found after combining lists.");
        return;
    }

    const body = `<p>Hello ${rowData[20]} Team,</p><p>To streamline the routing process, a scanned copy of the property contract is attached. The original will be sent via messenger or courier. Please refer to the scanned copy for your reference. Direct any concerns in a separate email thread.</p><p>Thank you for your cooperation.</p>`;
    try {
        MailApp.sendEmail(finalRecipients, subject, "", {
            htmlBody: body
        });
        logEmailStatus(finalRecipients, subject, "SUCCESS");
    } catch (e) {
        logEmailStatus(finalRecipients, subject, "FAILURE", e.message);
    }
}
