// ===============================================================
// --- GLOBAL CONFIGURATION VARIABLES ---
// ===============================================================

const SPREADSHEET_ID = "1Aa2O2XVinhL7x2zaBlCgpAcv7IcLAcqeUgGVfRNidzw"; 
const SHEET_NAME = "LE UNIFICADO";
const HEADER_ROW_NUMBER = 9;
const COLUMNS_TO_INCLUDE = [
    "Ref#",                          //DONE
    "Township or Property or Group", //DONE
    "Ball with",                     //DONE
    "Days pending with",             //DONE
    "Date Received Request",         //DONE
    "Kind of Contract",              //DONE
    "NOA Ref#",
    "NOA Copy",                      //DONE
    "Notarized Contact Copy",
    "MSP Contract Ref#",
    "PDF draft sent to prop",          //DONE
    "MSP Contract Draft", 
    "Date Drafted/ Rejected",         //DONE
    "Date emailed to GM",              //DONE
    "Date Routed to MOSD-COG  Reviewers",
    "Date Informed Agency  to Pick-up", //DONE
    "Date Return of Notarized Copy",   //DONE
    "Service Provider",  //DONE
    "Payor Company", //DONE
    "Type of Service", //DONE
    "Section", //DONE
    "Start Date",  //DONE
    "End Date", //DONE
    "Remarks", //DONE
    "Submit Contract signed by GM/CGM",  //DONE include in the PROPERTY tab only
    
];

const SEARCH_COLUMN = "Ref#";

const RBG_TK_SHEET_ID = "1m7bOgXL4UJHUd0euaYMguAakhhuElRPxBqZ6R_GTRj4";
const RBG_TK_SHEET_NAME = "Form Responses 1";


// ===============================================================
// --- USER AUTHENTICATION & PERMISSIONS ---
// ===============================================================

// MODIFICATION: Added a 'name' property to each user for display purposes.
const USER_CREDENTIALS = {
  "jecastro@megaworld-lifestyle.com": { 
    name: "Jeron Luther E.S. Castro",
    password: "qwe", 
    properties: ["ALL"] // Special keyword to grant access to all properties
  },
  "jmpizarro@megaworld-lifestyle.com": { 
    name: "Jaye Trich Pizarro",
    password: "adminpassword", 
    properties: ["ALL"] // Special keyword to grant access to all properties
  },
  "user_ub": {
    name: "Uptown Bonifacio User",
    password: "password123",
    properties: ["UB - Uptown Bonifacio"]
  },
  "user_mkh": {
    name: "McKinley & LCT User",
    password: "password456",
    properties: ["MKH - McKinley Hill", "LCT - Lucky Chinatown", "MKW - McKinley West"]
  },
  "user_multi": {
    name: "Multi-Property User",
    password: "password789",
    properties: ["CCF - Clark Cityfront", "MKH - McKinley Hill"]
  }
};


/**
 * Authenticates a user based on hardcoded credentials.
 * @param {string} username The username entered by the user.
 * @param {string} password The password entered by the user.
 * @returns {object} An object indicating success or failure. On success, it includes the user's full name and properties.
 */
function authenticateUser(username, password) {
  const cleanUsername = username.toLowerCase().trim();
  const user = USER_CREDENTIALS[cleanUsername];

  if (user && user.password === password) {
    // MODIFICATION: Return the 'name' property on successful login.
    return {
      status: "success",
      name: user.name, // <-- Send the user's full name to the client
      properties: user.properties,
    };
  } else {
    return {
      status: "error",
      message: "Invalid username or password.",
    };
  }
}


// ===============================================================
// --- CORE APPLICATION LOGIC (No changes below this line) ---
// ===============================================================

function doGet() {
    return HtmlService.createHtmlOutputFromFile("index")
        .addMetaTag("viewport", "width=device-width, initial-scale=1")
        .setTitle("MSP Contracts Workflow Monitoring")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===============================================================
// --- REVISED CODE: Only this function needs to be updated ---
// ===============================================================

function uploadFileAndLogData(fileData) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);
    } catch (e) {
        return { status: "error", message: "Server is busy, please try again." };
    }

    try {
        const FOLDER_ID = "1ROKJuZ6x5BqXXish9CQedJULJ3nBpx5T";
        const SUBMISSION_SHEET_ID = "1Aa2O2XVinhL7x2zaBlCgpAcv7IcLAcqeUgGVfRNidzw";
        const SUBMISSION_SHEET_NAME = "ContractFinderSubmissions";
        const submitterName = fileData.userName;
        const originalFileName = fileData.fileName;

        // --- LOOKUP LOGIC (No changes here) ---
        const sourceSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        const headers = sourceSheet.getRange(HEADER_ROW_NUMBER, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
        
        const refIndex = headers.indexOf(SEARCH_COLUMN);
        const townshipIndex = headers.indexOf("Township or Property or Group");
        const kindOfContractIndex = headers.indexOf("Kind of Contract");

        if (refIndex === -1 || townshipIndex === -1 || kindOfContractIndex === -1) {
             return { status: "error", message: "Configuration error: Could not find required columns (Ref#, Township, or Kind of Contract)." };
        }

        const dataRange = sourceSheet.getRange(HEADER_ROW_NUMBER + 1, 1, sourceSheet.getLastRow() - HEADER_ROW_NUMBER, headers.length);
        const allData = dataRange.getValues();
        
        let townshipValue = null;
        let kindOfContractValue = null;

        for (const row of allData) {
            if (String(row[refIndex]).trim() === String(fileData.ref).trim()) {
                townshipValue = String(row[townshipIndex]).trim();
                kindOfContractValue = String(row[kindOfContractIndex]).trim();
                break;
            }
        }
        
        if (!townshipValue || !kindOfContractValue) {
             return { status: "error", message: `Could not find required details (Township/Kind of Contract) for Ref#: ${fileData.ref}` };
        }

        const combinedName = `${townshipValue} - ${kindOfContractValue}`;
        const newFileName = combinedName.replace(/[\\/:"*?<>|]/g, '-');

        const decoded = Utilities.base64Decode(fileData.base64Data);
        const blob = Utilities.newBlob(decoded, "application/pdf", newFileName);
        const folder = DriveApp.getFolderById(FOLDER_ID);
        const newFile = folder.createFile(blob);
        const fileUrl = newFile.getUrl();

        const submissionSheet = SpreadsheetApp.openById(SUBMISSION_SHEET_ID).getSheetByName(SUBMISSION_SHEET_NAME);
        
        const range = submissionSheet.getRange("A1:F" + submissionSheet.getMaxRows());
        const values = range.getValues();
        let firstEmptyRow = values.findIndex(row => row.join("").trim() === "") + 1;
        if (firstEmptyRow === 0) firstEmptyRow = submissionSheet.getLastRow() + 1;

      
        submissionSheet.getRange(firstEmptyRow, 1, 1, 6).setValues([[
            new Date(),
            fileData.ref,
            kindOfContractValue,
            fileUrl,
            submitterName,
            newFileName, 
        ]]);

        return { status: "success", message: `Successfully uploaded: ${originalFileName}` };

    } catch (e) {
        console.error(`Upload failed for Ref#: ${fileData.ref}. Error: ${e.toString()}`);
        return { status: "error", message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}
function getInitialData(session) {
    try {
        const userProperties = session.properties || [];
        if (userProperties.length === 0) {
           throw new Error("Access denied. No properties assigned to this user.");
        }

        let allProcessedData = [];

        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const primarySheet = ss.getSheetByName(SHEET_NAME); 
        
        if (!primarySheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

        // --- LOGIC FOR PRIMARY SHEET ---
        const primaryHeaderRow = 9; // Row 9 is where data starts, so headers are on row 8
        const primaryStartRow = 11;
        const primaryHeaders = primarySheet.getRange(primaryHeaderRow, 1, 1, primarySheet.getLastColumn()).getDisplayValues()[0];
        
        const primaryNumRows = primarySheet.getLastRow() - primaryHeaderRow;
        let primaryData = [];
        if (primaryNumRows > 0) {
            primaryData = primarySheet.getRange(primaryStartRow, 1, primaryNumRows, primarySheet.getLastColumn()).getDisplayValues();
        }

        const headerMap = new Map(primaryHeaders.map((h, i) => [h.trim(), i]));
        const finalHeaders = COLUMNS_TO_INCLUDE.filter(h => headerMap.has(h));
        
        const cogData = primaryData.map(row => {
            const rowObject = { source: "COG-SCSU" };
            let rowHasValues = false;

            finalHeaders.forEach(header => {
                const colIndex = headerMap.get(header);
                const cellValue = row[colIndex] || "";
                rowObject[header] = cellValue;
                if (cellValue.trim() !== '') {
                    rowHasValues = true;
                }
            });

            return rowHasValues ? rowObject : null;
        }).filter(Boolean);

        allProcessedData.push(...cogData);
        
        // --- MODIFIED LOGIC FOR SECONDARY SHEET ---
        const rbgTkSpreadsheet = SpreadsheetApp.openById(RBG_TK_SHEET_ID);
        const rbgTkSheet = rbgTkSpreadsheet.getSheetByName(RBG_TK_SHEET_NAME);
        if (rbgTkSheet) {
            const rbgTkHeaderRow = 4; // Data starts at row 5, so headers are on row 4
            const rbgTkStartRow = 5;
            
            const rbgTkHeaders = rbgTkSheet.getRange(rbgTkHeaderRow, 1, 1, rbgTkSheet.getLastColumn()).getDisplayValues()[0].map(h => h.trim());
            
            const rbgTkNumRows = rbgTkSheet.getLastRow() - rbgTkHeaderRow;
            let rbgTkRows = [];
            if (rbgTkNumRows > 0) {
                rbgTkRows = rbgTkSheet.getRange(rbgTkStartRow, 1, rbgTkNumRows, rbgTkSheet.getLastColumn()).getDisplayValues();
            }
            
            const rbgTkData = rbgTkRows.map(row => {
                // --- CHANGE START ---
                // Check if Column V (index 21) or Column Y (index 24) has content.
                // If either column has data, skip this row.
                // In a zero-based array: Column V is at index 21, Column Y is at index 24.
                const colVValue = row[21] || "";
                const colYValue = row[24] || "";

                if (colVValue.trim() !== '' || colYValue.trim() !== '') {
                    return null; // Skip this row by returning null
                }
                // --- CHANGE END ---

                const rowObject = { source: "RBG-TK" };
                let valueCount = 0;

                rbgTkHeaders.forEach((header, index) => {
                    const cellValue = row[index] || "";
                    if (header) {
                       rowObject[header] = cellValue;
                    }
                    if (cellValue.trim() !== '') {
                        valueCount++;
                    }
                });

                // Only return the row object if it has at least 5 cells with content.
                return valueCount >= 5 ? rowObject : null;
            }).filter(Boolean); // The .filter(Boolean) conveniently removes all null entries.
            
            allProcessedData.push(...rbgTkData);
        }

        // Filtering logic remains the same
        const canSeeAll = userProperties.includes("ALL");
        const filteredData = canSeeAll ? allProcessedData : allProcessedData.filter(row => {
            const propertyKey = row.source === 'RBG-TK' ? 'PROPERTY' : 'Township or Property or Group';
            const property = row[propertyKey] ? row[propertyKey].trim() : '';
            return userProperties.includes(property);
        });

        return {
            headers: finalHeaders,
            allRows: filteredData,
        };
    } catch (error) {
        console.error("Error in getInitialData:", error.message, error.stack);
        return { error: `Server-side error: ${error.message}` };
    }
}
