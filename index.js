/* jshint esversion: 8 */
const assert = require('assert');
const async = require('async');
const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');
let request = require('request');
const FileCookieStore = require('tough-cookie-filestore');

// 1. Load Configuration
const configPath = path.join(__dirname, 'config.json');
if (!fs.existsSync(configPath)) {
    console.error('Error: config.json not found.');
    process.exit(-1);
}
const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));

// 2. Setup Request Jar and Cookies
const cookiesPath = path.join(__dirname, 'cookies.json');
if (!fs.existsSync(cookiesPath)) {
    fs.writeFileSync(cookiesPath, '{}');
}
let j = request.jar(new FileCookieStore(cookiesPath));
request = request.defaults({ jar: j });

const cookie_url = 'https://partner.steamgames.com';
config.authCookies.split('; ').forEach(function (cookie) {
    if (cookie.trim()) {
        j.setCookie(cookie.trim(), cookie_url);
    }
});

// 3. Initialize Google Sheets Auth
const auth = new google.auth.GoogleAuth({
    keyFile: path.join(__dirname, config.googleApiAuthFile),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

/**
 * Column to Index converter (A=0, B=1, etc.)
 */
function colToIndex(col) {
    if (!col || typeof col !== 'string') return 0;
    let index = 0;
    const cleanCol = col.toUpperCase().replace(/[^A-Z]/g, '');
    for (let i = 0; i < cleanCol.length; i++) {
        index = index * 26 + cleanCol.charCodeAt(i) - 64;
    }
    return index - 1;
}

/**
 * Gets the internal numeric Sheet ID for a given sheet name
 */
async function getSheetId(spreadsheetId, sheetName) {
    const res = await sheets.spreadsheets.get({ spreadsheetId });
    const sheet = res.data.sheets.find(s => s.properties.title === sheetName);
    return sheet ? sheet.properties.sheetId : 0;
}

/**
 * Fetches values AND notes from column A up to the furthest required column
 */
async function getDataFromSheets() {
    const colLetters = [config.keyColumn, config.resultColumn, config.filterColumn];
    const furthestCol = colLetters.map(c => c.toUpperCase()).sort((a, b) => {
        return colToIndex(b) - colToIndex(a);
    })[0];

    // Always start at A to ensure indices are absolute (A=0, B=1...)
    const range = `${config.sheetName}!A:${furthestCol}`;
    
    try {
        const response = await sheets.spreadsheets.get({
            spreadsheetId: config.spreadsheetId,
            ranges: [range],
            includeGridData: true,
            fields: 'sheets.data.rowData.values(userEnteredValue,note)'
        });

        const rowData = response.data.sheets[0].data[0].rowData || [];
        
        return rowData.map(row => {
            return (row.values || []).map(cell => ({
                value: cell.userEnteredValue ? (cell.userEnteredValue.stringValue || cell.userEnteredValue.numberValue || "").toString() : "",
                note: cell.note || ""
            }));
        });
    } catch (err) {
        console.error('Fetch Error:', err.message);
        process.exit(-1);
    }
}

/**
 * Updates cell values and notes via batchUpdate
 */
async function batchWriteResults(resultsArray) {
    const sheetId = await getSheetId(config.spreadsheetId, config.sheetName);
    const resColIdx = colToIndex(config.resultColumn);

    const requests = resultsArray.map((item, index) => {
        return {
            updateCells: {
                range: {
                    sheetId: sheetId,
                    startRowIndex: index,
                    endRowIndex: index + 1,
                    startColumnIndex: resColIdx,
                    endColumnIndex: resColIdx + 1
                },
                rows: [{
                    values: [{
                        userEnteredValue: { stringValue: item.status || "" },
                        note: item.note || ""
                    }]
                }],
                fields: 'userEnteredValue,note'
            }
        };
    });

    try {
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId: config.spreadsheetId,
            resource: { requests }
        });
        console.log(`Successfully synced ${resultsArray.length} rows to the sheet.`);
    } catch (err) {
        console.error('Batch Update Error:', err.message);
    }
}

/**
 * Queries the Steam Partner API for a single key
 */
function querySteam(key) {
    return new Promise((resolve, reject) => {
        let retries = 0;
        function doAttempt() {
            let url = `https://partner.steamgames.com/querycdkey/cdkey?cdkey=${key}&method=Query`;
            request({
                url: url,
                headers: {
                    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36',
                },
            }, function (err, resp, body) {
                if (err) return reject(err);
                
                let parts = body.split('<h2>Activation Details</h2>');
                if (parts.length < 2) {
                    if (retries < 5) {
                        retries++;
                        return setTimeout(doAttempt, 1000);
                    }
                    if (body.includes('login') || body.includes('Sign In')) {
                        return reject(new Error("Auth failed: Update cookies in config.json"));
                    }
                    return resolve({ status: "Error/RetryNeeded", note: "" });
                }

                let resultPart = parts[1].split('</table>')[0];
                let matches = resultPart.match(/<td>.*<\/td>/gu);
                if (!matches) return resolve({ status: "No Data", note: "" });

                let parsed = matches.map(line => line.replace(/<[^>]*>/gu, ''));
                
                // parsed[0] is Status (Activated, Fail, etc.)
                // parsed[1] is the Details (Game Name)
                resolve({ 
                    status: parsed[0] || "Unknown", 
                    note: parsed[1] || "" 
                });
            });
        }
        doAttempt();
    });
}

/**
 * Main execution flow
 */
async function run() {
    const keyIdx = colToIndex(config.keyColumn);
    const resIdx = colToIndex(config.resultColumn);
    const filIdx = colToIndex(config.filterColumn);
    let index = -1;

    console.log(`Starting process...`);
    console.log(`Mapping: Key=${config.keyColumn}(${keyIdx}), Result=${config.resultColumn}(${resIdx}), Filter=${config.filterColumn}(${filIdx})`);
    
    const rows = await getDataFromSheets();

    try {
        const results = await async.mapLimit(rows, 1, async (row) => {
            ++index;

            const keyCell = row[keyIdx] || { value: "", note: "" };
            const resCell = row[resIdx] || { value: "", note: "" };
            const filCell = row[filIdx] || { value: "", note: "" };

            const key = keyCell.value.trim();
            const existingResult = resCell.value.trim();
            const filterValue = filCell.value.trim();
            const existingNote = resCell.note;

            // 1. Skip Header
            if (index === 0) {
                return { status: existingResult, note: existingNote };
            }

            // 2. Skip if already 'Activated'
            if (existingResult.toLowerCase() === 'activated') {
                console.log(`[Row ${index + 1}] Skipping: Already activated.`);
                return { status: existingResult, note: existingNote };
            }

            // 3. Skip if filter condition not met or key empty
            if (!config.filterValues.includes(filterValue) || !key) {
                return { status: existingResult, note: existingNote };
            }

            // 4. Query Steam
            const data = await querySteam(key);
            console.log(`[Row ${index + 1}] Processed ${key}: ${data.status}`);
            return data;
        });

        console.log('Writing results and preserving existing notes...');
        await batchWriteResults(results);
        console.log('Done.');
    } catch (err) {
        console.error('Fatal Error during execution:', err);
    }
}

run();