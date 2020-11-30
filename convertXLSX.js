//
// Convert the ASCCP's published Excel spreadsheets to JSON files
// for use in the PapMap application
//
// Daniel Rosenberg MD
// August 2020
// Portland, Oregon
//

// const axios = require('axios');


// We really can't just download these from the web, since we need additional files to be included
// to account for 21-24 year olds, the odd way table 1 uses ALL to mean None, and for immunocompromised and pregnant patients;
// instead, we'll look in the current directory for files beginning with [1-5] and ending with .xlsx

// var urls = [
//     'https://cervixca.nlm.nih.gov/RiskTables/1-General%20Table%20for%20Screening_locked.xlsx',
//     'https://cervixca.nlm.nih.gov/RiskTables/2-General%20Table%20for%20Surveillance_locked.xlsx',
//     'https://cervixca.nlm.nih.gov/RiskTables/3-General%20Table%20for%20Risk%20Following%20Colposcpy_locked.xlsx',
//     'https://cervixca.nlm.nih.gov/RiskTables/4-General%20Table%20for%20Post-Colpo.xlsx',
//     'https://cervixca.nlm.nih.gov/RiskTables/5-General%20Table%20for%20Post-Treatment_locked.xlsx'
// ];
//
// var files = [
//     '1-General Table for Screening_locked.xlsx',I
//     '2-General Table for Surveillance_locked.xlsx',
//     '3-General Table for Risk Following Colposcpy_locked.xlsx',
//     '4-General Table for Post-Colpo.xlsx',
//     '5-General Table for Post-Treatment_locked.xlsx'
// ];

// In Table 1, we need the first 4 columns (age, past history, current HPV result, current Pap result), CIN3+ immediate risk (AR), CIN 3+ 5 year risk (BL),
// the Management column (CN), and Management confidence probability (CO)
// Table 1, row 8 jas a current Pap result of ALL - this is not a guideline but just the sum of the previous columns. Still it has a 5 year followup recommendation
// Clinical action thresholds: 5 year return 0-0.14% risk (of CIN3+ in the next 5 years; 3 year return 0.15-0.54%, 1 year return 0.55-9%
// Expedited treatment: 60-100% immediate CIN3+ risk, Expedited tx or colposcopy acceptable 25-59% immediate CIN3+ risk, colposcopy recommended 4-24%

const fs = require('fs');
const XLSX = require('xlsx');
const xlsDirectory = 'xlsx';
const jsonDirectory = 'json';

const scenario = [
    'Pap management',
    'Pap management with prior abnormal Pap',
    'Post-colposcopy plan',
    'Pap after colposcopy',
    'Pap after treatment',       // Excel file lists bx result just as CIN 3, but article CIN 2 or 3; "ALL" Pap means primary HPV screening without Pap
                                // Cotest-negative x 2 row with no current result equates to co-test neg x 1, current co-test negative in article
                                // and similar for the rest of the current Pap results - they are blank, but they are really included in the Current HPV result column
    'Post-treatment'
];


var columnCapture = [
    // Scenario (table) 1: capture columns Age (A), PAST HISTORY (most recent) (B), Current HPV result (C), current PAP result (D),
    // CIN 3+ Immediate risk (%) (AR), CIN 3+ 5 year risk (%) (BL), Management (CN), Management Confidence Probability (CO)
    [ 'Age', 'PAST HISTORY (most recent)', 'Current HPV Result', 'Current PAP Result', 'CIN3+ Immediate risk (%)', 'CIN3+ 5 year risk  (%)', 'Management', 'Management Confidence Probability', 'Notes', 'Figure' ],

    // Scenario 2: capture columns  Age (A), PAST HISTORY (previous 2) (B), PAST HISTORY (most recent ) (C), Current HPV result (D),
    // Current PAP result (E), CIN 3+ 5 yer risk (BM), Management (CO), Management Confidence Probability (CP)
    [ 'Age', 'PAST HISTORY (previous 2)', 'PAST HISTORY (most recent)', 'Current HPV Result', 'Current PAP Result', 'CIN3+ Immediate risk (%)', 'CIN3+ 5 year risk  (%)', 'Management', 'Management Confidence Probability', 'Notes', 'Figure'],

    // Scenario 3: capture columns Age (A), Referral Screen Result (B), Biopsy Result (C), CIN3+ 5 year risk (%) H, and Management (I)
    [ 'Age', 'Referral Screen Result', 'Biopsy Result', 'CIN3+ Immediate risk (%)', 'CIN3+ 5 year risk  (%)', 'Management', 'Notes'],

    // Scenario 4: capture columns: Age (A), Pre-Colpo Test Result (B), Pre-Colpo Test Result PAST HISTORY (C),
    // Current HPV Result (D), Current PAP Result (E), CIN3+ 5 yer risk (%) (BM), Management (CO), Management Confidence Probability (CP)
    [ 'Age', 'Pre-Colpo Test Result', 'Post-Colpo HPV Result - PAST HISTORY', 'Post-Colpo Test Result - PAST HISTORY',
        'Post-Colpo Test Result - Prior PAST HISTORY', 'Post-Colpo HPV Result - Prior PAST HISTORY',
        'Current HPV Result', 'Current PAP Result', 'CIN3+ Immediate risk (%)', 'CIN3+ 5 year risk  (%)', 'Management',
        'Management Confidence Probability', 'Figure', 'Notes'],

    // Scenario 5: capture columns: Age (A), Test Result Before Biopsy (B) even though it's always blank, Biopsy Result Before Treatment (C)
    // which is always CIN3 but means CIN2 or CIN3, Current HPV Result (D), Current Pap Result (E), CIN3+ 5 years risk (%) (BL), Management (CN),
    // Management Confidence Probability (CO)
    [ 'Age', 'Test Result Before Biopsy', 'Biopsy Result Before Treatment', 'Current HPV Result', 'Current PAP Result', 'CIN3+ Immediate risk (%)', 'CIN3+ 5 year risk  (%)', 'Management', 'Management Confidence Probability', 'Notes'],

    // Scenario 6: capture everything
    [ 'Age', 'Biopsy Result Before Treatment', 'Margins', 'Treatment', 'Management', 'Notes']


]
// Use XLSX to convert the Excel file to JSON
function convertToJSON(workbook, scenario) {

    // var table = [];
    let data = [];

    let sheets = workbook.SheetNames;
    console.log(`>> Found ${sheets.length} sheet(s).`)
    sheets.forEach(sheetIndex => {
        var worksheet = workbook.Sheets[sheetIndex];
        // var json = XLSX.sheet_to_json(worksheet);
        // console.log(`JSON: ${json}`);

        var headers = {};
        // var data = [];
        for (var z in worksheet) {
            if (z[0] === '!') continue;
            var tt = 0;
            for (var i = 0; i < z.length; i++) {
                if (!isNaN(z[i])) {
                    tt = i;
                    break;
                }
            }

            var col = z.substring(0, tt);
            var row = parseInt(z.substring(tt));
            var value = worksheet[z].v;


            if (row == 1 && value) {
                headers[col] = value;
                // console.log(`Header row: ${value}, col ${col}`);
                continue;
            }

            // Only bother with the columns we care about for a given scenario
            // if (value && col && columnCapture[scenario - 1].includes(col)) {
            if (value && col && columnCapture[scenario].includes(headers[col])) {

                // console.log(`Scenario ${scenario}, ${col}, row ${row}`);
                if (!data[row]) data[row] = {};
                data[row][headers[col]] = value;
                // console.log(`Row ${row}: ${value}`)

            }

        }
        // console.log(`Data: ${data}`);
    })
    return data;
}

// // Convert an ASCCP Excel file - grab it from the web
// function convertExcelURL(url) {
//
//
//     axios({
//         method: 'GET',
//         url: url,
//         responseType: 'arraybuffer', // blob vs arraybuffer
//         headers: {'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},
//     })
//         .then((response) => {
//             // const blob = new Blob([response.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8'});
//              var dataArray = new Uint8Array(response.data);
//              console.log(dataArray);
//              var workbook = XLSX.read(dataArray, {type: "array"})
//             convertToJSON(workbook);
//         })
//         .catch(response => {
//             console.log(`Unable to download ${url}`)
//         })
//
// }

function convertExcelFile(filename, scenario) {
    let workbook = XLSX.readFile(filename);

    let data = [];

    console.log(`Converting ${filename}, scenario ${scenario}`);
    data = convertToJSON(workbook, scenario);
    // console.log(`Data: ${data}`);
    return data;

}
// Main execution entry
const glob = require('glob-fs')({ gitignore: true });

// urls.forEach(convertExcelURL);

let json = [];
let data = [];
let moreData = [];
let matchedFiles = [];
const fileArg = xlsDirectory + `/[1-${columnCapture.length+1}]*.xlsx`;

// Calling readdirSync in the for loop would accumulate files even with distinct glob arguments
const files =  glob.readdirSync(fileArg, {});

console.log("Found files: " + files);

for (var i = 0; i < scenario.length; i++) {
    // var fileRegex = new RegExp(`${xlsDirectory}/${i+1}*.xlsx `, 'g');

    // Who knows while this doesn't work?


    // var files =  glob.readdirSync(xlsDirectory + "\/" + (i+1) + ".*\.xlsx", {});
    // var files =  glob.readdirSync(`${xlsDirectory}\/${i+1}.*\.xlsx`, {});

    // The scenario is the number at the beginning of the filename
    // If it's got a decimal (e.g., 3.1) it's still the same scenario - just a file we added
    // because it's not covered in the Excel files distributed by the NIH - we derived it from the article text
    // instead
    matchedFiles = files.filter(file => {
        var regexp = new RegExp(`${xlsDirectory}\/${i+1}.*`, 'g');
        return regexp.exec(file);
    })
    console.log(`Working on scenario ${i+1}, matched ${matchedFiles}`);

    // Go through Scenario 3, 3.1, etc.
    // We're storing scenario 1 in position 0 in the array, 2 in 1, etc.
    data = [];
    matchedFiles.forEach(filename => {
        console.log(`Filename ${filename}`)
        if (data.length == 0) {
            data.push(convertExcelFile(filename, i));
        } else {
            // We're concatenating the 2nd (3.1) etc. Excel files here into the same array, not as a secondary array
            // added onto the first
            moreData = (convertExcelFile(filename, i))
            moreData.forEach(item => {
                data[0].push(item);
            })
        }

    })
    json = JSON.stringify(data);
    fs.writeFile(`${jsonDirectory}/scenario-${i+1}.json`, json, function(err)  {
        if (err) throw err;
    });
}

