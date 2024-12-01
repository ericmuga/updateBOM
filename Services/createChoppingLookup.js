import fs from 'fs';
import xlsx from 'xlsx';

// Load the Excel file
const workbook = xlsx.readFile('./choppingLookup.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];

// Convert the sheet to JSON
const data = xlsx.utils.sheet_to_json(sheet);

// Group data by location and format it
const groupedData = data.reduce((result, { item_code, unit_measure, location }) => {
    if (!result[location]) {
        result[location] = [];
    }
    result[location].push({
        item_no: item_code || 'Unknown', // Fallback if item_code is missing
        uom: unit_measure ? unit_measure : 'unknown' // Fallback if unit_measure is missing
    });
    return result;
}, {});

// Save the result to a JSON file
fs.writeFileSync('output.json', JSON.stringify(groupedData, null, 4));

console.log('Data has been transformed and saved to output.json');
