import fs from 'fs';
import xlsx from 'xlsx';
import path from 'path';

// Path to your JSON file
const jsonFilePath = path.resolve('./mincingLines.json');

// Get today's date in YYYY-MM-DD format
const today = new Date().toISOString().split('T')[0];

// Function to create a single Excel file
const createSingleExcelFile = (filePath) => {
    // Read and parse the JSON data
    const jsonData = JSON.parse(fs.readFileSync(filePath, 'utf-8'));

    // Prepare a single array to store all consumption entries
    const allConsumptionEntries = [];

    jsonData.forEach((order) => {
        const { production_order_no, ProductionJournalLines } = order;

        // Filter consumption entries and add them to the combined array
        ProductionJournalLines.filter(line => line.type === 'consumption').forEach(entry => {
            allConsumptionEntries.push({
                Date: today,
                ProductionOrder: production_order_no,
                Item: entry.ItemNo,
                LocationCode: entry.LocationCode,
                UOM: entry.uom
            });
        });
    });

    // Create a worksheet from the combined data
    const worksheet = xlsx.utils.json_to_sheet(allConsumptionEntries);

    // Create a workbook and append the worksheet
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'All Consumption Data');

    // Save the workbook as a single Excel file
    const filename = `All_Consumption_Data.xlsx`;
    xlsx.writeFile(workbook, filename);

    console.log(`Created Excel file: ${filename}`);
};

// Run the function with the JSON file path
createSingleExcelFile(jsonFilePath);
