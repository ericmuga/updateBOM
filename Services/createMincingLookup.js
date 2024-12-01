import xlsx from 'xlsx';
import fs from 'fs';
import path from 'path';

const convertExcelToJson = (filePath) => {
    // Read the Excel file
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // Use the first sheet
    const worksheet = workbook.Sheets[sheetName];

    // Convert the worksheet to JSON
    const rawData = xlsx.utils.sheet_to_json(worksheet);

    // Group data into the desired format
    const groupedData = rawData.reduce((result, row) => {
        const fromLocation = row['From Location']; // From Location
        const toLocation = row['ToLoaction']; // To Location
        const outputItem = row['Output Item']; // Output Item
        const outputDescription = row['Output Item Description']; // Output Item Description
        const processLoss = parseFloat(row['ScrapPercentage']); // Process Loss
        const intakeItem = row['No_']; // Intake Item Number (No_)

        // Find or create the group by from/to locations and output item
        let group = result.find(
            (entry) =>
                entry.from === fromLocation &&
                entry.to === toLocation &&
                entry.output_item === outputItem
        );

        if (!group) {
            // Create a new group
            group = {
                from: fromLocation,
                to: toLocation,
                output_item: outputItem,
                output_description: outputDescription,
                process_loss: processLoss,
                intake_items: [],
            };
            result.push(group);
        }

        // Add intake item to the group if it doesn't already exist
        if (intakeItem && !group.intake_items.includes(intakeItem)) {
            group.intake_items.push(intakeItem);
        }

        return result;
    }, []);

    return groupedData;
};

// Path to your Excel file
const filePath = path.resolve('Mincing.xlsx');

// Convert and save the result as JSON
try {
    const jsonData = convertExcelToJson(filePath);
    const outputPath = path.resolve('Mincing.json');
    fs.writeFileSync(outputPath, JSON.stringify(jsonData, null, 2));
    console.log(`JSON data saved to ${outputPath}`);
} catch (error) {
    console.error('Error converting Excel to JSON:', error);
}
