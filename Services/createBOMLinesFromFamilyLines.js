import xlsx from "xlsx";

const createProductionBOMData = (inputFilePath) => {
  // Read the main Excel file with production BOM data
  const workbook = xlsx.readFile(inputFilePath);
  const sheetName = workbook.SheetNames[0];
  const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

  // Group BOM data by Family No_
  const groupedData = sheetData.reduce((acc, row) => {
    const familyNo = row["Family No_"];
    if (!acc[familyNo]) {
      acc[familyNo] = [];
    }
    acc[familyNo].push(row);
    return acc;
  }, {});

  // Generate BOM Headers and Lines
  const bomHeaders = [];
  const productionBOMLines = [];
  const bomMappings = {}; // To track items and their BOMs for substitution generation

  Object.keys(groupedData).forEach((familyNo) => {
    let suffix = 1; // Reset suffix for each Family No_

    groupedData[familyNo].forEach((row) => {
      // Create BOM No. with incrementing suffix
      const bomNo = `${familyNo}-${String(suffix).padStart(2, "0")}`;

      // Add to BOM Headers
      bomHeaders.push({
        "No.": bomNo,
        "Description": row["Description"], // Use the description from the row
        "Unit of Measure Code": row["Unit of Measure Code"], // Use the unit of measure code from the row
      });

      // Add to Production BOM Lines
      productionBOMLines.push({
        "Production BOM No.": bomNo,
        "Version Code": "", // Assuming Version Code is empty
        "Line No.": 10000, // Static Line No
        "Type": "Item", // Assuming Type is always Item
        "No.": row["Item No_"], // Added No. column
        "Description": row["Description"],
        "Unit of Measure Code": row["Unit of Measure Code"],
        "Quantity": parseFloat(row["Quantity"]),
        "Scrap %": 2, // Assuming Scrap % is 2
        "Quantity per": parseFloat(row["Quantity"]), // Quantity per matches Quantity
      });

      // Track item and its BOM for substitution
      if (!bomMappings[row["Item No_"]]) {
        bomMappings[row["Item No_"]] = [];
      }
      bomMappings[row["Item No_"]].push({
        bomNo,
        intakeItem: row["Item No_"],
      });

      suffix++; // Increment suffix for the next BOM
    });
  });

  return { bomHeaders, productionBOMLines, bomMappings, sheetData };
};

const generateSubstitutes = (bomMappings, sheetData) => {
  const substitutes = [];
  const existingSubstitutes = new Set();

  Object.keys(bomMappings).forEach((itemNo) => {
    const boms = bomMappings[itemNo];

    if (boms.length > 1) {
      // If the item has more than one BOM, find the intake items for each BOM
      for (let i = 0; i < boms.length; i++) {
        for (let j = i + 1; j < boms.length; j++) {
          const intakeItem1 = boms[i].intakeItem;
          const intakeItem2 = boms[j].intakeItem;

          // Check if this substitution already exists
          const key1 = `${intakeItem1}-${intakeItem2}`;
          const key2 = `${intakeItem2}-${intakeItem1}`;
          if (!existingSubstitutes.has(key1) && !existingSubstitutes.has(key2)) {
            substitutes.push({
              "Type": "Item",
              "No.": intakeItem1,
              "Variant Code": "",
              "Substitute Type": "Item",
              "Substitute No.": intakeItem2,
              "Substitute Variant Code": "",
              "Description": sheetData.find((row) => row["Item No_"] === intakeItem1)?.["Description"] || "",
              "Inventory": 0,
              "Interchangeable": true,
              "Relations Level": 0,
              "Quantity Avail. on Shpt. Date": 0,
              "Shipment Date": "",
            });

            substitutes.push({
              "Type": "Item",
              "No.": intakeItem2,
              "Variant Code": "",
              "Substitute Type": "Item",
              "Substitute No.": intakeItem1,
              "Substitute Variant Code": "",
              "Description": sheetData.find((row) => row["Item No_"] === intakeItem2)?.["Description"] || "",
              "Inventory": 0,
              "Interchangeable": true,
              "Relations Level": 0,
              "Quantity Avail. on Shpt. Date": 0,
              "Shipment Date": "",
            });

            existingSubstitutes.add(key1);
            existingSubstitutes.add(key2);
          }
        }
      }
    }
  });

  return substitutes;
};

const writeToExcel = (outputFilePath, bomHeaders, productionBOMLines, substitutes) => {
  // Convert JSON data to worksheets
  const headerWorksheet = xlsx.utils.json_to_sheet(bomHeaders);
  const linesWorksheet = xlsx.utils.json_to_sheet(productionBOMLines);
  const substitutesWorksheet = xlsx.utils.json_to_sheet(substitutes);

  // Create a new workbook and append the worksheets
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, headerWorksheet, "BOM Headers");
  xlsx.utils.book_append_sheet(workbook, linesWorksheet, "Production BOM Lines");
  xlsx.utils.book_append_sheet(workbook, substitutesWorksheet, "Item Substitutes");

  // Write the workbook to the specified file
  xlsx.writeFile(workbook, outputFilePath);

  console.log(`BOM Headers, Lines, and Substitutes written to ${outputFilePath}`);
};

// File paths
const inputFilePath = "./raw.xlsx"; // Path to your BOM data
const outputFilePath = "./FinalProductionBOMWithHeadersAndSubstitutes.xlsx"; // Path to save the final workbook

// Generate BOM and substitutes data
const { bomHeaders, productionBOMLines, bomMappings, sheetData } = createProductionBOMData(inputFilePath);
const substitutes = generateSubstitutes(bomMappings, sheetData);

// Write all data to the final workbook
writeToExcel(outputFilePath, bomHeaders, productionBOMLines, substitutes);
