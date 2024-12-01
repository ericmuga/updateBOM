import fs from 'fs';
import sql from 'mssql';
import dbConfig from '../config/dbConfig.js';
import xlsx from 'xlsx';

// Helper function to connect to the database
const connectToDatabase = async () => {
  try {
    return await sql.connect(dbConfig);
  } catch (error) {
    console.error('Error connecting to the database:', error);
    throw error;
  }
};

// Helper function to run a SQL query
const runQuery = async (query) => {
  try {
    const result = await sql.query(query);
    return result.recordset;
  } catch (error) {
    console.error('Error executing query:', error);
    throw error;
  }
};

// Fetch and structure data from the database, then save it to JSON
export const fetchData = async () => {
  try {
    await connectToDatabase();

    // Fetch distinct BOMNos
    const distinctBomNoQuery = `
      SELECT DISTINCT b.[No_] AS BOMNo
      FROM [fcl-bc-main].[dbo].[FCL$Production BOM Header] AS b
      INNER JOIN [fcl-bc-main].dbo.[FCL$Item Ledger Entry] AS a 
      ON a.[Document No_] = b.[No_]
      WHERE a.[Posting Date] >= '2024-01-01'
    `;
    const bomNos = await runQuery(distinctBomNoQuery);
    if (bomNos.length === 0) {
      console.log('No data found for the specified date range.');
      return;
    }

    const bomNosList = bomNos.map(row => `'${row.BOMNo}'`).join(',');

    // Fetch header data
    const headerQuery = `
      SELECT 
         b.[No_] AS BOMNo, b.[Description 2], b.[Search Name], b.[Unit of Measure Code],
         b.[Low-Level Code], b.[Creation Date], b.[Last Date Modified], b.[Status],
         b.[Version Nos_], b.[No_ Series], b.[Description],
         s.[BOMNo] AS SwapBOMNo, s.[FromLocation], s.[ToLoaction], s.[ScrapPercentage], 
         s.[Output Item], s.[Output Item Short Code], s.[Blocked],
         s.[Cost Calculation Type], s.[Linked BOM or Family], s.[Calculation Week],
         s.[Output Item Default Qty], s.[Default Batch Qty], s.[Packaging Template],
         s.[Active for Transfers], s.[Level Code], s.[Pieces Mandatory],
         s.[Override Average Cost]
      FROM [fcl-bc-main].[dbo].[FCL$Production BOM Header] AS b
      LEFT JOIN [fcl-bc-main].[dbo].[FCL$BOMSwapHeader] AS s ON s.[BOMNo] = b.[No_]
      WHERE b.[No_] IN (${bomNosList})
    `;
    const headers = await runQuery(headerQuery);

    // Fetch line data
    const linesQuery = `
      SELECT DISTINCT
         c.[Production BOM No_] AS BOMNo, c.[Line No_], c.[Type], c.[No_],
         c.[Description], c.[Unit of Measure Code], c.[Quantity], c.[Position],
         c.[Position 2], c.[Position 3], c.[Lead-Time Offset], c.[Routing Link Code],
         c.[Scrap _], c.[Variant Code], c.[Starting Date], c.[Ending Date], c.[Length],
         c.[Width], c.[Weight], c.[Depth], c.[Calculation Formula], c.[Quantity per],
         l.[BOMNo] AS SwapBOMNo, l.[ItemNo], l.[Location Code], l.[Short Code],
         l.[AutoAccummulate], l.[Units Per 100], l.[Linked BOM or Family],
         l.[LinkedToInput], l.[LinkedToOutput], l.[Default Qty], l.[Pieces Mandatory]
      FROM [fcl-bc-main].[dbo].[FCL$Production BOM Line] AS c
      LEFT JOIN [fcl-bc-main].[dbo].[FCL$BOMSwapLines] AS l ON l.[BOMNo] = c.[Production BOM No_] AND l.[ItemNo] = c.[No_]
      WHERE c.[Production BOM No_] IN (${bomNosList})
    `;
    const lines = await runQuery(linesQuery);

    // Combine header and line data into the desired structure
    const combinedData = headers.map(header => {
      const headerWithLines = {
        ...header,
        Lines: lines.filter(line => line.BOMNo === header.BOMNo)
      };
      return headerWithLines;
    });

    // Save the combined data to JSON
    fs.writeFileSync('grouped_output.json', JSON.stringify(combinedData, null, 2));
    console.log('Combined data with headers and lines saved to grouped_output.json');

    await sql.close();
  } catch (err) {
    console.error('Error fetching data:', err);
  }
};



// Read from initial JSON, apply grouping, and save to new JSON and Excel
export const fetchDataAndGroup = async () => {
  try {
    // Step 1: Read the initial data from grouped_output.json
    const data = JSON.parse(fs.readFileSync('grouped_output.json', 'utf-8'));

    // Step 2: Group data by Level Code and Output Item
    const groupedData = data.reduce((acc, item) => {
      const levelCode = item['Level Code'];
      const outputItem = item['Output Item'];

      // Initialize the structure for Level Code if it doesn’t exist
      if (!acc[levelCode]) acc[levelCode] = {};

      // Initialize the structure for Output Item within Level Code if it doesn’t exist
      if (!acc[levelCode][outputItem]) {
        acc[levelCode][outputItem] = {
          header: { ...item },
          lines: item.Lines
        };
        delete acc[levelCode][outputItem].header.Lines; // Remove lines from header
      }

      return acc;
    }, {});

    // Step 3: Save the final grouped data to a new JSON file
    fs.writeFileSync('final_grouped_output.json', JSON.stringify(groupedData, null, 2));
    console.log('Final grouped data saved to final_grouped_output.json');

    // Step 4: Prepare data for Excel export by flattening the nested structure
    const workbook = xlsx.utils.book_new();
    const finalSheetData = [];

    // Flatten the data for Excel
    Object.entries(groupedData).forEach(([levelCode, outputItems]) => {
      Object.entries(outputItems).forEach(([outputItem, details]) => {
        // Add each line item as a row in the Excel sheet
        details.lines.forEach(line => {
          finalSheetData.push({
            LevelCode: levelCode,
            OutputItem: outputItem,
            ...details.header, // Include all header details
            ...line            // Include all line details
          });
        });
      });
    });

    // Convert to Excel format and save
    const sheet = xlsx.utils.json_to_sheet(finalSheetData);
    xlsx.utils.book_append_sheet(workbook, sheet, 'GroupedData');
    xlsx.writeFile(workbook, 'final_grouped_output.xlsx');
    console.log('Final grouped data saved to final_grouped_output.xlsx');
  } catch (err) {
    console.error('Error processing and further grouping data:', err);
  }
};









const fetchConversionFactor = async (itemNo, unitCode) => {
    await connectToDatabase();
    const query = `
      SELECT [Qty_ per Unit of Measure] AS ConversionFactor
      FROM [fcl-bc-main].[dbo].[FCL$Item Unit of Measure]
      WHERE [Item No_] = '${itemNo}' AND [Code] = '${unitCode}'
    `;
    const result = await runQuery(query);
    await sql.close();
    return result.length > 0 ? result[0].ConversionFactor : 1;
  };
  
  // Main function to process BOM data
  export const processBOMData = async () => {
    // Load the JSON data
    const data = JSON.parse(fs.readFileSync('grouped_output.json', 'utf-8'));
  
    // Create a map of output items to their BOMNo for easy lookup
    const outputItemToBOMNo = new Map(data.map(item => [item['Output Item'], item.BOMNo]));
  
    // Process each line item in all BOMs
    for (const header of data) {
      for (const line of header.Lines) {
        // If the line item (No_) appears as an output item in the headers
        if (outputItemToBOMNo.has(line['No_'])) {
          // Replace No_ with the header BOMNo and set Type to 2
          line['No_'] = outputItemToBOMNo.get(line['No_']);
          line['Type'] = 2;
  
          // Update Quantity per
          line['Quantity per'] = line['Units Per 100']
            ? line['Units Per 100'] / 100
            : line['Quantity per'] / 100;
  
          // Fetch conversion factor and adjust Quantity
          const conversionFactor = await fetchConversionFactor(line['No_'], line['Unit of Measure Code']);
          line['Quantity'] = line['Quantity per'] * conversionFactor;
        }
      }
    }
  
    // Save updated data to a new JSON file
    fs.writeFileSync('updated_output.json', JSON.stringify(data, null, 2));
    console.log('Updated data saved to updated_output.json');

    //read the json file
    const jsonData = JSON.parse(fs.readFileSync('updated_output.json', 'utf-8'));
    
}// Run the processing function

// Function to create Excel workbook with header and lines worksheets
const createExcelWorkbook = async () => {
  // Load the JSON data
  const data = JSON.parse(fs.readFileSync('updated_output.json', 'utf-8'));

  // Prepare data for the 'Header' sheet
  const headerData = data.map(item => {
    // Exclude 'Lines' from the header data
    const { Lines, ...headerInfo } = item;
    return headerInfo;
  });

  // Prepare data for the 'Lines' sheet
  const lineData = data.flatMap(item => {
    return item.Lines.map(line => ({
      ...line,
      BOMNo: item.BOMNo // Include the BOMNo from the header in each line entry
    }));
  });

  // Create a new workbook
  const workbook = xlsx.utils.book_new();

  // Convert header data to a sheet and add it to the workbook
  const headerSheet = xlsx.utils.json_to_sheet(headerData);
  xlsx.utils.book_append_sheet(workbook, headerSheet, 'Header');

  // Convert line data to a sheet and add it to the workbook
  const lineSheet = xlsx.utils.json_to_sheet(lineData);
  xlsx.utils.book_append_sheet(workbook, lineSheet, 'Lines');

  // Write the workbook to a file
  xlsx.writeFile(workbook, 'bom_data.xlsx');
  console.log('Workbook created successfully as bom_data.xlsx');
};



await fetchData();

// Run the grouping function after initial data is fetched
await fetchDataAndGroup();

await processBOMData();

await createExcelWorkbook();
