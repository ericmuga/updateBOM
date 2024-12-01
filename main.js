import xlsx from 'xlsx';
import fs from 'fs';
import sql from 'mssql';
import dbConfig from './config/dbConfig.js';
import path from 'path';
import { fileURLToPath } from 'url';

// Create the equivalent of __dirname
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Define input, output, and processed directories
const inputDir = path.resolve(__dirname, '../updateBOM/Files');
const outputDir = path.resolve(__dirname, '../updateBOM/Output');
const processedDir = path.resolve(__dirname, '../updateBOM/Processed');

// Ensure output and processed directories exist
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}
if (!fs.existsSync(processedDir)) {
  fs.mkdirSync(processedDir, { recursive: true });
}

// Timeout configuration in milliseconds (10 seconds)
const requestTimeout = 10000; // 10 seconds

// Function to process each file
const processFile = async (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert the worksheet to JSON
  const data = xlsx.utils.sheet_to_json(worksheet);

  // Filter out blank rows and rows without an ItemNo
  const filteredData = data.filter(row => 
    row.ItemNo && row.ItemNo.trim() !== "" && 
    row.BOMNo && row.BOMNo.trim() !== ""
  );

  // Get the file name without extension for output JSON file
  const fileName = path.basename(filePath, path.extname(filePath));
  const outputFilePath = path.join(outputDir, `${fileName}.json`);

  // Save filtered data to JSON for preview
  fs.writeFileSync(outputFilePath, JSON.stringify(filteredData, null, 2));
  console.log(`Filtered data saved to ${outputFilePath}.`);

  // Process data in SQL Server
  await processData(filteredData);

  // Move the processed Excel file to the "Processed" directory
  const processedFilePath = path.join(processedDir, path.basename(filePath));
  fs.renameSync(filePath, processedFilePath);
  console.log(`Moved ${filePath} to ${processedFilePath}`);
};

// Function to process data in SQL Server
const processData = async (data) => {
    try {
      // Connect to the database
      const pool = await sql.connect(dbConfig);
  
      // Iterate over each row and process based on "Comments"
      let lineNo = 55000;
      for (const row of data) {
        lineNo++;
        const {
          BOMNo,
          Process,
          ItemNo,
          IntakeItemDescription,
          BaseUOM,
          UsagePerBatch,
          // UnitsPer100,
          LocationCode,
          AutoAccummulate,
          Comments,
          OutputItem,
          OutputItemDescription,
          Scrap=0,
        } = row;
        const UnitsPer100 = parseFloat(row[" Units Per 100 "]).toFixed(2);
        const ScrapValue = parseFloat(Scrap).toFixed(2)|| 0; // Use default if Scrap is missing or null
  // const IntakeItemDescription
        if (Comments === "Item Added") {
          // Insert into FCL$Production BOM Line
          const descriptionValue = IntakeItemDescription || ''; // Use default if IntakeItemDescription is missing or null
          const uom = BaseUOM || 'KG'; // Use default if IntakeItemDescription is missing or null
  
          await pool.request()
          .input("ProductionBOMNo", sql.VarChar, BOMNo)
          .input("VersionCode", sql.Int, 0)  
          .input("Type", sql.Int, 1)
          .input("LineNo", sql.Int, lineNo)
          .input("No_", sql.VarChar, ItemNo)
          .input("IntakeItemDescription", sql.VarChar, descriptionValue)
          .input("UnitOfMeasureCode", sql.VarChar, uom)
          .input("Quantity", sql.Float,UnitsPer100)
          .input("Position", sql.VarChar, '')
          .input("Position2", sql.VarChar, '')
          .input("Position3", sql.VarChar, '')
          .input("LeadTimeOffset", sql.VarChar, '')
          .input("RoutingLinkCode", sql.VarChar, '')
          .input("Scrap_", sql.Float, ScrapValue)
          .input("VariantCode", sql.VarChar, '')
          .input("LinkedBOMOrFamily", sql.VarChar, OutputItem)
          .input("StartingDate", sql.Date, '1753-01-01')
          .input("EndingDate", sql.Date, '1753-01-01')
          .input("Length", sql.Float, 0)
          .input("Width", sql.Float, 0)
          .input("Weight", sql.Float, 0)
          .input("Depth", sql.Float, 0)
          .input("CalculationFormula", sql.VarChar, '')
          .input("QuantityPer", sql.Float, 1)
          .query(`
            IF NOT EXISTS (
              SELECT 1 FROM [dbo].[FCL$Production BOM Line] 
              WHERE [Production BOM No_] = @ProductionBOMNo AND [No_] = @No_
            )
            BEGIN
            INSERT INTO [dbo].[FCL$Production BOM Line] (
              [Production BOM No_], 
              [Version Code],
              [Type],
              [Line No_], 
              [No_], 
              [Description],
              [Unit of Measure Code], 
              [Quantity], 
              [Position],
              [Position 2],
              [Position 3],
              [Lead-Time Offset], 
              [Routing Link Code],
              [Scrap _],
              [Variant Code],
              [Starting Date],
              [Ending Date],
              [Length],
              [Width],
              [Weight],
              [Depth],
              [Calculation Formula],
              [Quantity per]
            ) VALUES (
              @ProductionBOMNo,
              @VersionCode,
              @Type,
              @LineNo, 
              @No_, 
              @IntakeItemDescription,
              @UnitOfMeasureCode, 
              @Quantity,
              @Position, 
              @Position2,
              @Position3,
              @LeadTimeOffset,
              @RoutingLinkCode,
              @Scrap_,
              @VariantCode,
              @StartingDate,
              @EndingDate,
              @Length,
              @Width,
              @Weight,
              @Depth,
              @CalculationFormula,
              @QuantityPer
            )
            END
          `);
        const LocationCodeValue = LocationCode || ''; // Use default if LocationCode is missing or null
        const AutoAccummulateValue = AutoAccummulate || 0; // Convert "Yes" to 1, "No" to 0
          // Insert into FCL$BOMSwapLines
          await pool.request()
            .input("BOMNo", sql.VarChar, BOMNo)
            .input("ItemNo", sql.VarChar, ItemNo)
            .input("LocationCode", sql.VarChar, LocationCodeValue)
            .input("ShortCode", sql.VarChar, '')
            .input("AutoAccummulate", sql.Bit, AutoAccummulateValue)
            .input("UnitsPer100", sql.Float, UnitsPer100)
            .input("LinkedToInput", sql.Int, 0)
            .input("LinkedToOutput", sql.Int, 1)
            .input("DefaultQty", sql.Float, 0)
            .input("PiecesMandatory", sql.Int, 0)
            .input("linkedBOMOrFamily", sql.VarChar, '')
            .query(`
              IF NOT EXISTS (
              SELECT 1 FROM [dbo].[FCL$BOMSwapLines] 
              WHERE [BOMNo] = @BOMNo AND [ItemNo] = @ItemNo
            )
              BEGIN
              INSERT INTO [dbo].[FCL$BOMSwapLines] (
                [BOMNo], [ItemNo], [Location Code], [Short Code], [AutoAccummulate],
                [Units Per 100], [LinkedToInput], [LinkedToOutput], [Default Qty], [Pieces Mandatory],
            [Linked BOM or Family]
              ) VALUES (
                @BOMNo, @ItemNo, @LocationCode, @ShortCode, @AutoAccummulate,
                @UnitsPer100, @LinkedToInput, @LinkedToOutput, @DefaultQty, @PiecesMandatory,@linkedBOMOrFamily
              )
              END
            `);
          
        } else if (Comments === "Item Removed") {
          // Delete from FCL$Production BOM Line
          await pool.request()
            .input("ProductionBOMNo", sql.VarChar, BOMNo)
            .input("ItemNo", sql.VarChar, ItemNo)
            .query(`
              DELETE FROM [dbo].[FCL$Production BOM Line]
              WHERE [Production BOM No_] = @ProductionBOMNo AND [No_] = @ItemNo
            `);
  
          // Delete from FCL$BOMSwapLines
          await pool.request()
            .input("BOMNo", sql.VarChar, BOMNo)
            .input("ItemNo", sql.VarChar, ItemNo)
            .query(`
              DELETE FROM [dbo].[FCL$BOMSwapLines]
              WHERE [BOMNo] = @BOMNo AND [ItemNo] = @ItemNo
            `);
        }
      }
  
      console.log("Data processing complete.");
    } catch (err) {
      console.error("SQL error:", err);
    } finally {
      sql.close(); // Close the database connection
    }
  };

// Main function to process all files in the input directory
const processAllFiles = async () => {
  try {
    const files = fs.readdirSync(inputDir).filter(file => file.endsWith('.xlsx'));
    for (const file of files) {
      const filePath = path.join(inputDir, file);
      console.log(`Processing file: ${file}`);
      await processFile(filePath);
    }
    console.log("All files processed.");
  } catch (err) {
    console.error("Error processing files:", err);
  }
};

// Execute the main function
processAllFiles();
