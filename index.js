import xlsx from 'xlsx';
import fs from 'fs';
import sql from 'mssql';
import dbConfig from './config/dbConfig.js';
import path from 'path';
import { fileURLToPath } from 'url';

// Create the equivalent of __dirname
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Load the Excel file using the path with __dirname
const workbook = xlsx.readFile(path.resolve(__dirname, "../updateBOM/Files/otherChopping.xlsx"));
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert the worksheet to JSON
const data = xlsx.utils.sheet_to_json(worksheet);

// Filter out blank rows and rows without an ItemNo
const filteredData = data.filter(row => row.ItemNo && row.ItemNo.trim() !== "");

// Save filtered data to a JSON file for preview
fs.writeFile("filteredData.json", JSON.stringify(filteredData, null, 2), (err) => {
  if (err) {
    console.error("Error writing to file:", err);
  } else {
    console.log("Filtered data saved to 'filteredData.json'. You can review it before inserting into the database.");
  }
});

// Function to insert or delete data in SQL Server
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
        UnitsPer100,
        LocationCode,
        AutoAccummulate,
        Comments,
        OutputItem,
        OutputItemDescription
      } = row;

      if (Comments === "Item Added") {
        // Insert into FCL$Production BOM Line
        await pool.request()
          .input("ProductionBOMNo", sql.VarChar, BOMNo)
          .input("VersionCode", sql.Int, 0)  
          .input("Type", sql.Int, 1)
          .input("LineNo", sql.Int, lineNo)
          .input("No_", sql.VarChar, ItemNo)
          .input("Description", sql.VarChar, IntakeItemDescription)
          .input("UnitOfMeasureCode", sql.VarChar, BaseUOM)
          .input("Quantity", sql.Float, UsagePerBatch)
          .input("Position", sql.VarChar, '')
          .input("Position2", sql.VarChar, '')
          .input("Position3", sql.VarChar, '')
          .input("LeadTimeOffset", sql.VarChar, '')
          .input("RoutingLinkCode", sql.VarChar, '')
          .input("Scrap_", sql.Float, 0)
          .input("VariantCode", sql.VarChar, '')
          .input("LinkedBOMOrFamily", sql.VarChar, OutputItem)
          .input("StartingDate", sql.Date, '1753-01-01')
          .input("EndingDate", sql.Date, '1753-01-01')
          .input("Length", sql.Float, 0)
          .input("Width", sql.Float, 0)
          .input("Weight", sql.Float, 0)
          .input("Depth", sql.Float, 0)
          .input("CalculationFormula", sql.VarChar, '')
          .input("QuantityPer", sql.Float, UsagePerBatch)
          .query(`
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
              [Linked BOM or Family],
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
              @Description,
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
          `);

        // Insert into FCL$BOMSwapLines
        await pool.request()
          .input("BOMNo", sql.VarChar, BOMNo)
          .input("ItemNo", sql.VarChar, ItemNo)
          .input("LocationCode", sql.VarChar, LocationCode)
          .input("ShortCode", sql.VarChar, '')
          .input("AutoAccummulate", sql.Bit, AutoAccummulate)
          .input("UnitsPer100", sql.Float, UnitsPer100)
          .input("LinkedToInput", sql.VarChar, 0)
          .input("LinkedToOutput", sql.VarChar, 1)
          .input("DefaultQty", sql.VarChar, 0)
          .input("PiecesMandatory", sql.VarChar, 0)
          .query(`
            INSERT INTO [dbo].[FCL$BOMSwapLines] (
              [BOMNo], [ItemNo], [Location Code], [ShortCode], [AutoAccummulate],
              [Units Per 100], [LinkedToInput], [LinkedToOutput], [Default Qty], [Pieces Mandatory]
            ) VALUES (
              @BOMNo, @ItemNo, @LocationCode, @ShortCode, @AutoAccummulate,
              @UnitsPer100, @LinkedToInput, @LinkedToOutput, @DefaultQty, @PiecesMandatory
            )
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

// Execute the data processing function with filtered data
processData(filteredData);
