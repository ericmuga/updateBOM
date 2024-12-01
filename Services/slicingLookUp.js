import fs from 'fs';
import sql from 'mssql';
import dbConfig from '../config/dbConfig.js';

// Helper function to connect to the database
const connectToDatabase = async () => {
  try {
    return await sql.connect(dbConfig);
  } catch (error) {
    console.error('Error connecting to the database:', error);
    throw error;
  }
};

// Function to generate and save the lookup
const generateAndSaveLookup = async () => {
  const query = `
    SELECT 
      [Item No_] AS [Output],
      [FamilyNo],
      (SELECT [Intake Item] 
       FROM [FCL$Family BOM Swap Header] AS b 
       WHERE a.[FamilyNo] = b.[Family No_]) AS [Intake]
    FROM [FCL$FamilyBOMSwaps] AS a 
    WHERE [FamilyNo] IN (
      '1220H00','1220H01','1220H02','1220H03','1220H04','1220H05','1220H06','1220H07',
      '1220H08','1220H09','1220H10','1220H11','1220H12','1220H14','1220H15','1220H16',
      '1220H17','1220H18','1220H19','1220H20','1220H21','1220H22','1220H23','1220H24',
      '1220H25','1220H26','1220H27','1220H28','1220H29','1220H31','1220H32','1220H33',
      '1220H34','1220H35','1220H36','1220H37','1220H38','1220H39','1220H40','1220H41',
      '1220H42','1220H43','1220H44','1220H45','1220H46','1220H47','1220H48','1220H49',
      '1220H50','1220H51','1220H52','1220H53','1220H54','1220H55','1220H56','1220H57',
      '1220H58','1220H59','1220H60','1220H61','1220H62','1220H63','1220H64','1220H65',
      '1220H66','1220H67','1220H68','1220H69','1220H70','1220H71','1220H72','1220H73',
      '1220H74','1220H75','1220H76','1220H77','1220H78','1220H79','1220H80','1220H81',
      '1220H82','1220H83','1220H84','1220H85','1220H86','1220H87','1220H88','1220H89',
      '1220H90','1220H91','1220H92','1220H93','1220H94','1220H95','1220H96','1220H97',
      '1220H98','1220H99','1220I00','1220I01','1220I02','1220I03','1220I04','1220I05',
      '1220I06','1220I07','1220I08','1220I09','1220I10','1220I11','1220I12','1220I13'
    )
    AND [Percentage of Main Prod] > 0
  `;

  try {
    // Connect to the database
    const pool = await connectToDatabase();

    // Execute the query
    const result = await pool.request().query(query);

    // Transform the result into the desired lookup format
    const lookup = result.recordset.map((row, index) => ({
      process_code: 8, // Process code for slicing
      shortcode: "SL",
      process_name: "Slicing parts for slices, portion",
      intake_item: row.Intake,
      output_item: row.Output,
      input_location: "1570",
      output_location: "1570",
      production_order_series: `P${String(8).padStart(2, "0")}`,
      process_loss: 0.00,
    }));

    // Convert the lookup to a JS module
    const lookupFileContent = `const processLookup = ${JSON.stringify(
      lookup,
      null,
      2
    )};\n\nexport default processLookup;`;

    // Save the lookup as a JS file
    const filePath = './processLookup.js';
    fs.writeFileSync(filePath, lookupFileContent, 'utf8');

    console.log(`Lookup saved successfully to ${filePath}`);
  } catch (error) {
    console.error('Error generating and saving lookup:', error);
    throw error;
  }
};

// Call the function
generateAndSaveLookup();
