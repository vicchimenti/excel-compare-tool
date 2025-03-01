// This file serves as an adapter between the original CLI tool and our Electron UI
const { program } = require('commander');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

// Function to compare Excel files
async function compareExcel(file1Path, file2Path, options = {}) {
  // Validate inputs
  if (!file1Path || !file2Path) {
    throw new Error('Two Excel files must be provided for comparison');
  }

  if (!fs.existsSync(file1Path)) {
    throw new Error(`File not found: ${file1Path}`);
  }

  if (!fs.existsSync(file2Path)) {
    throw new Error(`File not found: ${file2Path}`);
  }

  console.log(`Comparing: ${file1Path} and ${file2Path}`);

  try {
    // Read workbooks
    const workbook1 = XLSX.readFile(file1Path);
    const workbook2 = XLSX.readFile(file2Path);

    // Track differences
    const differences = [];

    // Get sheets from both workbooks
    const sheets1 = workbook1.SheetNames;
    const sheets2 = workbook2.SheetNames;

    // Compare sheet names
    const commonSheets = sheets1.filter(sheet => sheets2.includes(sheet));
    
    // For each common sheet, compare cell by cell
    for (const sheetName of commonSheets) {
      const sheet1 = workbook1.Sheets[sheetName];
      const sheet2 = workbook2.Sheets[sheetName];
      
      // Convert sheets to JSON for easier handling
      const data1 = XLSX.utils.sheet_to_json(sheet1, { header: 1, defval: null });
      const data2 = XLSX.utils.sheet_to_json(sheet2, { header: 1, defval: null });
      
      // If we have options.keyColumn, use that for matching rows
      if (options.keyColumn !== undefined && data1.length > 0 && data2.length > 0) {
        // Assume first row contains headers
        const headers1 = data1[0];
        const headers2 = data2[0];
        
        // Find the index of the key column in each file
        let keyIndex1 = -1;
        let keyIndex2 = -1;
        
        if (typeof options.keyColumn === 'number') {
          // If keyColumn is a number, use it directly as index
          keyIndex1 = options.keyColumn;
          keyIndex2 = options.keyColumn;
        } else if (typeof options.keyColumn === 'string') {
          // If keyColumn is a string, find the matching header
          keyIndex1 = headers1.findIndex(h => h === options.keyColumn);
          keyIndex2 = headers2.findIndex(h => h === options.keyColumn);
        }
        
        // Check if we found the key column in both files
        if (keyIndex1 >= 0 && keyIndex2 >= 0) {
          // Create maps of rows keyed by the key column value
          const rowMap1 = new Map();
          const rowMap2 = new Map();
          
          // Skip header row (index 0)
          for (let i = 1; i < data1.length; i++) {
            const row = data1[i];
            if (row && row[keyIndex1] !== undefined && row[keyIndex1] !== null) {
              // Convert key to string to ensure consistent matching
              const key = String(row[keyIndex1]);
              rowMap1.set(key, { row, index: i });
            }
          }
          
          for (let i = 1; i < data2.length; i++) {
            const row = data2[i];
            if (row && row[keyIndex2] !== undefined && row[keyIndex2] !== null) {
              // Convert key to string to ensure consistent matching
              const key = String(row[keyIndex2]);
              rowMap2.set(key, { row, index: i });
            }
          }
          
          // Find keys that exist in both maps for comparison
          const commonKeys = [...rowMap1.keys()].filter(key => rowMap2.has(key));
          
          // Find keys that exist only in one file
          const onlyInFile1 = [...rowMap1.keys()].filter(key => !rowMap2.has(key));
          const onlyInFile2 = [...rowMap2.keys()].filter(key => !rowMap1.has(key));
          
          // Compare rows with common keys
          for (const key of commonKeys) {
            const { row: row1, index: rowIndex1 } = rowMap1.get(key);
            const { row: row2, index: rowIndex2 } = rowMap2.get(key);
            
            // Get max columns to ensure we check all cells
            const maxCols = Math.max(
              row1 ? row1.length : 0,
              row2 ? row2.length : 0
            );
            
            // Compare each cell in the matched rows
            for (let col = 0; col < maxCols; col++) {
              const value1 = row1 && col < row1.length ? row1[col] : null;
              const value2 = row2 && col < row2.length ? row2[col] : null;
              
              // Check if values are different (accounting for null/undefined)
              if (value1 !== value2 && !(value1 === null && value2 === undefined) && !(value1 === undefined && value2 === null)) {
                // Convert column index to Excel column letter
                const colLetter = XLSX.utils.encode_col(col);
                const cellRef1 = `${colLetter}${rowIndex1 + 1}`;
                const cellRef2 = `${colLetter}${rowIndex2 + 1}`;
                const headerName = headers1[col] || `Column ${colLetter}`;
                
                differences.push({
                  sheet: sheetName,
                  keyValue: key,
                  cell1: cellRef1,
                  cell2: cellRef2,
                  column: headerName,
                  value1: value1,
                  value2: value2
                });
              }
            }
          }
          
          // Add rows that only exist in file 1
          for (const key of onlyInFile1) {
            const { index: rowIndex } = rowMap1.get(key);
            differences.push({
              sheet: sheetName,
              keyValue: key,
              cell1: `Row ${rowIndex + 1}`,
              cell2: 'N/A',
              column: 'Entire Row',
              value1: 'Row exists',
              value2: 'Row missing'
            });
          }
          
          // Add rows that only exist in file 2
          for (const key of onlyInFile2) {
            const { index: rowIndex } = rowMap2.get(key);
            differences.push({
              sheet: sheetName,
              keyValue: key,
              cell1: 'N/A',
              cell2: `Row ${rowIndex + 1}`,
              column: 'Entire Row',
              value1: 'Row missing',
              value2: 'Row exists'
            });
          }
        } else {
          // Key column not found in one or both files, fall back to direct comparison
          console.warn(`Key column "${options.keyColumn}" not found in one or both files for sheet "${sheetName}". Using direct comparison instead.`);
          compareDirectly(sheetName, data1, data2, differences);
        }
      } else {
        // No key column specified, use direct comparison
        compareDirectly(sheetName, data1, data2, differences);
      }
    }
    
    // Check for sheets that exist in one file but not the other
    const uniqueSheets1 = sheets1.filter(sheet => !sheets2.includes(sheet));
    const uniqueSheets2 = sheets2.filter(sheet => !sheets1.includes(sheet));
    
    for (const sheet of uniqueSheets1) {
      differences.push({
        sheet: sheet,
        keyValue: 'N/A',
        cell1: 'N/A',
        cell2: 'N/A',
        column: 'Sheet',
        value1: 'Sheet exists',
        value2: 'Sheet missing'
      });
    }
    
    for (const sheet of uniqueSheets2) {
      differences.push({
        sheet: sheet,
        keyValue: 'N/A',
        cell1: 'N/A',
        cell2: 'N/A',
        column: 'Sheet',
        value1: 'Sheet missing',
        value2: 'Sheet exists'
      });
    }
    
    // Get column names from the first sheet for UI display
    let columns = [];
    if (commonSheets.length > 0) {
      const firstSheetName = commonSheets[0];
      const sheet1 = workbook1.Sheets[firstSheetName];
      const data = XLSX.utils.sheet_to_json(sheet1, { header: 1, defval: null });
      if (data.length > 0) {
        columns = data[0].map((col, index) => ({
          name: col || `Column ${XLSX.utils.encode_col(index)}`,
          index: index
        }));
      }
    }
    
    return {
      differences,
      file1: path.basename(file1Path),
      file2: path.basename(file2Path),
      columns: columns
    };
  } catch (error) {
    console.error('Error comparing Excel files:', error);
    throw new Error(`Failed to compare Excel files: ${error.message}`);
  }
}

// Helper function for direct comparison (original logic)
function compareDirectly(sheetName, data1, data2, differences) {
  // Get max dimensions to ensure we check all cells
  const maxRows = Math.max(data1.length, data2.length);
  const maxCols = Math.max(
    ...data1.map(row => row ? row.length : 0),
    ...data2.map(row => row ? row.length : 0)
  );
  
  // Compare each cell
  for (let row = 0; row < maxRows; row++) {
    for (let col = 0; col < maxCols; col++) {
      const value1 = data1[row] && col < data1[row].length ? data1[row][col] : null;
      const value2 = data2[row] && col < data2[row].length ? data2[row][col] : null;
      
      // Check if values are different (accounting for null/undefined)
      if (value1 !== value2 && !(value1 === null && value2 === undefined) && !(value1 === undefined && value2 === null)) {
        // Convert column index to Excel column letter
        const colLetter = XLSX.utils.encode_col(col);
        const cellRef = `${colLetter}${row + 1}`;
        
        // Try to get header name from first row if available
        let headerName = `Column ${colLetter}`;
        if (data1.length > 0 && data1[0] && col < data1[0].length) {
          headerName = data1[0][col] || headerName;
        }
        
        differences.push({
          sheet: sheetName,
          keyValue: 'N/A',
          cell1: cellRef,
          cell2: cellRef,
          column: headerName,
          value1: value1,
          value2: value2
        });
      }
    }
  }
}

// If this file is run directly from command line
if (require.main === module) {
  program
    .arguments('<oldFile> <newFile>')
    .description('Compare two Excel files')
    .option('-k, --key-column <column>', 'Column name or index to use as key for matching rows')
    .action(async (oldFile, newFile, cmdOptions) => {
      try {
        const options = {};
        if (cmdOptions.keyColumn) {
          // Try to convert to number if it's a numeric string
          const numericKey = Number(cmdOptions.keyColumn);
          options.keyColumn = isNaN(numericKey) ? cmdOptions.keyColumn : numericKey;
        }
        
        const results = await compareExcel(oldFile, newFile, options);
        
        console.log('\nComparison Results:');
        console.log(`Comparing ${results.file1} and ${results.file2}`);
        
        if (results.differences.length === 0) {
          console.log('No differences found. Files are identical.');
        } else {
          console.log(`Found ${results.differences.length} differences:`);
          
          results.differences.forEach((diff, index) => {
            console.log(`\nDifference #${index + 1}:`);
            console.log(`Sheet: ${diff.sheet}`);
            console.log(`Key Value: ${diff.keyValue}`);
            console.log(`Column: ${diff.column}`);
            console.log(`Cell in File 1: ${diff.cell1}`);
            console.log(`Cell in File 2: ${diff.cell2}`);
            console.log(`File 1 value: ${diff.value1}`);
            console.log(`File 2 value: ${diff.value2}`);
          });
        }
      } catch (error) {
        console.error('Error:', error.message);
        process.exit(1);
      }
    })
    .parse(process.argv);
  
  // Show help if no arguments provided
  if (program.args.length === 0) {
    program.help();
  }
} else {
  // When imported as a module, export the compare function
  module.exports = compareExcel;
}