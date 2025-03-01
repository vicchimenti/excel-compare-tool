// This file serves as an adapter between the original CLI tool and our Electron UI
const { program } = require('commander');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

// Function to compare Excel files
async function compareExcel(file1Path, file2Path) {
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
            
            differences.push({
              sheet: sheetName,
              cell: cellRef,
              value1: value1,
              value2: value2
            });
          }
        }
      }
    }
    
    // Check for sheets that exist in one file but not the other
    const uniqueSheets1 = sheets1.filter(sheet => !sheets2.includes(sheet));
    const uniqueSheets2 = sheets2.filter(sheet => !sheets1.includes(sheet));
    
    for (const sheet of uniqueSheets1) {
      differences.push({
        sheet: sheet,
        cell: 'N/A',
        value1: 'Sheet exists',
        value2: 'Sheet missing'
      });
    }
    
    for (const sheet of uniqueSheets2) {
      differences.push({
        sheet: sheet,
        cell: 'N/A',
        value1: 'Sheet missing',
        value2: 'Sheet exists'
      });
    }
    
    return {
      differences,
      file1: path.basename(file1Path),
      file2: path.basename(file2Path)
    };
  } catch (error) {
    console.error('Error comparing Excel files:', error);
    throw new Error(`Failed to compare Excel files: ${error.message}`);
  }
}

// If this file is run directly from command line
if (require.main === module) {
  program
    .arguments('<oldFile> <newFile>')
    .description('Compare two Excel files')
    .action(async (oldFile, newFile) => {
      try {
        const results = await compareExcel(oldFile, newFile);
        
        console.log('\nComparison Results:');
        console.log(`Comparing ${results.file1} and ${results.file2}`);
        
        if (results.differences.length === 0) {
          console.log('No differences found. Files are identical.');
        } else {
          console.log(`Found ${results.differences.length} differences:`);
          
          results.differences.forEach((diff, index) => {
            console.log(`\nDifference #${index + 1}:`);
            console.log(`Sheet: ${diff.sheet}`);
            console.log(`Cell: ${diff.cell}`);
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