#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { program } = require('commander');
const chalk = require('chalk');

program
  .name('excel-compare')
  .description('Compare two Excel files and identify differences')
  .version('1.0.0')
  .argument('<oldFile>', 'Path to the old Excel file')
  .argument('<newFile>', 'Path to the new Excel file')
  .option('-s, --sheet <name>', 'Specific sheet name to compare (defaults to first sheet)')
  .option('-i, --id-column <name>', 'Column name to use as identifier (defaults to "ID")')
  .option('-o, --output <file>', 'Save output to a file')
  .option('--json', 'Output results as JSON')
  .parse(process.argv);

const options = program.opts();
const [oldFilePath, newFilePath] = program.args;

// Save console.log output if output file is specified
if (options.output) {
  const logFile = fs.createWriteStream(options.output);
  const originalConsoleLog = console.log;
  console.log = function() {
    originalConsoleLog.apply(console, arguments);
    const text = Array.from(arguments).join(' ');
    logFile.write(text + '\n');
  };
}

// Main function to compare Excel files
async function compareExcelFiles(oldFilePath, newFilePath, options) {
  try {
    // Read the Excel files
    const oldFile = fs.readFileSync(oldFilePath);
    const newFile = fs.readFileSync(newFilePath);

    // Import SheetJS
    const oldWorkbook = XLSX.read(oldFile, {
      cellDates: true
    });

    const newWorkbook = XLSX.read(newFile, {
      cellDates: true
    });

    console.log(chalk.blue("Old workbook sheets:"), oldWorkbook.SheetNames);
    console.log(chalk.blue("New workbook sheets:"), newWorkbook.SheetNames);

    // Determine which sheet to use
    const sheetName = options.sheet || oldWorkbook.SheetNames[0];
    if (!oldWorkbook.SheetNames.includes(sheetName)) {
      console.error(chalk.red(`Sheet "${sheetName}" not found in old file`));
      return;
    }
    if (!newWorkbook.SheetNames.includes(sheetName)) {
      console.error(chalk.red(`Sheet "${sheetName}" not found in new file`));
      return;
    }

    // Get worksheets
    const oldSheet = oldWorkbook.Sheets[sheetName];
    const newSheet = newWorkbook.Sheets[sheetName];

    // Convert to array of arrays for easier processing
    const oldData = XLSX.utils.sheet_to_json(oldSheet, { header: 1 });
    const newData = XLSX.utils.sheet_to_json(newSheet, { header: 1 });

    console.log(chalk.blue(`Old sheet "${sheetName}" has ${oldData.length} rows`));
    console.log(chalk.blue(`New sheet "${sheetName}" has ${newData.length} rows`));

    // Get headers
    const oldHeaders = oldData[0];
    const newHeaders = newData[0];

    // Find indexes of key identifying fields
    const findColumnIndex = (headers, name) => {
      return headers.findIndex(h => h && h.toString().includes(name));
    };

    // Find the ID column (use the one specified in options or look for "ID")
    const idColumnName = options.idColumn || "ID";
    const oldIdCol = findColumnIndex(oldHeaders, idColumnName);
    const newIdCol = findColumnIndex(newHeaders, idColumnName);

    if (oldIdCol === -1 || newIdCol === -1) {
      console.error(chalk.red(`Could not find column "${idColumnName}" in one or both files.`));
      console.log(chalk.yellow("Available columns in old file:"), oldHeaders);
      console.log(chalk.yellow("Available columns in new file:"), newHeaders);
      return;
    }

    // Try to find other useful columns
    const oldNameCol = findColumnIndex(oldHeaders, "Name");
    const newNameCol = findColumnIndex(newHeaders, "Name");
    const oldTitleCol = findColumnIndex(oldHeaders, "Title");
    const newTitleCol = findColumnIndex(newHeaders, "Title");

    console.log(chalk.blue(`Old file column indexes - ID: ${oldIdCol}, Name: ${oldNameCol}, Title: ${oldTitleCol}`));
    console.log(chalk.blue(`New file column indexes - ID: ${newIdCol}, Name: ${newNameCol}, Title: ${newTitleCol}`));

    // Create sets of IDs for quick lookup
    const oldIds = new Set(oldData.slice(1).map(row => String(row[oldIdCol])));
    const newIds = new Set(newData.slice(1).map(row => String(row[newIdCol])));

    // Find rows present only in old file (removed)
    const removedIds = [...oldIds].filter(id => !newIds.has(id));
    console.log(chalk.red(`Found ${removedIds.length} rows present only in old file (removed)`));

    // Find rows present only in new file (added)
    const addedIds = [...newIds].filter(id => !oldIds.has(id));
    console.log(chalk.green(`Found ${addedIds.length} rows present only in new file (added)`));

    // Find common IDs (potential modifications)
    const commonIds = [...oldIds].filter(id => newIds.has(id));
    console.log(chalk.blue(`Found ${commonIds.length} IDs present in both files (potential modifications)`));

    // Function to create a map of rows by ID for quick lookup
    function createRowMap(data, idCol) {
      const map = new Map();
      for (let i = 1; i < data.length; i++) {
        const id = String(data[i][idCol]);
        if (id) map.set(id, data[i]);
      }
      return map;
    }

    const oldRowsById = createRowMap(oldData, oldIdCol);
    const newRowsById = createRowMap(newData, newIdCol);

    // Function to compare two rows and identify differences
    function findDifferences(oldRow, newRow, oldHeaders, newHeaders) {
      const diffs = [];
      const maxCols = Math.max(oldRow.length, newRow.length);

      for (let i = 0; i < maxCols; i++) {
        // Skip if the column doesn't exist in one of the files
        if (i >= oldRow.length || i >= newRow.length) continue;

        const oldVal = oldRow[i] === undefined || oldRow[i] === null ? "" : String(oldRow[i]);
        const newVal = newRow[i] === undefined || newRow[i] === null ? "" : String(newRow[i]);

        if (oldVal !== newVal) {
          diffs.push({
            columnIndex: i,
            columnName: i < oldHeaders.length ? oldHeaders[i] : `Column ${i}`,
            oldValue: oldVal,
            newValue: newVal
          });
        }
      }

      return diffs;
    }

    // Check for modifications in common rows
    const modifiedRows = [];

    for (const id of commonIds) {
      const oldRow = oldRowsById.get(id);
      const newRow = newRowsById.get(id);

      if (oldRow && newRow) {
        const differences = findDifferences(oldRow, newRow, oldHeaders, newHeaders);

        if (differences.length > 0) {
          modifiedRows.push({
            id,
            name: newRow[newNameCol] || oldRow[oldNameCol],
            title: newRow[newTitleCol] || oldRow[oldTitleCol],
            differences
          });
        }
      }
    }

    console.log(chalk.yellow(`Found ${modifiedRows.length} modified rows`));

    // Prepare results for potential JSON output
    const results = {
      addedRows: [],
      removedRows: [],
      modifiedRows: []
    };

    // Output detailed information about changes
    // 1. Added rows
    if (addedIds.length > 0) {
      console.log(chalk.green("\nADDED ROWS:"));
      for (const id of addedIds) {
        const row = newRowsById.get(id);
        if (row) {
          const name = row[newNameCol] || "Unknown";
          const title = row[newTitleCol] || "Unknown";
          console.log(chalk.green(`ID: ${id}, Name: ${name}, Title: ${title}`));
          
          results.addedRows.push({ id, name, title });
        }
      }
    }

    // 2. Removed rows
    if (removedIds.length > 0) {
      console.log(chalk.red("\nREMOVED ROWS:"));
      for (const id of removedIds) {
        const row = oldRowsById.get(id);
        if (row) {
          const name = row[oldNameCol] || "Unknown";
          const title = row[oldTitleCol] || "Unknown";
          console.log(chalk.red(`ID: ${id}, Name: ${name}, Title: ${title}`));
          
          results.removedRows.push({ id, name, title });
        }
      }
    }

    // 3. Modified rows
    if (modifiedRows.length > 0) {
      console.log(chalk.yellow("\nMODIFIED ROWS:"));

      for (const mod of modifiedRows) {
        console.log(chalk.yellow(`\nID: ${mod.id}, Name: ${mod.name || 'Unknown'}`));
        console.log(chalk.yellow(`Title: ${mod.title || 'Unknown'}`));
        console.log(chalk.yellow(`Changes (${mod.differences.length}):`));

        const modifiedRowData = {
          id: mod.id,
          name: mod.name,
          title: mod.title,
          changes: []
        };

        for (const diff of mod.differences) {
          // Truncate long values for display
          let oldVal = diff.oldValue;
          let newVal = diff.newValue;

          if (oldVal && oldVal.length > 100) oldVal = oldVal.substring(0, 97) + "...";
          if (newVal && newVal.length > 100) newVal = newVal.substring(0, 97) + "...";

          console.log(`  - ${diff.columnName}: ${chalk.red(`"${oldVal}"`)} → ${chalk.green(`"${newVal}"`)}`);
          
          modifiedRowData.changes.push({
            field: diff.columnName,
            oldValue: diff.oldValue,
            newValue: diff.newValue
          });
        }
        
        results.modifiedRows.push(modifiedRowData);
      }
    }

    // Summary of changes
    console.log(chalk.blue("\n\nSUMMARY OF CHANGES:"));
    console.log(chalk.blue(`- Total rows in old file: ${oldData.length - 1}`));
    console.log(chalk.blue(`- Total rows in new file: ${newData.length - 1}`));
    console.log(chalk.green(`- Added rows: ${addedIds.length}`));
    console.log(chalk.red(`- Removed rows: ${removedIds.length}`));
    console.log(chalk.yellow(`- Modified rows: ${modifiedRows.length}`));

    // Output as JSON if requested
    if (options.json) {
      console.log("\nJSON OUTPUT:");
      console.log(JSON.stringify(results, null, 2));
    }

  } catch (error) {
    console.error(chalk.red('Error processing files:'), error);
    process.exit(1);
  }
}

// Run the comparison
compareExcelFiles(oldFilePath, newFilePath, options);