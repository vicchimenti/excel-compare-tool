# Excel Comparison Tool

A command-line utility for comparing two Excel files to identify differences. This tool helps you quickly spot added rows, removed rows, and modifications between two versions of an Excel file.

## Features

- **Row-based comparison** using a unique identifier column
- **Detailed change reporting** for added, removed, and modified rows
- **Customizable comparison** with options for different sheets and ID columns
- **Colorized output** for better readability in the terminal
- **Multiple output formats** including text and JSON
- **Support for output redirection** to save results to a file

## Installation

### Prerequisites

- [Node.js](https://nodejs.org/) (version 12 or higher)

### Setup

1. Clone this repository:

   git clone https://github.com/vicchimenti/excel-compare-tool.git
   cd excel-compare-tool

2. Install dependencies:

   npm install

### Usage

#### Basic Comparison

node index.js path/to/old-file.xlsx path/to/new-file.xlsx

### Command Line Options

Usage: excel-compare [options] <oldFile> <newFile>

#### Arguments

  oldFile                       Path to the old Excel file
  newFile                       Path to the new Excel file

#### Options

  -V, --version                 output the version number
  -s, --sheet <name>            Specific sheet name to compare (defaults to first sheet)
  -i, --id-column <name>        Column name to use as identifier (defaults to "ID")
  -o, --output <file>           Save output to a file
  --json                        Output results as JSON
  -h, --help                    display help for command

### Examples

#### Compare specific sheets

node index.js old-data.xlsx new-data.xlsx --sheet "Sales Data"

#### Use a different identifier column

node index.js old-data.xlsx new-data.xlsx --id-column "Employee Number"

#### Save results to a file

node index.js old-data.xlsx new-data.xlsx --output comparison-results.txt

#### Get JSON output

node index.js old-data.xlsx new-data.xlsx --json > results.json

## How It Works

The tool follows these steps to compare Excel files:

Loads both Excel files using the SheetJS library
Identifies the sheet(s) to compare
Extracts rows and headers from each file
Uses the specified ID column to match corresponding rows between files
Identifies rows that:

Exist only in the old file (removed)
Exist only in the new file (added)
Exist in both files but have differences (modified)

For modified rows, provides a detailed comparison of changed values
Outputs a summary and details of all differences

### Use Cases

Track changes in data exports over time
Verify data migrations between systems
Compare different versions of a dataset
Validate changes after bulk updates
Quality assurance for data processing

### Troubleshooting

Common Issues

"Could not find column 'ID' in one or both files": Use the --id-column option to specify the correct identifier column name
Empty comparison results: Verify that the files contain matching data and the correct sheet name is specified
Error reading files: Check that the file paths are correct and the files are valid Excel format

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

### SheetJS

[https://github.com/SheetJS/sheetjs] for Excel file parsing

### Commander.js

[https://github.com/tj/commander.js/] for command line parsing

### Chalk

[https://github.com/chalk/chalk] for terminal colors
