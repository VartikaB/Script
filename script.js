const XLSX = require('xlsx');
const fs = require('fs');

// Read the JSON file
const jsonData = fs.readFileSync('input.json', 'utf8');

// Parse the JSON data
const data = JSON.parse(jsonData);

// Create a new workbook
const workbook = XLSX.utils.book_new();

// Function to create a sheet from data and add it to the workbook
function createSheetFromData(data, sheetName) {
  const worksheet = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
}

// Function to recursively process nested data and create sheets
function processNestedData(obj, parentSheetName = '') {
  for (const key in obj) {
    if (obj.hasOwnProperty(key)) {
      const value = obj[key];
      const valueType = typeof value;

      if (valueType === 'object' && value !== null) {
        const sheetName = `${parentSheetName}_${key}`;
        createSheetFromData([value], sheetName);

        if (parentSheetName !== '') {
          // Create a hyperlink from parent sheet to child sheet
          const parentSheet = workbook.Sheets[parentSheetName];
          parentSheet[`${key}_link`] = {
            t: 's',
            v: sheetName,
            l: {
              Target: `'${sheetName}'!A1`, // Link to the child sheet
              Tooltip: 'Click to navigate',
              TargetMode: 'Internal',
            },
          };
        }

        processNestedData(value, sheetName);
      }
    }
  }
}

// Create a sheet for the main JSON object
createSheetFromData([data], 'MainSheet');
processNestedData(data);

// Write the workbook to a file
XLSX.writeFile(workbook, 'nested_data.xlsx');







