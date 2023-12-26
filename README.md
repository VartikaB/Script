**Excel Sheet Generation from Nested JSON**
This script demonstrates the process of converting nested JSON data into an Excel spreadsheet using Node.js and the 'xlsx' library.

**Instructions:**
**Prerequisites:**


Install the required package by running **npm install xlsx** in the terminal.
Usage:

Place your input JSON file (input.json) containing nested data in the same directory as the script.
Run the Script:

Execute the script by running node script.js in the terminal.


**Code Explanation:**
Libraries Used:
xlsx: Used for Excel file manipulation.
fs: File system module for reading the input JSON file.
Code Breakdown:
**Read JSON Data:**

Read the contents of the input.json file and parse it into a JavaScript object.
**Create Workbook:**

Initialize a new Excel workbook using XLSX.utils.book_new().
**Sheet Creation Functions:**

createSheetFromData: Generates a new sheet from provided data and adds it to the workbook.
processNestedData: Recursively processes nested data and creates sheets for each level in the hierarchy.
**Conversion Process:**

The script iterates through the JSON data, generating sheets for each object.
For nested objects, it creates child sheets and establishes hyperlinks from parent sheets to their respective child sheets within the same Excel file.
**Output:**

The generated Excel file (nested_data.xlsx) contains the main sheet ('MainSheet') with hyperlinks to its child JSON object sheets, preserving the hierarchical structure.
**Running the Example:**
Place your JSON data in a file named input.json in the same directory as the script.
Run the script using node script.js.
Check the output file nested_data.xlsx to view the generated Excel sheet with linked data.
