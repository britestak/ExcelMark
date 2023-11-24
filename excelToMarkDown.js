//Required libraries
const xlsx = require('xlsx');
const fs = require('fs');

// Method to convert array to markdown table
function arrayToMarkdownTable(array) {
    let markdown = '';
    array.forEach((row, index) => {
        markdown += '| ' + row.join(' | ') + ' |\n';
        if (index === 0) {
            markdown += '|' + row.map(() => '---').join('|') + '|\n';
        }
    });
    return markdown;
}

// Method to read Excel and convert to markdown
function excelToMarkdown(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 }); //Considering header as in the row 1
    return arrayToMarkdownTable(data);
}

// Kindly replace the required file in 'sampleExcel.xlsx'
const markdownTable = excelToMarkdown('sampleExcel.xlsx');


// Output to console or save to file
console.log(markdownTable); // logging the command prompt just for reference
fs.writeFileSync('outputMarkDown.md', markdownTable); // 
