const XLSX = require('xlsx');
const XLSXPopulate = require('xlsx-populate');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

async function downloadFile(fileId, destination) {
    try {
        const url = `https://drive.usercontent.google.com/download?id=${fileId}&export=download&confirm=t`;
        const response = await axios({
            url,
            method: 'GET',
            responseType: 'stream'
        });

        const writer = fs.createWriteStream(destination);
        response.data.pipe(writer);

        return new Promise((resolve, reject) => {
            writer.on('finish', () => {
                console.log(`File successfully downloaded to ${destination}`);
                resolve(true);
            });
            writer.on('error', reject);
        });
    } catch (error) {
        console.error('Error downloading file:', error.message);
        return false;
    }
}

async function updateExcelFile(filePath, sheetName, updates) {
    try {
        // Load the workbook while preserving styles
        const workbook = await XLSXPopulate.fromFileAsync(filePath);
        const worksheet = workbook.sheet(sheetName);

        if (!worksheet) {
            console.error(`Sheet ${sheetName} not found`);
            return false;
        }

        // Find ASIN column and row (case-insensitive)
        let asinCol = -1;
        let headerRow = -1;
        const asinVariations = ['ASIN','Asins'];
        
        // Search first 10 rows for headers
        for (let row = 1; row <= 10; row++) {
            for (let col = 1; col <= 50; col++) { // Search first 50 columns
                const cellValue = worksheet.cell(row, col).value();
                if (cellValue && asinVariations.includes(cellValue.toString().trim())) {
                    asinCol = col;
                    headerRow = row;
                    console.log(`Found ASIN column with name "${cellValue}" at column ${col}, row ${row}`);
                    break;
                }
            }
            if (asinCol !== -1) break;
        }

        if (asinCol === -1 || headerRow === -1) {
            console.error('Could not find ASIN column (tried variations: ' + asinVariations.join(', ') + ')');
            return false;
        }

        // Map column headers to their indices
        const colIndices = {};
        const headerValues = {}; // Store actual header names
        for (let col = 1; col <= 50; col++) {
            const header = worksheet.cell(headerRow, col).value();
            if (header) {
                const headerKey = header.toString().trim();
                colIndices[headerKey.toUpperCase()] = col; // Store uppercase version for case-insensitive matching
                headerValues[headerKey.toUpperCase()] = headerKey; // Store original header
            }
        }

        // Log available columns for debugging
        console.log('\nAvailable columns in sheet:', Object.values(headerValues));

        // Find the row with matching ASIN
        let targetRow = -1;
        const range = worksheet.usedRange();
        const maxRow = range.endCell().rowNumber();
        
        for (let row = headerRow + 1; row <= maxRow; row++) {
            const cellValue = worksheet.cell(row, asinCol).value();
            if (cellValue && cellValue.toString().trim().toUpperCase() === updates.asin.toUpperCase()) {
                targetRow = row;
                break;
            }
        }

        if (targetRow === -1) {
            console.error(`ASIN ${updates.asin} not found in the sheet`);
            return false;
        }

        console.log(`\nFound ASIN ${updates.asin} at row ${targetRow}`);

        // Update only the values that exist in the sheet
        let updatedFields = [];
        let skippedFields = [];
        
        Object.entries(updates).forEach(([field, value]) => {
            if (field.toLowerCase() !== 'asin') {
                // Try to find the column case-insensitively
                const upperField = field.toUpperCase();
                if (colIndices[upperField] !== undefined) {
                    worksheet.cell(targetRow, colIndices[upperField]).value(value);
                    updatedFields.push(`${headerValues[upperField]}: ${value}`);
                } else {
                    skippedFields.push(field);
                }
            }
        });

        if (updatedFields.length > 0) {
            console.log('\nUpdated fields:', updatedFields.join(', '));
        }
        if (skippedFields.length > 0) {
            console.log('\nSkipped fields (not found in sheet):', skippedFields.join(', '));
        }

        // Save the workbook
        await workbook.toFileAsync(filePath);
        console.log(`\nSaved updates to file: ${filePath}`);
        return true;

    } catch (error) {
        console.error('Error processing Excel file:', error.message);
        return false;
    }
}

// Example usage:
async function main() {
    const fileId = '11xfaf0nGgscOpYpE9U3tfRelyCxsDoOJ';
    const excelFile = 'ICERWORKSHEET.xlsx';

    // Example updates for different sheets
    const updates = {
        'AMS NFL': {
            asin: 'B084TRRKBY',
            'FBA INV': 8,
            'OTS QOH': 1,
            'QOHQTY': 456,
            'AMZ VC INV': 789,
            'WIP QTY': 10,
            'WIP ETA': '2024-07-01'
        },
        'AMS NBA': {
            asin: 'B01LZZHGGM',
            'AMZ VC INV': 200
        },
        'AMS WNBA per Size': {
            asin: 'B0DPJDRKKD',
            'DF INV': 6,
            'AMZ VC': 150
        }
    };

    // Download the file first
    const downloaded = await downloadFile(fileId, excelFile);
    if (!downloaded) {
        console.error('Failed to download file');
        return;
    }

    // Update each sheet
    for (const [sheetName, sheetUpdates] of Object.entries(updates)) {
        console.log(`\n=== Processing sheet: ${sheetName} ===`);
        await updateExcelFile(excelFile, sheetName, sheetUpdates);
    }
}

// If running directly (not as a module)
if (require.main === module) {
    main().catch(console.error);
}

// Export functions for use as a module
module.exports = {
    downloadFile,
    updateExcelFile
}; 