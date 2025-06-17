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

async function updateExcelFile(filePath, updates) {
    try {
        // Load the workbook while preserving styles
        const workbook = await XLSXPopulate.fromFileAsync(filePath);
        const worksheet = workbook.sheet('AMS NFL');

        if (!worksheet) {
            console.error('Sheet AMS NFL not found');
            return false;
        }

        // Find ASIN column and row
        let asinCol = -1;
        let headerRow = -1;
        
        // Search first 10 rows for headers (adjust if needed)
        for (let row = 1; row <= 10; row++) {
            for (let col = 1; col <= 20; col++) { // Search first 20 columns
                const cellValue = worksheet.cell(row, col).value();
                if (cellValue === 'ASIN') {
                    asinCol = col;
                    headerRow = row;
                    break;
                }
            }
            if (asinCol !== -1) break;
        }

        if (asinCol === -1 || headerRow === -1) {
            console.error('Could not find ASIN column');
            return false;
        }

        // Map column headers to their indices
        const colIndices = {};
        for (let col = 1; col <= 20; col++) {
            const header = worksheet.cell(headerRow, col).value();
            if (header) {
                colIndices[header.toString().trim()] = col;
            }
        }

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
        Object.entries(updates).forEach(([field, value]) => {
            if (field !== 'asin' && colIndices[field] !== undefined) {
                worksheet.cell(targetRow, colIndices[field]).value(value);
                console.log(`Updated ${field} to ${value}`);
            }
        });

        // Save the workbook
        await workbook.toFileAsync(filePath);
        console.log(`\nSaved updates to file: ${filePath}`);
        return true;

    } catch (error) {
        console.error('Error processing Excel file:', error.message);
        return false;
    }
}

async function main() {
    const fileId = '11xfaf0nGgscOpYpE9U3tfRelyCxsDoOJ';
    const excelFile = 'ICERWORKSHEET.xlsx';

    // Updates for specific ASIN
    const updates = {
        asin: 'B084TRRKBY',
        'FBA INV': 98,
        'OTS QOH': 123,
        'QOHQTY': 456,
        'AMZ VC INV': 789,
        'WIP QTY': 10,
        'WIP ETA': '2024-07-01'
    };

    // First download, then update
    const downloaded = await downloadFile(fileId, excelFile);
    if (downloaded) {
        await updateExcelFile(excelFile, updates);
    }
}

main().catch(console.error); 