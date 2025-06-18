const XLSX = require('xlsx');
const XLSXPopulate = require('xlsx-populate');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');
const readline = require('readline');

// Load credentials from file
const credentials = JSON.parse(fs.readFileSync('client_secret.json'));

const SCOPES = ['https://www.googleapis.com/auth/drive.file'];
const TOKEN_PATH = 'token.json';

// Create an OAuth2 client
function getOAuth2Client() {
    return new google.auth.OAuth2(
        credentials.installed.client_id,
        credentials.installed.client_secret,
        'urn:ietf:wg:oauth:2.0:oob'
    );
}

// Get and store new token after prompting for user authorization
async function getNewToken(oAuth2Client) {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
    });
    console.log('Authorize this app by visiting this url:', authUrl);
    console.log('\nAfter authorizing, you will see a code on the page. Copy that code and paste it here.');

    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });

    return new Promise((resolve, reject) => {
        rl.question('Enter the code from that page here: ', async (code) => {
            rl.close();
            try {
                const { tokens } = await oAuth2Client.getToken(code);
                oAuth2Client.setCredentials(tokens);
                // Store the token to disk for later program executions
                fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
                console.log('Token stored to', TOKEN_PATH);
                resolve(oAuth2Client);
            } catch (err) {
                reject('Error retrieving access token: ' + err);
            }
        });
    });
}

// Create Google Drive client with OAuth2 authentication
async function createDriveClient() {
    const oAuth2Client = getOAuth2Client();

    try {
        // Check if we have previously stored a token
        if (fs.existsSync(TOKEN_PATH)) {
            const token = JSON.parse(fs.readFileSync(TOKEN_PATH));
            oAuth2Client.setCredentials(token);
        } else {
            // If no token, get a new one
            await getNewToken(oAuth2Client);
        }
        return google.drive({ version: 'v3', auth: oAuth2Client });
    } catch (error) {
        console.error('Error creating drive client:', error);
        throw error;
    }
}

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

        // Find ASIN column and row
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

        // Process each ASIN's updates
        for (const asinUpdate of updates) {
            const { asin, ...fieldUpdates } = asinUpdate;

            // Find the row with matching ASIN
            let targetRow = -1;
            const range = worksheet.usedRange();
            const maxRow = range.endCell().rowNumber();
            
            for (let row = headerRow + 1; row <= maxRow; row++) {
                const cellValue = worksheet.cell(row, asinCol).value();
                if (cellValue && cellValue.toString().trim().toUpperCase() === asin.toUpperCase()) {
                    targetRow = row;
                    break;
                }
            }

            if (targetRow === -1) {
                console.error(`\nASIN ${asin} not found in the sheet`);
                continue; // Skip to next ASIN
            }

            console.log(`\nFound ASIN ${asin} at row ${targetRow}`);

            // Update only the values that exist in the sheet
            let updatedFields = [];
            let skippedFields = [];
            
            Object.entries(fieldUpdates).forEach(([field, value]) => {
                // Try to find the column case-insensitively
                const upperField = field.toUpperCase();
                if (colIndices[upperField] !== undefined) {
                    worksheet.cell(targetRow, colIndices[upperField]).value(value);
                    updatedFields.push(`${headerValues[upperField]}: ${value}`);
                } else {
                    skippedFields.push(field);
                }
            });

            if (updatedFields.length > 0) {
                console.log('Updated fields:', updatedFields.join(', '));
            }
            if (skippedFields.length > 0) {
                console.log('Skipped fields (not found in sheet):', skippedFields.join(', '));
            }
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

// Helper: Find file by name in Google Drive
async function findFileIdByName(fileName) {
    const drive = await createDriveClient();
    const res = await drive.files.list({
        q: `name='${fileName.replace(/'/g, "\\'")}' and trashed=false`,
        fields: 'files(id, name)',
        spaces: 'drive',
        pageSize: 1
    });
    if (res.data.files && res.data.files.length > 0) {
        return res.data.files[0].id;
    }
    return null;
}

// Function to upload file to Google Drive
async function uploadToDrive(filePath, fileId = null) {
    try {
        const drive = await createDriveClient();
        const fileMetadata = {
            name: path.basename(filePath)
        };
        const media = {
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            body: fs.createReadStream(filePath)
        };

        let response;
        if (fileId) {
            // Update existing file
            response = await drive.files.update({
                fileId: fileId,
                resource: fileMetadata,
                media: media,
                fields: 'id'
            });
            console.log('File updated successfully:', response.data.id);
        } else {
            // Create new file
            response = await drive.files.create({
                resource: fileMetadata,
                media: media,
                fields: 'id'
            });
            console.log('New file created successfully:', response.data.id);
        }
        return response.data.id;
    } catch (error) {
        console.error('Error uploading file to Drive:', error.message);
        throw error;
    }
}

async function main() {
    const excelFile = 'ICERWORKSHEET.xlsx';

    const updates = {
        'AMS NBA': [
            {
                asin: 'B01LZZHGGM',
                'AMZ VC INV': 400,
                'OTSQOH': 90
            }
        ]
    };

    try {
        // Update each sheet
        for (const [sheetName, sheetUpdates] of Object.entries(updates)) {
            console.log(`\n=== Processing sheet: ${sheetName} ===`);
            await updateExcelFile(excelFile, sheetName, sheetUpdates);
        }

        try {
            // Find file by name in Drive
            const fileId = await findFileIdByName(excelFile);
            if (fileId) {
                console.log('File already exists in Drive. Updating...');
                await uploadToDrive(excelFile, fileId);
            } else {
                console.log('File not found in Drive. Creating new file...');
                await uploadToDrive(excelFile);
            }
        } catch (error) {
            console.error('Failed to upload file to Drive:', error.message);
        }
    } catch (error) {
        console.error('Error in main:', error.message);
    }
}

// If running directly (not as a module)
if (require.main === module) {
    main().catch(console.error);
}

// Export functions for use as a module
module.exports = {
    downloadFile,
    updateExcelFile,
    uploadToDrive
}; 