const express = require('express');
console.log('Express loaded successfully');
const app = express();
console.log('Express app created');
const XLSX = require('xlsx');
const path = require('path');

app.get("/", (req, res) => {
  res.send("Hello, Express!");
});


app.get('/debug-fba-asin', (req, res) => {
  const filePath = path.join(__dirname, 'ICERWORKSHEET.xlsx');
  const sheetName = 'AMS NFL';
  const workbook = XLSX.readFile(filePath);
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    return res.status(404).send('Sheet AMS NFL not found');
  }
  const data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
  if (data.length === 0) {
    return res.status(404).send('No data found in AMS NFL sheet');
  }
  // Get column names
  const columns = Object.keys(data[0]);
  // Get first 5 rows
  const sampleRows = data.slice(0, 5);
  res.json({ columns, sampleRows });
});

app.get('/update-fba-b084trrkby', (req, res) => {
  const filePath = path.join(__dirname, 'ICERWORKSHEET.xlsx');
  const sheetName = 'AMS NFL';
  const asinToUpdate = 'B084TRRKBY';
  const newFbaInvValue = 98;

  const workbook = XLSX.readFile(filePath);
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    return res.status(404).send('Sheet AMS NFL not found');
  }

  // Find the header row and the columns for ASIN and FBA INV
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  let asinColIdx = -1, fbaInvColIdx = -1, headerRowIdx = -1;

  for (let R = range.s.r; R <= range.e.r; ++R) {
    let foundAsin = false, foundFbaInv = false;
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = { c: C, r: R };
      const cellRef = XLSX.utils.encode_cell(cellAddress);
      const cell = worksheet[cellRef];
      if (cell && String(cell.v).trim().toUpperCase() === 'ASIN') {
        asinColIdx = C;
        foundAsin = true;
      }
      if (cell && String(cell.v).trim().toUpperCase() === 'FBA INV') {
        fbaInvColIdx = C;
        foundFbaInv = true;
      }
    }
    if (foundAsin && foundFbaInv) {
      headerRowIdx = R;
      break;
    }
  }

  if (asinColIdx === -1 || fbaInvColIdx === -1) {
    return res.status(404).send('ASIN or FBA INV column not found');
  }

  // Now, find the row with the matching ASIN
  let updated = false;
  for (let R = headerRowIdx + 1; R <= range.e.r; ++R) {
    const asinCellRef = XLSX.utils.encode_cell({ c: asinColIdx, r: R });
    const asinCell = worksheet[asinCellRef];
    if (asinCell && String(asinCell.v).trim().toUpperCase() === asinToUpdate) {
      const fbaInvCellRef = XLSX.utils.encode_cell({ c: fbaInvColIdx, r: R });
      worksheet[fbaInvCellRef] = { t: 'n', v: newFbaInvValue };
      updated = true;
      break;
    }
  }

  if (!updated) {
    return res.status(404).send('ASIN not found');
  }

  XLSX.writeFile(workbook, filePath);
  res.send(`FBA INV value updated to ${newFbaInvValue} for ASIN: ${asinToUpdate} (formatting preserved)`);
});

app.get('/debug-headers', (req, res) => {
  const filePath = path.join(__dirname, 'ICERWORKSHEET.xlsx');
  const sheetName = 'AMS NFL';
  const workbook = XLSX.readFile(filePath);
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    return res.status(404).send('Sheet AMS NFL not found');
  }
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  let output = [];
  for (let R = range.s.r; R <= Math.min(range.e.r, range.s.r + 4); ++R) {
    let row = [];
    for (let C = range.s.c; C <= Math.min(range.e.c, range.s.c + 20); ++C) {
      const cellAddress = { c: C, r: R };
      const cellRef = XLSX.utils.encode_cell(cellAddress);
      const cell = worksheet[cellRef];
      row.push(cell ? cell.v : '');
    }
    output.push(row);
  }
  res.json({ rows: output });
});

app.get('/update-asin-fields', (req, res) => {
  const filePath = path.join(__dirname, 'ICERWORKSHEET.xlsx');
  const sheetName = 'AMS NFL';
  const asinToUpdate = req.query.asin;
  if (!asinToUpdate) {
    return res.status(400).send('Please provide an asin query parameter.');
  }

  // Remove asin from the query to get the fields to update
  const updates = { ...req.query };
  delete updates.asin;

  if (Object.keys(updates).length === 0) {
    return res.status(400).send('Please provide at least one field to update.');
  }

  const workbook = XLSX.readFile(filePath);
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    return res.status(404).send('Sheet AMS NFL not found');
  }

  // Find the header row and the columns for ASIN and the fields to update
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  let headerRowIdx = -1;
  const colMap = {}; // { fieldName: colIdx }

  for (let R = range.s.r; R <= range.e.r; ++R) {
    let foundAsin = false;
    let foundAll = true;
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = { c: C, r: R };
      const cellRef = XLSX.utils.encode_cell(cellAddress);
      const cell = worksheet[cellRef];
      if (cell) {
        const val = String(cell.v).trim();
        if (val.toUpperCase() === 'ASIN') {
          colMap['ASIN'] = C;
          foundAsin = true;
        }
        // Map all requested fields
        for (const field of Object.keys(updates)) {
          if (val.toUpperCase() === field.toUpperCase()) {
            colMap[field] = C;
          }
        }
      }
    }
    // Check if all fields are found
    foundAll = Object.keys(updates).every(f => colMap[f] !== undefined);
    if (foundAsin && foundAll) {
      headerRowIdx = R;
      break;
    }
  }

  if (colMap['ASIN'] === undefined) {
    return res.status(404).send('ASIN column not found');
  }
  for (const field of Object.keys(updates)) {
    if (colMap[field] === undefined) {
      return res.status(404).send(`Column not found: ${field}`);
    }
  }

  // Now, find the row with the matching ASIN
  let updated = false;
  for (let R = headerRowIdx + 1; R <= range.e.r; ++R) {
    const asinCellRef = XLSX.utils.encode_cell({ c: colMap['ASIN'], r: R });
    const asinCell = worksheet[asinCellRef];
    if (asinCell && String(asinCell.v).trim().toUpperCase() === asinToUpdate.toUpperCase()) {
      // Update all requested fields
      for (const [field, value] of Object.entries(updates)) {
        const colIdx = colMap[field];
        const cellRef = XLSX.utils.encode_cell({ c: colIdx, r: R });
        // Try to parse as number, otherwise keep as string
        const numVal = Number(value);
        worksheet[cellRef] = isNaN(numVal)
          ? { t: 's', v: value }
          : { t: 'n', v: numVal };
      }
      updated = true;
      break;
    }
  }

  if (!updated) {
    return res.status(404).send('ASIN not found');
  }

  XLSX.writeFile(workbook, filePath);
  res.send(`Fields updated for ASIN: ${asinToUpdate}`);
});

// Start the server
const port = 3000;
console.log('Attempting to start server...');
app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});
console.log('After listen() call');
