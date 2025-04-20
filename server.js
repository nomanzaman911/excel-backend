const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

const sourcePath = path.join(__dirname, 'excel.xlsx');
const tempPath = path.join('/tmp', 'excel.xlsx');

app.post('/calculate', async (req, res) => {
  const { quantity } = req.body;

  try {
    // Copy excel file to /tmp if not there already
    if (!fs.existsSync(tempPath)) {
      fs.copyFileSync(sourcePath, tempPath);
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(tempPath);
    const sheet = workbook.getWorksheet(1);

    // Write quantity to A2
    sheet.getCell('A2').value = quantity;

    // Simulate calculation (since Excel formulas don't auto-run)
    const result = quantity * 10;

    // Optional: write result to B2 in Excel file
    sheet.getCell('B2').value = result;
    await workbook.xlsx.writeFile(tempPath);

    res.json({ result }); // Send result back to frontend
  } catch (err) {
    console.error('Server Error:', err.message);
    res.status(500).json({ error: 'Server error' });
  }
});

app.listen(3000, () => {
  console.log('Server running on port 3000');
});
