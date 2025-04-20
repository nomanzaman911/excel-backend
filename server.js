const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());

app.post('/calculate', async (req, res) => {
  const quantity = req.body.quantity;

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, 'excel.xlsx'));
    const worksheet = workbook.getWorksheet(1);

    worksheet.getCell('A2').value = quantity;

    // Get the formula result from B2
    const totalCell = worksheet.getCell('B2');

    await workbook.xlsx.writeFile(path.join(__dirname, 'excel.xlsx'));

    res.json({
      input: quantity,
      total: totalCell.result || totalCell.value
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Error processing Excel file' });
  }
});

app.listen(PORT, () => {
  console.log(`✅ Server running at http://localhost:${PORT}`);
});
