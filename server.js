const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');

const app = express();
app.use(cors());
app.use(express.json());

app.post('/calculate', async (req, res) => {
  const { quantity } = req.body;

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('excel.xlsx');

    const sheet = workbook.getWorksheet(1);

    // ✅ Write input value to A2
    sheet.getCell('A2').value = quantity;

    // ✅ Recalculate workbook formulas (only if formula is in B2)
    // Note: ExcelJS can't calculate Excel formulas itself (e.g. "=A2*10")
    // So if your B2 cell has a formula, you need to pre-calculate or use Excel locally
    // BUT you can use a helper JS formula here:
    const result = quantity * 10; // Replace this with your formula logic if needed

    // Optionally, update cell B2 too:
    sheet.getCell('B2').value = result;

    await workbook.xlsx.writeFile('excel.xlsx'); // Save changes

    res.json({ result });
  } catch (err) {
    console.error(err);
    res.status(500).send('Something went wrong!');
  }
});

app.listen(3000, () => {
  console.log('Server is running on port 3000');
});
