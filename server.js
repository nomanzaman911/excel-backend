const express = require('express');
const axios = require('axios');
const dotenv = require('dotenv');
dotenv.config();

const app = express();
const port = process.env.PORT || 3000;

const getAccessToken = async () => {
  const response = await axios.post(
    `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
    new URLSearchParams({
      client_id: process.env.CLIENT_ID,
      scope: 'https://graph.microsoft.com/.default',
      client_secret: process.env.CLIENT_SECRET,
      grant_type: 'client_credentials'
    }),
    {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    }
  );
  return response.data.access_token;
};

const updateAndReadExcel = async (quantity) => {
  const accessToken = await getAccessToken();
  const fileId = process.env.EXCEL_FILE_ID;
  const sheet = process.env.EXCEL_WORKSHEET_NAME || 'Sheet1';
  const inputCell = process.env.EXCEL_CELL_INPUT || 'A2';
  const outputCell = process.env.EXCEL_CELL_OUTPUT || 'B2';

  // 🔄 Update quantity in Excel
  await axios.patch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${sheet}')/range(address='${inputCell}')`,
    {
      values: [[parseInt(quantity)]]
    },
    {
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    }
  );

  // 🔁 Read calculated result from Excel
  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${sheet}')/range(address='${outputCell}')`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    }
  );

  return response.data.values[0][0];
};

app.get('/calculate', async (req, res) => {
  const quantity = req.query.quantity;
  if (!quantity) return res.status(400).json({ error: 'Missing quantity' });

  try {
    const result = await updateAndReadExcel(quantity);
    res.json({ result });
  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
