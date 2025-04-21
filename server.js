const express = require('express');
const axios = require('axios');
const qs = require('qs');
const dotenv = require('dotenv');
dotenv.config();

const app = express();
const port = process.env.PORT || 3000;

// Root route — this is what makes "/" not show "Cannot GET /"
app.get('/', (req, res) => {
  res.send('Excel API Backend is running ✔️');
});

// /calculate endpoint
app.get('/calculate', async (req, res) => {
  const quantity = req.query.quantity;

  try {
    // Step 1: Get access token
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
      qs.stringify({
        grant_type: 'client_credentials',
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
      }),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      }
    );

    const accessToken = tokenResponse.data.access_token;

    // Step 2: Update the quantity in Excel
    await axios.patch(
      `https://graph.microsoft.com/v1.0/users/${process.env.USER_EMAIL}/drive/items/${process.env.EXCEL_FILE_ID}/workbook/worksheets('Sheet1')/range(address='A1')`,
      { values: [[Number(quantity)]] },
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
      }
    );

    // Step 3: Read the calculated result from cell B1
    const readResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${process.env.USER_EMAIL}/drive/items/${process.env.EXCEL_FILE_ID}/workbook/worksheets('Sheet1')/range(address='B1')`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

    const calculatedValue = readResponse.data.values[0][0];
    res.json({ result: calculatedValue });

  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
