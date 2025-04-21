require('dotenv').config();
const express = require('express');
const axios = require('axios');
const qs = require('qs');

const app = express();
const PORT = process.env.PORT || 3000;

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const EXCEL_FILE_ID = process.env.EXCEL_FILE_ID;
const USER_EMAIL = process.env.USER_EMAIL;

async function getAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const tokenData = {
    grant_type: 'client_credentials',
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default'
  };

  const response = await axios.post(tokenUrl, qs.stringify(tokenData), {
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    }
  });

  return response.data.access_token;
}

async function updateExcelAndFetchResult(quantity) {
  const accessToken = await getAccessToken();
  const workbookUrl = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/root:/calculator.xlsx:/workbook/worksheets('Sheet1')/range(address='A1')`;

  // Update cell A1 with quantity
  await axios.patch(workbookUrl, {
    values: [[quantity]]
  }, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    }
  });

  // Read result from cell B1
  const resultUrl = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/root:/calculator.xlsx:/workbook/worksheets('Sheet1')/range(address='B1')`;
  const response = await axios.get(resultUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });

  return response.data.values[0][0];
}

app.get('/calculate', async (req, res) => {
  try {
    const quantity = parseInt(req.query.quantity);
    const calculatedValue = await updateExcelAndFetchResult(quantity);
    res.json({ result: calculatedValue });
  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
