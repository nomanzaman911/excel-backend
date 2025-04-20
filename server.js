const express = require('express');
const axios = require('axios');
const qs = require('qs');
const app = express();
const PORT = process.env.PORT || 3000;

// Your Microsoft App credentials
const TENANT_ID = '6940843a-674d-4941-9ca2-dc5603f278df';
const CLIENT_ID = '53f1b63e-e169-4121-a255-c0a966ca514e';
const CLIENT_SECRET = 'qmB8Q~phRIvnOQl5R5WHcLxu3~by0Z2pkqGx9cAq';
const EXCEL_FILE_ID = 'IQSUM4nCBCV6Qb9EouonniYQAWYCUuKDO82wYx0B2DRX01Q';
const SHEET_NAME = 'Sheet1'; // adjust if your sheet name is different

const getToken = async () => {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const formData = {
    grant_type: 'client_credentials',
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default'
  };

  const response = await axios.post(tokenUrl, qs.stringify(formData), {
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    }
  });

  return response.data.access_token;
};

const updateCell = async (token, value) => {
  const url = `https://graph.microsoft.com/v1.0/me/drive/items/${EXCEL_FILE_ID}/workbook/worksheets('${SHEET_NAME}')/range(address='A2')`;

  await axios.patch(
    url,
    { values: [[value]] },
    {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    }
  );
};

const readCell = async (token) => {
  const url = `https://graph.microsoft.com/v1.0/me/drive/items/${EXCEL_FILE_ID}/workbook/worksheets('${SHEET_NAME}')/range(address='B2')`;

  const response = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  return response.data.values[0][0];
};

app.get('/calculate', async (req, res) => {
  const quantity = req.query.quantity;
  if (!quantity) return res.status(400).send('Quantity is required');

  try {
    const token = await getToken();
    await updateCell(token, quantity);
    const result = await readCell(token);
    res.json({ result });
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).send('Something went wrong');
  }
});

app.listen(PORT, () => {
  console.log(`✅ Server running on port ${PORT}`);
});
