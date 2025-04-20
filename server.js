const express = require('express');
const axios = require('axios');
const qs = require('qs');
const cors = require('cors');
require('dotenv').config();

const app = express();
app.use(cors()); // ✅ Allow frontend access from other domains

const PORT = process.env.PORT || 3000;

// Environment variables (set in Render)
const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const excelFileId = process.env.EXCEL_FILE_ID;

const getAccessToken = async () => {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const data = qs.stringify({
    grant_type: 'client_credentials',
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default'
  });

  const headers = {
    'Content-Type': 'application/x-www-form-urlencoded'
  };

  const response = await axios.post(url, data, { headers });
  return response.data.access_token;
};

app.get('/calculate', async (req, res) => {
  const quantity = req.query.quantity || '1';

  try {
    const accessToken = await getAccessToken();

    const baseURL = `https://graph.microsoft.com/v1.0/me/drive/items/${excelFileId}/workbook/worksheets('Sheet1')`;

    // Update cell A2 with quantity
    await axios.patch(
      `${baseURL}/range(address='A2')`,
      { values: [[quantity]] },
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    // Read result from B2
    const response = await axios.get(
      `${baseURL}/range(address='B2')`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const result = response.data.values[0][0];
    res.json({ result });

  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Calculation failed' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
