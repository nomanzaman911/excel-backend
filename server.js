require('dotenv').config();
const express = require('express');
const axios = require('axios');
const bodyParser = require('body-parser');

const app = express();
app.use(bodyParser.json());

const {
  CLIENT_ID,
  TENANT_ID,
  CLIENT_SECRET,
  USER_EMAIL,
  EXCEL_FILE_NAME
} = process.env;

// Get access token
async function getAccessToken() {
  const response = await axios.post(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    new URLSearchParams({
      client_id: CLIENT_ID,
      scope: 'https://graph.microsoft.com/.default',
      client_secret: CLIENT_SECRET,
      grant_type: 'client_credentials',
    }),
    { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
  );
  return response.data.access_token;
}

// Get file ID by file name
async function getExcelFileId(token) {
  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/root/children`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  const file = response.data.value.find(f => f.name === EXCEL_FILE_NAME);
  if (!file) throw new Error("Excel file not found in OneDrive.");
  return file.id;
}

// POST or GET /calculate
app.all('/calculate', async (req, res) => {
  try {
    const quantity = req.method === 'POST' ? req.body.quantity : req.query.quantity;
    if (!quantity) return res.status(400).json({ error: 'Missing quantity' });

    const token = await getAccessToken();
    const fileId = await getExcelFileId(token);

    // Write quantity to A2
    await axios.patch(
      `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/items/${fileId}/workbook/worksheets('Sheet1')/range(address='A2')`,
      { values: [[quantity]] },
      { headers: { Authorization: `Bearer ${token}` } }
    );

    // Read result from B2
    const resultRes = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/items/${fileId}/workbook/worksheets('Sheet1')/range(address='B2')`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const result = resultRes.data.values?.[0]?.[0] || null;
    res.json({ result });

  } catch (err) {
    console.error('Error:', err.response?.data || err.message);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Server running on port ${PORT}`));
