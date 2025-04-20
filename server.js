const express = require('express');
const axios = require('axios');
const dotenv = require('dotenv');
const app = express();
dotenv.config();

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const PORT = process.env.PORT || 3000;

const EXCEL_FILE_PATH = '/calculator.xlsx'; // File in your root OneDrive
const WORKSHEET_NAME = 'Sheet1'; // Update if your sheet is named differently

// Get access token
async function getAccessToken() {
  const response = await axios.post('https://login.microsoftonline.com/' + process.env.TENANT_ID + '/oauth2/v2.0/token', new URLSearchParams({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  }));
  return response.data.access_token;
}

// Update Excel input cell and get result
app.get('/calculate', async (req, res) => {
  try {
    const quantity = req.query.quantity;
    const token = await getAccessToken();

    // Update cell A2 with quantity
    await axios.patch(
      `https://graph.microsoft.com/v1.0/me/drive/root:${EXCEL_FILE_PATH}:/workbook/worksheets('${WORKSHEET_NAME}')/range(address='A2')`,
      { values: [[parseInt(quantity)]] },
      { headers: { Authorization: `Bearer ${token}` } }
    );

    // Read calculated result from B2
    const resultResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/root:${EXCEL_FILE_PATH}:/workbook/worksheets('${WORKSHEET_NAME}')/range(address='B2')`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const calculatedValue = resultResponse.data.values[0][0];
    res.json({ result: calculatedValue });
  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
