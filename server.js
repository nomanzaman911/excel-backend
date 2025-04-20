const express = require('express');
const axios = require('axios');
const cors = require('cors');
require('dotenv').config();

const app = express();
app.use(cors());

const PORT = process.env.PORT || 3000;

// Get access token using Microsoft identity platform (client credentials flow)
async function getAccessToken() {
  const url = 'https://login.microsoftonline.com/' + process.env.TENANT_ID + '/oauth2/v2.0/token';
  const params = new URLSearchParams();
  params.append('client_id', process.env.CLIENT_ID);
  params.append('scope', 'https://graph.microsoft.com/.default');
  params.append('client_secret', process.env.CLIENT_SECRET);
  params.append('grant_type', 'client_credentials');

  const response = await axios.post(url, params);
  return response.data.access_token;
}

app.get('/calculate', async (req, res) => {
  const quantity = req.query.quantity || 1;

  try {
    const accessToken = await getAccessToken();

    // Update cell A2 in the Excel file
    await axios.patch(
      `https://graph.microsoft.com/v1.0/users/nomanzaman@outlook.com/drive/root:/calculator.xlsx:/workbook/worksheets('Sheet1')/range(address='A2')`,
      { values: [[parseInt(quantity)]] },
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    // Read calculated value from cell B2
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/users/nomanzaman@outlook.com/drive/root:/calculator.xlsx:/workbook/worksheets('Sheet1')/range(address='B2')`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const calculatedValue = response.data.values[0][0];
    res.json({ result: calculatedValue });
  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
