require('dotenv').config();
const express = require('express');
const axios = require('axios');
const qs = require('querystring');

const app = express();
const PORT = process.env.PORT || 3000;

// Get access token from Microsoft Identity platform
async function getAccessToken() {
  const url = 'https://login.microsoftonline.com/' + process.env.TENANT_ID + '/oauth2/v2.0/token';

  const response = await axios.post(url, qs.stringify({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  }));

  return response.data.access_token;
}

// Call Excel API to set value and get calculated result
async function calculateResult(quantity) {
  const accessToken = await getAccessToken();
  const filePath = '/users/nomanzaman@outlook.com/drive/root:/calculator.xlsx';

  const setUrl = `https://graph.microsoft.com/v1.0${filePath}:/workbook/worksheets('Sheet1')/range(address='A2')`;
  const getUrl = `https://graph.microsoft.com/v1.0${filePath}:/workbook/worksheets('Sheet1')/range(address='B2')`;

  // Set quantity to A2
  await axios.patch(setUrl, {
    values: [[quantity]]
  }, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    }
  });

  // Get calculated value from B2
  const response = await axios.get(getUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });

  return response.data.values[0][0];
}

app.get('/calculate', async (req, res) => {
  const quantity = parseInt(req.query.quantity);

  if (isNaN(quantity)) {
    return res.status(400).json({ error: 'Invalid quantity' });
  }

  try {
    const result = await calculateResult(quantity);
    res.json({ result });
  } catch (err) {
    console.error('Error:', err.response?.data || err.message);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
