const express = require('express');
const fetch = require('node-fetch');
const bodyParser = require('body-parser');

const app = express();
const port = 10000;

app.use(bodyParser.json());

// HARDCODED CREDENTIALS
const CLIENT_ID = '53f1b63e-e169-4121-a255-c0a966ca514e';
const TENANT_ID = '6940843a-674d-4941-9ca2-dc5603f278df';
const CLIENT_SECRET = 'qmB8Q~phRIvnOQl5R5WHcLxu3~by0Z2pkqGx9cAq';
const EXCEL_FILE_ID = '24C52B39E61CD77F!sc28933942504417abf44a2ea279e2610';
const USER_EMAIL = 'nomanzaman@outlook.com';

const getAccessToken = async () => {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('client_id', CLIENT_ID);
  params.append('scope', 'https://graph.microsoft.com/.default');
  params.append('client_secret', CLIENT_SECRET);
  params.append('grant_type', 'client_credentials');

  const res = await fetch(url, {
    method: 'POST',
    body: params
  });

  const data = await res.json();
  return data.access_token;
};

app.get('/', (req, res) => {
  res.send('✅ Excel API Backend is running');
});

app.get('/calculate', async (req, res) => {
  const quantity = parseInt(req.query.quantity || 1);

  try {
    const accessToken = await getAccessToken();

    const updateCellUrl = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/items/${EXCEL_FILE_ID}/workbook/worksheets('Sheet1')/range(address='A1')`;
    const updateRes = await fetch(updateCellUrl, {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ values: [[quantity]] })
    });

    const readCellUrl = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/items/${EXCEL_FILE_ID}/workbook/worksheets('Sheet1')/range(address='B1')`;
    const resultRes = await fetch(readCellUrl, {
      method: 'GET',
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    const resultData = await resultRes.json();
    const calculatedValue = resultData.values?.[0]?.[0] ?? null;

    if (calculatedValue !== null) {
      res.json({ result: calculatedValue });
    } else {
      res.status(500).json({ error: 'No result returned from Excel' });
    }

  } catch (err) {
    console.error('Error:', err);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
});
