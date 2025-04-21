const express = require('express');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ConfidentialClientApplication } = require('@azure/msal-node');
require('isomorphic-fetch');

const app = express();
const port = process.env.PORT || 10000;

// Hardcoded values
const CLIENT_ID = '53f1b63e-e169-4121-a255-c0a966ca514e';
const TENANT_ID = '6940843a-674d-4941-9ca2-dc5603f278df';
const CLIENT_SECRET = 'qmB8Q~phRIvnOQl5R5WHcLxu3~by0Z2pkqGx9cAq';
const EXCEL_FILE_ID = '24C52B39E61CD77F!sc28933942504417abf44a2ea279e2610';
const USER_EMAIL = 'nomanzaman@outlook.com';

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: CLIENT_SECRET
  }
};

const cca = new ConfidentialClientApplication(msalConfig);

app.get('/', (req, res) => {
  res.send('Excel API Backend is running');
});

app.get('/calculate', async (req, res) => {
  const quantity = parseFloat(req.query.quantity);
  if (isNaN(quantity)) {
    return res.status(400).json({ error: 'Invalid quantity' });
  }

  try {
    const tokenResponse = await cca.acquireTokenByClientCredential({
      scopes: ['https://graph.microsoft.com/.default']
    });

    const client = Client.init({
      authProvider: done => done(null, tokenResponse.accessToken)
    });

    const address = 'Sheet1!A1';
    await client
      .api(`/users/${USER_EMAIL}/drive/items/${EXCEL_FILE_ID}/workbook/worksheets('Sheet1')/range(address='${address}')`)
      .patch({
        values: [[quantity]]
      });

    const resultRange = await client
      .api(`/users/${USER_EMAIL}/drive/items/${EXCEL_FILE_ID}/workbook/worksheets('Sheet1')/range(address='Sheet1!B1')`)
      .get();

    const calculatedValue = resultRange.values[0][0];
    res.json({ result: calculatedValue });

  } catch (error) {
    console.error('Error:', JSON.stringify(error, null, 2));
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
