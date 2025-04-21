const express = require('express');
const open = require('open');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const app = express();
const port = 3000;

// === Azure App Config ===
const config = {
  auth: {
    clientId: '1140a629-6ea1-41ec-9655-d5e1afab2408',
    authority: 'https://login.microsoftonline.com/6940843a-674d-4941-9ca2-dc5603f278df',
    clientSecret: 'wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ',
  }
};

const REDIRECT_URI = 'http://localhost:3000/auth/callback';
const SCOPES = ["user.read", "files.readwrite", "offline_access"];

const pca = new msal.ConfidentialClientApplication(config);

let accessToken = null;

// === Step 1: Start Login Flow ===
app.get('/auth', (req, res) => {
  const authCodeUrlParams = {
    scopes: SCOPES,
    redirectUri: REDIRECT_URI
  };

  pca.getAuthCodeUrl(authCodeUrlParams).then((response) => {
    res.redirect(response);
  }).catch(err => res.status(500).send(err));
});

// === Step 2: Handle Redirect with Auth Code ===
app.get('/auth/callback', (req, res) => {
  const tokenRequest = {
    code: req.query.code,
    scopes: SCOPES,
    redirectUri: REDIRECT_URI,
  };

  pca.acquireTokenByCode(tokenRequest).then((response) => {
    accessToken = response.accessToken;
    res.send("✅ Authentication successful. You can now go to /calculate?quantity=5");
  }).catch(err => res.status(500).send(err));
});

// === Step 3: Use Graph API to Access Excel File ===
app.get('/calculate', async (req, res) => {
  if (!accessToken) return res.send({ error: "User not authenticated. Go to /auth first." });

  const quantity = parseInt(req.query.quantity || '1');

  try {
    const client = Client.init({
      authProvider: done => done(null, accessToken)
    });

    const filePath = '/calculator.xlsx';
    const sheetName = 'Sheet1';
    const inputCell = 'A1';
    const resultCell = 'B1';

    // Update quantity
    await client
      .api(`/me/drive/root:${filePath}:/workbook/worksheets('${sheetName}')/range(address='${inputCell}')`)
      .patch({ values: [[quantity]] });

    // Read result
    const result = await client
      .api(`/me/drive/root:${filePath}:/workbook/worksheets('${sheetName}')/range(address='${resultCell}')`)
      .get();

    res.send({ result: result.values?.[0]?.[0] || "No result returned from Excel" });

  } catch (err) {
    console.error("Error:", JSON.stringify(err, null, 2));
    res.send({ error: "Failed to calculate result" });
  }
});

// === Default Route ===
app.get('/', (req, res) => {
  res.send("✅ Excel API Backend is running.<br><br>➡️ <a href='/auth'>Click here to log in</a>");
});

app.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
});
