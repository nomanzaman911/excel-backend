const express = require('express');
const session = require('express-session');
const fetch = require('node-fetch');
const qs = require('querystring');
const app = express();

const clientId = "1140a629-6ea1-41ec-9655-d5e1afab2408";
const clientSecret = "wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ";
const tenantId = "common"; // Use "common" for personal accounts
const redirectUri = "https://excel-backend-1-y1fk.onrender.com/auth/callback";

app.use(express.json());
app.use(session({ secret: 'secret', resave: false, saveUninitialized: true }));

app.get('/auth', (req, res) => {
  const params = new URLSearchParams({
    client_id: clientId,
    response_type: 'code',
    redirect_uri: redirectUri,
    response_mode: 'query',
    scope: 'https://graph.microsoft.com/Files.ReadWrite https://graph.microsoft.com/User.Read offline_access',
  });

  res.redirect(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${params}`);
});

app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;

  const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    body: qs.stringify({
      client_id: clientId,
      scope: 'https://graph.microsoft.com/Files.ReadWrite https://graph.microsoft.com/User.Read offline_access',
      code,
      redirect_uri: redirectUri,
      grant_type: 'authorization_code',
      client_secret: clientSecret
    }),
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });

  const tokenData = await tokenRes.json();
  req.session.accessToken = tokenData.access_token;
  res.redirect('https://theglowup.com.au/index.html'); // Replace with actual URL
});

app.post('/calculate', async (req, res) => {
  const accessToken = req.session.accessToken;
  if (!accessToken) return res.status(401).json({ error: "Not signed in" });

  const quantity = req.body.quantity;
  const filePath = '/calculator.xlsx';
  const workbookRange = 'Sheet1!A1';

  // Update the input cell
  await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:${filePath}:/workbook/worksheets('Sheet1')/range(address='${workbookRange}')`, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ values: [[quantity]] })
  });

  // Read the result cell
  const resultCell = 'Sheet1!B1'; // Adjust based on your sheet
  const resultRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:${filePath}:/workbook/worksheets('Sheet1')/range(address='${resultCell}')`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  const resultData = await resultRes.json();
  const result = resultData.values?.[0]?.[0];
  res.json({ result });
});

app.listen(10000, () => {
  console.log('✅ Server listening on port 10000');
});
