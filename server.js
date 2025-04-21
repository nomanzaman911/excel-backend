const express = require('express');
const session = require('express-session');
const axios = require('axios');
const cors = require('cors');
const querystring = require('querystring');

const app = express();
const port = 10000;

const clientId = '1140a629-6ea1-41ec-9655-d5e1afab2408';
const clientSecret = 'wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ';
const redirectUri = 'https://excel-backend-1-y1fk.onrender.com/auth/callback';
const tenantId = 'common';

app.use(cors({
  origin: 'http://theglowup.com.au',
  credentials: true
}));

app.use(express.json());
app.use(session({
  secret: 'secret',
  resave: false,
  saveUninitialized: true
}));

app.get('/', (req, res) => {
  res.send('✅ Excel API Backend is running');
});

app.get('/auth', (req, res) => {
  const params = querystring.stringify({
    client_id: clientId,
    response_type: 'code',
    redirect_uri: redirectUri,
    response_mode: 'query',
    scope: 'offline_access Files.ReadWrite.All',
  });
  res.redirect(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${params}`);
});

app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;
  try {
    const tokenRes = await axios.post(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      querystring.stringify({
        client_id: clientId,
        scope: 'offline_access Files.ReadWrite.All',
        code: code,
        redirect_uri: redirectUri,
        grant_type: 'authorization_code',
        client_secret: clientSecret
      }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    req.session.accessToken = tokenRes.data.access_token;
    res.redirect('http://theglowup.com.au');
  } catch (err) {
    console.error('Token Error:', err.response?.data || err);
    res.status(500).send('Auth failed');
  }
});

app.post('/calculate', async (req, res) => {
  const accessToken = req.session.accessToken;
  const quantity = req.body.quantity;

  if (!accessToken) return res.status(401).json({ error: 'Not signed in' });

  try {
    // Set A1 value in Excel
    await axios.patch(`https://graph.microsoft.com/v1.0/me/drive/root:/calculator.xlsx:/workbook/worksheets('Sheet1')/range(address='A1')`, {
      values: [[quantity]]
    }, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    // Get B1 value from Excel
    const resultRes = await axios.get(`https://graph.microsoft.com/v1.0/me/drive/root:/calculator.xlsx:/workbook/worksheets('Sheet1')/range(address='B1')`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    const result = resultRes.data.values?.[0]?.[0];
    res.json({ result });
  } catch (err) {
    console.error('Calculation Error:', err.response?.data || err);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
  console.log('==> Your service is live 🎉');
});
