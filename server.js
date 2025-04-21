const express = require('express');
const session = require('express-session');
const axios = require('axios');
const qs = require('querystring');
const cors = require('cors');

const app = express();
app.use(express.json());

// Allow frontend site CORS
app.use(cors({
  origin: 'http://theglowup.com.au',
  credentials: true
}));

app.use(session({
  secret: 'excel-secret-key',
  resave: false,
  saveUninitialized: true
}));

// Config
const clientId = '1140a629-6ea1-41ec-9655-d5e1afab2408';
const clientSecret = 'wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ';
const tenantId = 'common'; // Use "common" for personal accounts
const redirectUri = 'https://excel-backend-1-y1fk.onrender.com/auth/callback';
const scope = 'offline_access Files.ReadWrite User.Read openid profile';
const excelPath = '/me/drive/root:/calculator.xlsx:/workbook/worksheets/Sheet1';


// Step 1: Redirect user to Microsoft sign-in
app.get('/auth', (req, res) => {
  const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?` +
    qs.stringify({
      client_id: clientId,
      response_type: 'code',
      redirect_uri: redirectUri,
      response_mode: 'query',
      scope
    });
  res.redirect(authUrl);
});

// Step 2: Callback with auth code
app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;
  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  try {
    const tokenRes = await axios.post(tokenUrl, qs.stringify({
      client_id: clientId,
      scope,
      code,
      redirect_uri: redirectUri,
      grant_type: 'authorization_code',
      client_secret: clientSecret
    }), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });

    req.session.accessToken = tokenRes.data.access_token;
    res.redirect('http://theglowup.com.au');
  } catch (err) {
    console.error('Token error:', err.response?.data || err.message);
    res.send('Auth failed');
  }
});

// Auth check
app.get('/auth/status', (req, res) => {
  res.json({ authenticated: !!req.session.accessToken });
});


// Main calculation endpoint
app.post('/calculate', async (req, res) => {
  const token = req.session.accessToken;
  const inputValue = req.body.value;

  if (!token) return res.status(401).json({ error: 'Not authenticated' });

  try {
    // 1. Set A1
    await axios.patch(
      `https://graph.microsoft.com/v1.0${excelPath}/range(address='A1')`,
      { values: [[inputValue]] },
      { headers: { Authorization: `Bearer ${token}` } }
    );

    // 2. Read B1
    const result = await axios.get(
      `https://graph.microsoft.com/v1.0${excelPath}/range(address='B1')`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const calculatedValue = result.data.values[0][0];
    res.json({ result: calculatedValue });

  } catch (err) {
    console.error('Calculation error:', err.response?.data || err.message);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});


// Start the server
app.listen(10000, () => {
  console.log('✅ Server listening on port 10000');
});
