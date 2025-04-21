const express = require('express');
const session = require('express-session');
const axios = require('axios');
const cors = require('cors');
const querystring = require('querystring');
const app = express();
const port = 10000;

// Allow requests from your frontend domain
app.use(cors({
  origin: 'http://theglowup.com.au',
  credentials: true
}));

app.use(express.json());
app.use(session({
  secret: 'excel-secret',
  resave: false,
  saveUninitialized: true
}));

// Replace with your own values
const CLIENT_ID = '1140a629-6ea1-41ec-9655-d5e1afab2408';
const CLIENT_SECRET = 'wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ';
const REDIRECT_URI = 'http://localhost:3000/auth/callback'; // Same as registered in Azure
const TENANT = 'common'; // or use your tenant ID
const EXCEL_FILE_PATH = '/calculator.xlsx'; // Must be in OneDrive root
const WORKSHEET = 'Sheet1'; // Sheet name

let userAccessToken = null;

// Auth endpoint to redirect to Microsoft login
app.get('/auth', (req, res) => {
  const params = querystring.stringify({
    client_id: CLIENT_ID,
    response_type: 'code',
    redirect_uri: REDIRECT_URI,
    response_mode: 'query',
    scope: 'offline_access Files.ReadWrite User.Read',
  });
  res.redirect(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/authorize?${params}`);
});

// Callback after login
app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;
  if (!code) return res.status(400).send('No code provided');

  try {
    const tokenResponse = await axios.post(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, querystring.stringify({
      client_id: CLIENT_ID,
      scope: 'offline_access Files.ReadWrite User.Read',
      code: code,
      redirect_uri: REDIRECT_URI,
      grant_type: 'authorization_code',
      client_secret: CLIENT_SECRET
    }), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });

    userAccessToken = tokenResponse.data.access_token;
    res.send('✅ Signed in successfully! You can now use the calculator.');
  } catch (err) {
    console.error('Token error:', err.response?.data || err.message);
    res.status(500).send('Error getting token');
  }
});

// Calculation route
app.post('/calculate', async (req, res) => {
  const quantity = req.body.quantity;
  if (!userAccessToken) return res.status(401).send('Not signed in');

  try {
    // Update quantity cell A1
    await axios.patch(`https://graph.microsoft.com/v1.0/me/drive/root:${EXCEL_FILE_PATH}:/workbook/worksheets('${WORKSHEET}')/range(address='A1')`, {
      values: [[quantity]]
    }, {
      headers: { Authorization: `Bearer ${userAccessToken}` }
    });

    // Read calculated result from B1
    const resultResponse = await axios.get(`https://graph.microsoft.com/v1.0/me/drive/root:${EXCEL_FILE_PATH}:/workbook/worksheets('${WORKSHEET}')/range(address='B1')`, {
      headers: { Authorization: `Bearer ${userAccessToken}` }
    });

    const result = resultResponse.data.values[0][0];
    res.json({ result });
  } catch (err) {
    console.error('Calculation error:', err.response?.data || err.message);
    res.status(500).send({ error: 'Failed to calculate result' });
  }
});

// Test route
app.get('/', (req, res) => {
  res.send('✅ Excel API Backend is running');
});

app.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
});
