const express = require('express');
const session = require('express-session');
const axios = require('axios');
const qs = require('qs');
const cors = require('cors');
const app = express();

const CLIENT_ID = '1140a629-6ea1-41ec-9655-d5e1afab2408';
const CLIENT_SECRET = 'wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ';
const TENANT_ID = '6940843a-674d-4941-9ca2-dc5603f278df';
const REDIRECT_URI = 'https://excel-backend-1-y1fk.onrender.com/auth/callback';
const AUTHORITY = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

app.use(cors({
  origin: 'http://theglowup.com.au',
  credentials: true
}));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: 'excel_secret',
  resave: false,
  saveUninitialized: true
}));

app.get('/', (req, res) => {
  res.send('✅ Excel API Backend is running');
});

app.get('/auth', (req, res) => {
  const authUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?client_id=${CLIENT_ID}&response_type=code&redirect_uri=${encodeURIComponent(REDIRECT_URI)}&response_mode=query&scope=${encodeURIComponent('openid profile offline_access Files.ReadWrite.All User.Read')}`;
  res.redirect(authUrl);
});

app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;
  if (!code) return res.status(400).send('No code provided');

  try {
    const tokenRes = await axios.post(AUTHORITY, qs.stringify({
      client_id: CLIENT_ID,
      scope: 'Files.ReadWrite.All User.Read',
      code: code,
      redirect_uri: REDIRECT_URI,
      grant_type: 'authorization_code',
      client_secret: CLIENT_SECRET
    }), { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } });

    req.session.accessToken = tokenRes.data.access_token;
    res.redirect('http://theglowup.com.au');
  } catch (err) {
    console.error('Token error:', err.response?.data || err.message);
    res.status(500).send('Auth failed');
  }
});

app.post('/calculate', async (req, res) => {
  const token = req.session.accessToken;
  const quantity = req.body.quantity;

  if (!token) return res.status(401).send('Not authenticated');

  try {
    const headers = { Authorization: `Bearer ${token}` };

    // Write input value to A1
    await axios.patch(
      'https://graph.microsoft.com/v1.0/me/drive/root:/calculator.xlsx:/workbook/worksheets/Sheet1/range(address=\'A1\')',
      { values: [[quantity]] },
      { headers }
    );

    // Read output from B1
    const resultRes = await axios.get(
      'https://graph.microsoft.com/v1.0/me/drive/root:/calculator.xlsx:/workbook/worksheets/Sheet1/range(address=\'B1\')',
      { headers }
    );

    const result = resultRes.data.values?.[0]?.[0];
    res.json({ result });
  } catch (err) {
    console.error('Calculation error:', err.response?.data || err.message);
    res.status(500).send('Failed to get result');
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => {
  console.log(`✅ Server listening on port ${PORT}`);
});
