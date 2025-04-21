const express = require('express');
const session = require('express-session');
const axios = require('axios');
const querystring = require('querystring');
const cors = require('cors');

const app = express();
const port = 10000;

app.use(express.json());
app.use(cors({
  origin: 'http://theglowup.com.au',
  credentials: true
}));
app.use(session({ secret: 'secret', resave: false, saveUninitialized: true }));

const CLIENT_ID = '1140a629-6ea1-41ec-9655-d5e1afab2408';
const CLIENT_SECRET = 'wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ';
const REDIRECT_URI = 'http://localhost:3000/auth/callback';
const TENANT_ID = '6940843a-674d-4941-9ca2-dc5603f278df';
const AUTHORITY = `https://login.microsoftonline.com/${TENANT_ID}`;
const SCOPES = 'https://graph.microsoft.com/.default offline_access Files.ReadWrite';

let accessToken = '';

app.get('/', (req, res) => {
  res.send('✅ Excel API Backend is running');
});

app.get('/auth', (req, res) => {
  const authUrl = `${AUTHORITY}/oauth2/v2.0/authorize?` + querystring.stringify({
    client_id: CLIENT_ID,
    response_type: 'code',
    redirect_uri: REDIRECT_URI,
    response_mode: 'query',
    scope: 'offline_access Files.ReadWrite User.Read',
    state: '12345'
  });
  res.redirect(authUrl);
});

app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;

  try {
    const response = await axios.post(`${AUTHORITY}/oauth2/v2.0/token`, querystring.stringify({
      client_id: CLIENT_ID,
      scope: 'offline_access Files.ReadWrite User.Read',
      code,
      redirect_uri: REDIRECT_URI,
      grant_type: 'authorization_code',
      client_secret: CLIENT_SECRET
    }), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });

    accessToken = response.data.access_token;
    req.session.accessToken = accessToken;
    res.redirect('http://theglowup.com.au'); // Redirect back to your frontend
  } catch (error) {
    console.error(error.response?.data || error.message);
    res.send('Authentication failed');
  }
});

app.post('/calculate', async (req, res) => {
  const quantity = req.body.quantity;

  try {
    const headers = { Authorization: `Bearer ${accessToken}` };

    // 1. Write input value to A1
    await axios.patch(`https://graph.microsoft.com/v1.0/me/drive/root:/calculator.xlsx:/workbook/worksheets('Sheet1')/range(address='A1')`, {
      values: [[quantity]]
    }, { headers });

    // 2. Read calculated result from B1
    const resultRes = await axios.get(`https://graph.microsoft.com/v1.0/me/drive/root:/calculator.xlsx:/workbook/worksheets('Sheet1')/range(address='B1')`, { headers });

    const result = resultRes.data.values[0][0];
    res.json({ result });
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
  console.log('==> Your service is live 🎉');
});
