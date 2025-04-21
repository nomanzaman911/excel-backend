const express = require('express');
const session = require('express-session');
const axios = require('axios');
const cors = require('cors');
const qs = require('querystring');

const app = express();
const port = 10000;

const CLIENT_ID = '1140a629-6ea1-41ec-9655-d5e1afab2408';
const CLIENT_SECRET = 'wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ';
const REDIRECT_URI = 'http://localhost:10000/auth/callback';
const TENANT_ID = 'common'; // works with personal account
const EXCEL_FILE_PATH = '/calculator.xlsx';
const INPUT_CELL = 'Sheet1!A1';
const RESULT_CELL = 'Sheet1!B1';

app.use(cors({
  origin: 'http://theglowup.com.au',
  credentials: true
}));
app.use(express.json());
app.use(session({ secret: 'excel-secret', resave: false, saveUninitialized: true }));

app.get('/', (req, res) => res.send('✅ Excel API Backend is running'));

app.get('/auth', (req, res) => {
  const params = qs.stringify({
    client_id: CLIENT_ID,
    response_type: 'code',
    redirect_uri: REDIRECT_URI,
    response_mode: 'query',
    scope: 'openid profile offline_access User.Read Files.ReadWrite',
  });
  res.redirect(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?${params}`);
});

app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;
  try {
    const response = await axios.post(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      qs.stringify({
        client_id: CLIENT_ID,
        scope: 'User.Read Files.ReadWrite',
        code,
        redirect_uri: REDIRECT_URI,
        grant_type: 'authorization_code',
        client_secret: CLIENT_SECRET,
      }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );
    req.session.accessToken = response.data.access_token;
    res.redirect('http://theglowup.com.au'); // redirect back to your site
  } catch (err) {
    res.status(500).send('Auth failed');
  }
});

app.post('/calculate', async (req, res) => {
  const accessToken = req.session.accessToken;
  const { input } = req.body;
  try {
    await axios.patch(
      `https://graph.microsoft.com/v1.0/me/drive/root:${EXCEL_FILE_PATH}:/workbook/worksheets('Sheet1')/range(address='${INPUT_CELL}')`,
      { values: [[input]] },
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/root:${EXCEL_FILE_PATH}:/workbook/worksheets('Sheet1')/range(address='${RESULT_CELL}')`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const result = response.data.values[0][0];
    res.json({ result });
  } catch (err) {
    console.error(err.response?.data || err);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
});
