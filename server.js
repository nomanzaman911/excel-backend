const express = require('express');
const session = require('express-session');
const { Issuer } = require('openid-client');
const axios = require('axios');

const app = express();
app.use(express.json());

const clientId = '1140a629-6ea1-41ec-9655-d5e1afab2408';
const clientSecret = 'wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ';
const redirectUri = 'http://localhost:3000/auth/callback'; // Replace with your deployed frontend URI if needed

app.use(
  session({
    secret: 'mysecret',
    resave: false,
    saveUninitialized: true,
  })
);

let client;

(async () => {
  const microsoftIssuer = await Issuer.discover('https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration');
  client = new microsoftIssuer.Client({
    client_id: clientId,
    client_secret: clientSecret,
    redirect_uris: [redirectUri],
    response_types: ['code'],
  });
})();

app.get('/auth', (req, res) => {
  const url = client.authorizationUrl({
    scope: 'openid profile offline_access Files.ReadWrite',
  });
  res.redirect(url);
});

app.get('/auth/callback', async (req, res) => {
  const params = client.callbackParams(req);
  const tokenSet = await client.callback(redirectUri, params);
  req.session.tokenSet = tokenSet;
  res.send('Signed in! You can close this tab and return to the app.');
});

app.post('/calculate', async (req, res) => {
  try {
    const token = req.session.tokenSet?.access_token;
    if (!token) return res.status(401).json({ error: 'Unauthorized' });

    const workbookUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/calculator.xlsx:/workbook/worksheets('Sheet1')/range(address='A1')`;

    await axios.patch(workbookUrl, { values: [[req.body.quantity]] }, {
      headers: { Authorization: `Bearer ${token}` },
    });

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/root:/calculator.xlsx:/workbook/worksheets('Sheet1')/range(address='B1')`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const result = response.data.values[0][0];
    res.json({ result });
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).json({ error: 'Calculation failed' });
  }
});

app.listen(10000, () => {
  console.log('✅ Server listening on port 10000');
});
