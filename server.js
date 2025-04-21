const express = require('express');
const axios = require('axios');
const querystring = require('querystring');
const cors = require('cors');
const app = express();

app.use(cors());
app.use(express.json());

const PORT = 10000;

// === Credentials (replace only if you regenerate secret) ===
const CLIENT_ID = '1140a629-6ea1-41ec-9655-d5e1afab2408';
const CLIENT_SECRET = 'wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ';
const TENANT_ID = 'common'; // Use 'common' for personal accounts like Outlook.com
const REDIRECT_URI = 'http://localhost:3000/auth/callback';

// === Step 1: Redirect user to sign in ===
app.get('/auth', (req, res) => {
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    response_type: 'code',
    redirect_uri: REDIRECT_URI,
    response_mode: 'query',
    scope: [
      'https://graph.microsoft.com/Files.ReadWrite',
      'https://graph.microsoft.com/User.Read',
      'offline_access',
      'openid',
      'profile'
    ].join(' '),
  });

  const authUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?${params}`;
  res.redirect(authUrl);
});

// === Step 2: Callback from Microsoft after sign-in ===
app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;

  try {
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      querystring.stringify({
        client_id: CLIENT_ID,
        scope: [
          'https://graph.microsoft.com/Files.ReadWrite',
          'https://graph.microsoft.com/User.Read',
          'offline_access',
          'openid',
          'profile'
        ].join(' '),
        code,
        redirect_uri: REDIRECT_URI,
        grant_type: 'authorization_code',
        client_secret: CLIENT_SECRET,
      }),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      }
    );

    const accessToken = tokenResponse.data.access_token;

    res.send(`
      <h2>✅ Authorization successful!</h2>
      <p>Access token obtained.</p>
      <code>${accessToken}</code>
    `);
  } catch (error) {
    console.error('Token error:', error.response?.data || error.message);
    res.status(500).send('❌ Failed to exchange code for token.');
  }
});

app.listen(PORT, () => {
  console.log(`✅ Server listening on port ${PORT}`);
  console.log(`==> Your service is live 🎉`);
});
