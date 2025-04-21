import express from 'express';
import fetch from 'node-fetch';
import open from 'open';

const app = express();
const port = 10000;

// Replace these with your actual app credentials
const clientId = '1140a629-6ea1-41ec-9655-d5e1afab2408';
const clientSecret = 'wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ';
const tenantId = '6940843a-674d-4941-9ca2-dc5603f278df';
const redirectUri = 'http://localhost:10000/auth/callback';
const scopes = 'openid offline_access Files.ReadWrite';

let accessToken = null;

app.get('/auth/signin', async (req, res) => {
  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodeURIComponent(redirectUri)}&response_mode=query&scope=${encodeURIComponent(scopes)}`;
  res.redirect(authUrl);
});

app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;

  const tokenRes = await fetch(`https://login.microsoftonline.com/common/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: 'authorization_code',
      code,
      redirect_uri: redirectUri,
      scope: scopes
    })
  });

  const tokenJson = await tokenRes.json();
  accessToken = tokenJson.access_token;

  if (accessToken) {
    res.send('✅ Sign-in complete. You can now use the Excel API!');
  } else {
    res.send('❌ Failed to sign in.');
  }
});

app.get('/excel/read', async (req, res) => {
  if (!accessToken) return res.status(401).send('Unauthorized');

  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/calculator.xlsx:/workbook/worksheets('Sheet1')/range(address='A1')`;

  const graphRes = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  const data = await graphRes.json();
  res.json(data);
});

app.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
  console.log(`🔗 Open http://localhost:${port}/auth/signin to sign in`);
});
