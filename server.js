const express = require("express");
const session = require("express-session");
const axios = require("axios");
const cors = require("cors");
const qs = require("querystring");

const app = express();
const port = 10000;

const clientId = "1140a629-6ea1-41ec-9655-d5e1afab2408";
const clientSecret = "wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ";
const tenantId = "common"; // Use 'common' for personal Microsoft accounts
const redirectUri = "http://localhost:3000/auth/callback"; // This must match Azure app settings
const excelFilePath = "/drive/root:/calculator.xlsx"; // Your file is in OneDrive root

app.use(express.json());

// ✅ CORS FIX
app.use(cors({
  origin: "http://theglowup.com.au",
  credentials: true
}));

// ✅ SESSION
app.use(session({
  secret: "keyboard cat",
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: false,
    httpOnly: true,
    sameSite: "lax"
  }
}));

// ✅ Home
app.get("/", (req, res) => {
  res.send("✅ Excel API Backend is running");
});

// ✅ Step 1: Redirect to Microsoft Login
app.get("/auth", (req, res) => {
  const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?` +
    qs.stringify({
      client_id: clientId,
      response_type: "code",
      redirect_uri: redirectUri,
      response_mode: "query",
      scope: "Files.ReadWrite offline_access",
    });
  res.redirect(authUrl);
});

// ✅ Step 2: Callback from Microsoft with auth code
app.get("/auth/callback", async (req, res) => {
  const code = req.query.code;
  try {
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      qs.stringify({
        client_id: clientId,
        scope: "Files.ReadWrite offline_access",
        code: code,
        redirect_uri: redirectUri,
        grant_type: "authorization_code",
        client_secret: clientSecret,
      }),
      { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );

    req.session.accessToken = tokenResponse.data.access_token;
    res.redirect("http://theglowup.com.au"); // ✅ Go back to your frontend
  } catch (err) {
    console.error("Token exchange failed", err.response?.data || err.message);
    res.status(500).send("Auth failed");
  }
});

// ✅ Calculate endpoint (POST)
app.post("/calculate", async (req, res) => {
  const token = req.session.accessToken;
  const quantity = req.body.quantity;

  if (!token) {
    return res.status(401).json({ error: "Not authenticated" });
  }

  try {
    // Write input to Excel
    await axios.patch(
      `https://graph.microsoft.com/v1.0/me${excelFilePath}:/workbook/worksheets('Sheet1')/range(address='A1')`,
      { values: [[quantity]] },
      { headers: { Authorization: `Bearer ${token}` } }
    );

    // Read calculated result from Excel
    const resultResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/me${excelFilePath}:/workbook/worksheets('Sheet1')/range(address='B1')`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const result = resultResponse.data.values[0][0];
    res.json({ result });
  } catch (err) {
    console.error("Excel interaction failed", err.response?.data || err.message);
    res.status(500).json({ error: "Failed to calculate result" });
  }
});

// ✅ Start the server
app.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
  console.log("==> Your service is live 🎉");
});
