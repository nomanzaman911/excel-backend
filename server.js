const express = require("express");
const session = require("express-session");
const axios = require("axios");
const cors = require("cors");

const app = express();
app.use(express.json());
app.use(cors({ origin: "http://theglowup.com.au", credentials: true }));

app.use(
  session({
    secret: "your_secret_key",
    resave: false,
    saveUninitialized: true
  })
);

const CLIENT_ID = "1140a629-6ea1-41ec-9655-d5e1afab2408";
const CLIENT_SECRET = "wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ";
const TENANT_ID = "6940843a-674d-4941-9ca2-dc5603f278df";
const REDIRECT_URI = "https://excel-backend-1-y1fk.onrender.com/auth/callback";
const FILE_PATH = "/calculator.xlsx"; // Excel file in root of OneDrive

const SCOPES = "offline_access Files.ReadWrite";

app.get("/", (req, res) => {
  res.send("✅ Excel API Backend is running");
});

// Start auth
app.get("/auth", (req, res) => {
  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${CLIENT_ID}&response_type=code&redirect_uri=${encodeURIComponent(
    REDIRECT_URI
  )}&response_mode=query&scope=${encodeURIComponent(SCOPES)}&state=12345`;
  res.redirect(authUrl);
});

// Handle callback
app.get("/auth/callback", async (req, res) => {
  const code = req.query.code;
  if (!code) return res.send("No code received.");

  try {
    const tokenRes = await axios.post(`https://login.microsoftonline.com/common/oauth2/v2.0/token`, new URLSearchParams({
      client_id: CLIENT_ID,
      scope: SCOPES,
      code: code,
      redirect_uri: REDIRECT_URI,
      grant_type: "authorization_code",
      client_secret: CLIENT_SECRET
    }), {
      headers: { "Content-Type": "application/x-www-form-urlencoded" }
    });

    req.session.access_token = tokenRes.data.access_token;
    res.redirect("http://theglowup.com.au");
  } catch (err) {
    console.error("Token error:", err.response.data);
    res.status(500).send("Token exchange failed");
  }
});

// Calculate
app.post("/calculate", async (req, res) => {
  const { quantity } = req.body;
  const token = req.session.access_token;

  if (!token) return res.status(401).send("Not authenticated");

  try {
    // Write value to A1
    await axios.patch(
      `https://graph.microsoft.com/v1.0/me/drive/root:${FILE_PATH}:/workbook/worksheets('Sheet1')/range(address='A1')`,
      { values: [[quantity]] },
      { headers: { Authorization: `Bearer ${token}` } }
    );

    // Read result from B1
    const resultRes = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/root:${FILE_PATH}:/workbook/worksheets('Sheet1')/range(address='B1')`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const result = resultRes.data.values[0][0];
    res.json({ result });
  } catch (err) {
    console.error("Calculation error:", err.response?.data || err.message);
    res.status(500).json({ error: "Failed to calculate result" });
  }
});

app.listen(10000, () => {
  console.log("✅ Server listening on port 10000");
  console.log("==> Your service is live 🎉");
});
