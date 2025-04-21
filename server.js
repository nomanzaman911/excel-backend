const express = require("express");
const session = require("express-session");
const axios = require("axios");
const qs = require("querystring");
const cors = require("cors");

const app = express();
const port = 10000;

// Replace these with your actual app values
const CLIENT_ID = "1140a629-6ea1-41ec-9655-d5e1afab2408";
const CLIENT_SECRET = "wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ";
const TENANT_ID = "common";
const REDIRECT_URI = "https://excel-backend-1-y1fk.onrender.com/auth/callback";

app.use(cors({ origin: "http://theglowup.com.au", credentials: true }));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(
  session({
    secret: "secret",
    resave: false,
    saveUninitialized: true,
  })
);

app.get("/", (req, res) => {
  res.send("✅ Excel API Backend is running");
});

app.get("/auth", (req, res) => {
  const authUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?client_id=${CLIENT_ID}&response_type=code&redirect_uri=${REDIRECT_URI}&response_mode=query&scope=offline_access%20User.Read%20Files.ReadWrite`;
  res.redirect(authUrl);
});

app.get("/auth/callback", async (req, res) => {
  const code = req.query.code;
  if (!code) return res.send("Auth failed: No code received");

  try {
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      qs.stringify({
        client_id: CLIENT_ID,
        scope: "offline_access User.Read Files.ReadWrite",
        code,
        redirect_uri: REDIRECT_URI,
        grant_type: "authorization_code",
        client_secret: CLIENT_SECRET,
      }),
      {
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
      }
    );
    req.session.accessToken = tokenResponse.data.access_token;
    res.redirect("http://theglowup.com.au");
  } catch (error) {
    console.error("Token error:", error.response.data);
    res.send("Failed to get access token");
  }
});

app.post("/calculate", async (req, res) => {
  const accessToken = req.session.accessToken;
  const quantity = req.body.quantity;

  if (!accessToken) return res.status(401).send("Not authenticated");

  try {
    const filePath = "/calculator.xlsx";

    // Update input cell A1
    await axios.patch(
      `https://graph.microsoft.com/v1.0/me/drive/root:${filePath}:/workbook/worksheets('Sheet1')/range(address='A1')`,
      { values: [[quantity]] },
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    // Read result cell B1
    const resultResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/root:${filePath}:/workbook/worksheets('Sheet1')/range(address='B1')`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const result = resultResponse.data.values[0][0];
    res.json({ result });
  } catch (err) {
    console.error("Calculation error:", err.response?.data || err.message);
    res.status(500).send("Failed to calculate");
  }
});

app.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
});
