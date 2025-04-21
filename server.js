const express = require("express");
const session = require("express-session");
const fetch = require("node-fetch");
const querystring = require("querystring");
const cors = require("cors");
const bodyParser = require("body-parser");

const app = express();
const port = process.env.PORT || 10000;

// Replace with your actual Azure app info:
const CLIENT_ID = "1140a629-6ea1-41ec-9655-d5e1afab2408";
const CLIENT_SECRET = "wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ";
const TENANT_ID = "common"; // Use 'common' for personal accounts
const REDIRECT_URI = "http://localhost:3000/auth/callback"; // Or your FreeHosting URL
const EXCEL_FILE_PATH = "/calculator.xlsx"; // File in root of OneDrive
const EXCEL_WORKSHEET = "Sheet1";
const INPUT_CELL = "A1";
const OUTPUT_CELL = "B1";

app.use(cors());
app.use(bodyParser.json());

app.use(
  session({
    secret: "keyboard cat",
    resave: false,
    saveUninitialized: true,
  })
);

app.get("/", (req, res) => {
  res.send("✅ Excel API Backend is running");
});

// Step 1: Redirect user to Microsoft login
app.get("/auth/login", (req, res) => {
  const params = querystring.stringify({
    client_id: CLIENT_ID,
    response_type: "code",
    redirect_uri: REDIRECT_URI,
    response_mode: "query",
    scope: "openid profile offline_access User.Read Files.ReadWrite",
  });
  res.redirect(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?${params}`);
});

// Step 2: Handle redirect & exchange code for access token
app.get("/auth/callback", async (req, res) => {
  const code = req.query.code;
  const tokenParams = {
    client_id: CLIENT_ID,
    scope: "https://graph.microsoft.com/.default",
    code: code,
    redirect_uri: REDIRECT_URI,
    grant_type: "authorization_code",
    client_secret: CLIENT_SECRET,
  };

  try {
    const tokenRes = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: querystring.stringify(tokenParams),
    });

    const tokenData = await tokenRes.json();
    if (tokenData.error) throw new Error(JSON.stringify(tokenData));

    req.session.access_token = tokenData.access_token;
    console.log("✅ Access token obtained.");

    // ✅ Redirect to frontend after login
    res.redirect("http://theglowup.com.au/index.html"); // <-- Replace this with your real frontend URL
  } catch (err) {
    console.error("❌ Auth error:", err);
    res.status(500).send("Authentication failed.");
  }
});

// Step 3: Handle Excel calculation
app.post("/calculate", async (req, res) => {
  const token = req.session.access_token;
  const quantity = req.body.quantity;

  if (!token) return res.status(401).json({ error: "Not authenticated" });

  try {
    // 1. Update input cell
    await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:${EXCEL_FILE_PATH}:/workbook/worksheets('${EXCEL_WORKSHEET}')/range(address='${INPUT_CELL}')`, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ values: [[quantity]] }),
    });

    // 2. Get result from output cell
    const resultRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:${EXCEL_FILE_PATH}:/workbook/worksheets('${EXCEL_WORKSHEET}')/range(address='${OUTPUT_CELL}')`, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    const resultData = await resultRes.json();
    const result = resultData.values?.[0]?.[0];

    if (result === undefined) {
      return res.status(500).json({ error: "No result returned from Excel" });
    }

    res.json({ result });
  } catch (err) {
    console.error("❌ Excel error:", err);
    res.status(500).json({ error: "Failed to calculate result" });
  }
});

app.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
});
