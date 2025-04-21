const express = require("express");
const fetch = require("isomorphic-fetch");
const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

const app = express();
const port = process.env.PORT || 10000;

// Config values (hardcoded)
const CLIENT_ID = "53f1b63e-e169-4121-a255-c0a966ca514e";
const TENANT_ID = "6940843a-674d-4941-9ca2-dc5603f278df";
const CLIENT_SECRET = "qmB8Q~phRIvnOQl5R5WHcLxu3~by0Z2pkqGx9cAq";
const EXCEL_FILE_ID = "24C52B39E61CD77F!sc28933942504417abf44a2ea279e2610";
const USER_EMAIL = "nomanzaman@outlook.com";

// Token retrieval
async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append("client_id", CLIENT_ID);
  params.append("scope", "https://graph.microsoft.com/.default");
  params.append("client_secret", CLIENT_SECRET);
  params.append("grant_type", "client_credentials");

  const response = await fetch(url, {
    method: "POST",
    body: params,
  });

  const data = await response.json();
  return data.access_token;
}

// Graph client setup
async function getGraphClient() {
  const token = await getAccessToken();

  const client = Client.init({
    authProvider: (done) => {
      done(null, token);
    },
  });

  return client;
}

// Route to calculate
app.get("/calculate", async (req, res) => {
  const quantity = parseFloat(req.query.quantity);

  try {
    const client = await getGraphClient();

    // Update input cell (A1)
    await client.api(`/users/${USER_EMAIL}/drive/items/${EXCEL_FILE_ID}/workbook/worksheets('Sheet1')/range(address='A1')`)
      .patch({
        values: [[quantity]],
      });

    // Get result from cell (B1)
    const result = await client
      .api(`/users/${USER_EMAIL}/drive/items/${EXCEL_FILE_ID}/workbook/worksheets('Sheet1')/range(address='B1')`)
      .get();

    console.log("📦 Raw Excel Result:", JSON.stringify(result, null, 2)); // 🧠 Debug log

    const calculatedValue = result?.values?.[0]?.[0];

    if (calculatedValue !== undefined && calculatedValue !== null) {
      res.json({ result: calculatedValue });
    } else {
      res.status(500).json({ error: "No result returned from Excel" });
    }
  } catch (error) {
    console.error("❌ Error:", error);
    res.status(500).json({ error: "Failed to calculate result" });
  }
});

app.get("/", (req, res) => {
  res.send("✅ Excel API Backend is running");
});

app.listen(port, () => {
  console.log(`🚀 Server listening on port ${port}`);
});
