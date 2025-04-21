const express = require("express");
const { Client } = require("@microsoft/microsoft-graph-client");
const { ClientSecretCredential } = require("@azure/identity");
require("isomorphic-fetch");

const app = express();
const port = process.env.PORT || 10000;

// Hardcoded config values
const CLIENT_ID = "53f1b63e-e169-4121-a255-c0a966ca514e";
const TENANT_ID = "6940843a-674d-4941-9ca2-dc5603f278df";
const CLIENT_SECRET = "qmB8Q~phRIvnOQl5R5WHcLxu3~by0Z2pkqGx9cAq";
const EXCEL_FILE_ID = "24C52B39E61CD77F!sc28933942504417abf44a2ea279e2610";
const USER_EMAIL = "nomanzaman@outlook.com";

const credential = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

async function getAccessToken() {
  const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");
  return tokenResponse.token;
}

function getGraphClient(token) {
  return Client.init({
    authProvider: done => done(null, token),
  });
}

app.get("/", (req, res) => {
  res.send("Excel API Backend is running");
});

app.get("/calculate", async (req, res) => {
  const quantity = parseFloat(req.query.quantity || 0);
  if (isNaN(quantity)) return res.status(400).json({ error: "Invalid quantity" });

  try {
    const token = await getAccessToken();
    const client = getGraphClient(token);

    const range = "Sheet1!A1";
    await client.api(`/users/${USER_EMAIL}/drive/items/${EXCEL_FILE_ID}/workbook/worksheets/Sheet1/range(address='${range}')`)
      .patch({ values: [[quantity]] });

    const resultRange = "Sheet1!B1";
    const resultResponse = await client.api(`/users/${USER_EMAIL}/drive/items/${EXCEL_FILE_ID}/workbook/worksheets/Sheet1/range(address='${resultRange}')`).get();
    const calculatedValue = resultResponse.values?.[0]?.[0];

    res.json({ result: calculatedValue });
  } catch (error) {
    console.error("Calculation error:", error);
    res.status(500).json({ error: "Failed to calculate result" });
  }
});

app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
