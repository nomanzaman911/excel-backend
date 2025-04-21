import express from "express";
import fetch from "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";

const app = express();
const port = 10000;

// HARDCODED CREDENTIALS
const CLIENT_ID = "53f1b63e-e169-4121-a255-c0a966ca514e";
const TENANT_ID = "6940843a-674d-4941-9ca2-dc5603f278df";
const CLIENT_SECRET = "qmB8Q~phRIvnOQl5R5WHcLxu3~by0Z2pkqGx9cAq";
const EXCEL_FILE_ID = "24C52B39E61CD77F!sc28933942504417abf44a2ea279e2610";
const USER_EMAIL = "nomanzaman@outlook.com"; // This must match Excel file owner

// Auth setup
const credential = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
const graphClient = Client.initWithMiddleware({
  authProvider: {
    getAccessToken: async () => {
      const token = await credential.getToken("https://graph.microsoft.com/.default");
      return token.token;
    },
  },
});

app.get("/", (req, res) => {
  res.send("✅ Excel API Backend is running");
});

app.get("/calculate", async (req, res) => {
  try {
    const quantity = req.query.quantity;
    if (!quantity) return res.status(400).json({ error: "Missing quantity param" });

    // Update Excel A1 with quantity
    await graphClient
      .api(`/users/${USER_EMAIL}/drive/items/${EXCEL_FILE_ID}/workbook/worksheets('Sheet1')/range(address='A1')`)
      .patch({ values: [[quantity]] });

    // Wait briefly for calculation (optional, depending on Excel formula delays)
    await new Promise(resolve => setTimeout(resolve, 1000));

    // Read back calculated result from B1
    const resultResponse = await graphClient
      .api(`/users/${USER_EMAIL}/drive/items/${EXCEL_FILE_ID}/workbook/worksheets('Sheet1')/range(address='B1')`)
      .get();

    console.log("📦 Raw Excel Result:", JSON.stringify(resultResponse, null, 2));

    const result = resultResponse?.values?.[0]?.[0];
    if (result === undefined) return res.status(500).json({ error: "No result returned from Excel" });

    res.json({ result });
  } catch (error) {
    console.error("Error:", JSON.stringify(error, null, 2));
    res.status(500).json({ error: "Failed to calculate result" });
  }
});

app.listen(port, () => {
  console.log(`✅ Server listening on port ${port}`);
});
