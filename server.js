const express = require('express');
const fetch = require('isomorphic-fetch');
const msal = require('@azure/msal-node');
const graph = require('@microsoft/microsoft-graph-client');

const app = express();
app.use(express.json());

const config = {
  auth: {
    clientId: "1140a629-6ea1-41ec-9655-d5e1afab2408",
    authority: "https://login.microsoftonline.com/common",
    clientSecret: "wR18Q~Yo~udBKwLQDdAF~dT2JphoPZFEJKxdMdtJ"
  }
};

const REDIRECT_URI = "http://localhost:3000/auth/callback";
const pca = new msal.ConfidentialClientApplication(config);

// Redirect user to Microsoft login
app.get("/auth", (req, res) => {
  const authCodeUrlParams = {
    scopes: ["https://graph.microsoft.com/.default", "offline_access", "Files.ReadWrite"],
    redirectUri: REDIRECT_URI
  };

  pca.getAuthCodeUrl(authCodeUrlParams).then(url => {
    res.redirect(url);
  });
});

// Receive auth code and exchange for token
let accessToken = null;

app.get("/auth/callback", async (req, res) => {
  const tokenRequest = {
    code: req.query.code,
    scopes: ["https://graph.microsoft.com/.default", "offline_access", "Files.ReadWrite"],
    redirectUri: REDIRECT_URI
  };

  try {
    const response = await pca.acquireTokenByCode(tokenRequest);
    accessToken = response.accessToken;
    res.send("✅ Authentication complete! You can now send calculation requests.");
  } catch (err) {
    console.error("Auth error:", err);
    res.status(500).send("Auth failed");
  }
});

// Excel Calculation API
app.post("/calculate", async (req, res) => {
  if (!accessToken) return res.status(401).send({ error: "Not authenticated" });

  const { quantity } = req.body;

  try {
    const client = graph.Client.init({
      authProvider: (done) => done(null, accessToken)
    });

    // Adjust path to your file and cells
    const drivePath = "/me/drive/root:/calculator.xlsx";
    const worksheet = "Sheet1";
    const inputCell = "A1";
    const outputCell = "B1";

    // Update input value
    await client
      .api(`${drivePath}:/workbook/worksheets('${worksheet}')/range(address='${inputCell}')`)
      .patch({ values: [[quantity]] });

    // Get calculated result
    const result = await client
      .api(`${drivePath}:/workbook/worksheets('${worksheet}')/range(address='${outputCell}')`)
      .get();

    const calculatedValue = result.values?.[0]?.[0];

    if (calculatedValue !== undefined) {
      res.send({ result: calculatedValue });
    } else {
      res.status(500).send({ error: "No result returned from Excel" });
    }
  } catch (err) {
    console.error("Excel error:", err);
    res.status(500).send({ error: "Failed to calculate result" });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`✅ Server listening on port ${PORT}`));
