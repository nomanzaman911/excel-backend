
const TENANT_ID = "6940843a-674d-4941-9ca2-dc5603f278df";
const CLIENT_ID = "53f1b63e-e169-4121-a255-c0a966ca514e";
const CLIENT_SECRET = "qmB8Q~phRIvnOQl5R5WHcLxu3~by0Z2pkqGx9cAq";
const EXCEL_FILE_ID = "24C52B39E61CD77F!sc28933942504417abf44a2ea279e2610";
const USER_EMAIL = "nomanzaman@outlook.com";






const express = require('express');
const axios = require('axios');
const cors = require('cors');
require('dotenv').config();

const app = express();
const port = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());

// Environment variables from Render (or .env locally)
const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  EXCEL_FILE_ID,
  USER_EMAIL // The Microsoft account email where the Excel is stored
} = process.env;

// Function to get access token
async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', CLIENT_ID);
  params.append('client_secret', CLIENT_SECRET);
  params.append('scope', 'https://graph.microsoft.com/.default');

  const response = await axios.post(url, params);
  return response.data.access_token;
}

// Route to update Excel cell and get calculated result
app.get('/calculate', async (req, res) => {
  const quantity = req.query.quantity;

  try {
    const token = await getAccessToken();

    // Update cell A2 with quantity
    await axios.patch(
      `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/items/${EXCEL_FILE_ID}/workbook/worksheets('Sheet1')/range(address='A2')`,
      { values: [[quantity]] },
      { headers: { Authorization: `Bearer ${token}` } }
    );

    // Read result from cell B2 (assumes your formula is there)
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/items/${EXCEL_FILE_ID}/workbook/worksheets('Sheet1')/range(address='B2')`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const calculatedValue = response.data.values[0][0];

    res.json({ result: calculatedValue });
  } catch (err) {
    console.error('Error:', err.response?.data || err.message);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

// Start server
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
