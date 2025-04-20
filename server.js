const express = require('express');
const axios = require('axios');
const dotenv = require('dotenv');
const cors = require('cors');

dotenv.config();

const app = express();
app.use(cors());

const PORT = process.env.PORT || 3000;

async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', process.env.CLIENT_ID);
  params.append('client_secret', process.env.CLIENT_SECRET);
  params.append('scope', 'https://graph.microsoft.com/.default');

  const res = await axios.post(url, params);
  return res.data.access_token;
}

app.get('/calculate', async (req, res) => {
  const quantity = parseInt(req.query.quantity, 10);

  try {
    const token = await getAccessToken();

    // Set A2 to quantity
    await axios.patch(
      `https://graph.microsoft.com/v1.0${process.env.EXCEL_FILE_PATH}/worksheets('Sheet1')/range(address='A2')`,
      { values: [[quantity]] },
      { headers: { Authorization: `Bearer ${token}` } }
    );

    // Read result from B2
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0${process.env.EXCEL_FILE_PATH}/worksheets('Sheet1')/range(address='B2')`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const result = response.data.values?.[0]?.[0];
    res.json({ result });
  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
