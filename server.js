import express from 'express';
import axios from 'axios';
import bodyParser from 'body-parser';

const app = express();
const port = process.env.PORT || 3000;

app.use(bodyParser.json());
app.use(express.static('public'));

// MS Graph API credentials and config
const CLIENT_ID = '53f1b63e-e169-4121-a255-c0a966ca514e';
const TENANT_ID = '6940843a-674d-4941-9ca2-dc5603f278df';
const CLIENT_SECRET = 'qmB8Q~phRIvnOQl5R5WHcLxu3~by0Z2pkqGx9cAq';
const USER_EMAIL = 'nomanzaman@outlook.com';
const EXCEL_FILE_ID = '24C52B39E61CD77F!sc28933942504417abf44a2ea279e2610';
const EXCEL_ITEM_PATH = '/drive/root:/calculator.xlsx:/workbook';

// Get token
async function getAccessToken() {
  const response = await axios.post(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, new URLSearchParams({
    client_id: CLIENT_ID,
    scope: 'https://graph.microsoft.com/.default',
    client_secret: CLIENT_SECRET,
    grant_type: 'client_credentials'
  }), {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
}

// Update Excel cell and get result
async function calculate(quantity) {
  const token = await getAccessToken();
  const headers = { Authorization: `Bearer ${token}` };

  const updateUrl = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}${EXCEL_ITEM_PATH}/worksheets('Sheet1')/range(address='A2')`;
  await axios.patch(updateUrl, { values: [[quantity]] }, { headers });

  const readUrl = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}${EXCEL_ITEM_PATH}/worksheets('Sheet1')/range(address='B2')`;
  const response = await axios.get(readUrl, { headers });
  return response.data.values[0][0];
}

app.get('/', (req, res) => {
  res.send('Excel API Backend is running');
});

app.get('/calculate', async (req, res) => {
  const quantity = req.query.quantity;
  try {
    const result = await calculate(quantity);
    res.json({ result });
  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to calculate result' });
  }
});

app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
