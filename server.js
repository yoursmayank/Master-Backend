const express = require('express');
const axios = require('axios');
const cors = require('cors');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3001;

// Validate required environment variables
const requiredEnv = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET'];
const missingEnv = requiredEnv.filter(key => !process.env[key]);

if (missingEnv.length > 0) {
  console.error('Missing environment variables:', missingEnv.join(', '));
  process.exit(1);
}

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const DATAVERSE_URL = 'https://jhalaniextrusion.api.crm8.dynamics.com';
const SCOPE = DATAVERSE_URL + '/.default';

app.use(cors());
app.use(express.json());

async function getAccessToken() {
  const tokenUrl = 'https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token';

  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', CLIENT_ID);
  params.append('client_secret', CLIENT_SECRET);
  params.append('scope', SCOPE);

  const response = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });

  return response.data.access_token;
}

app.get('/api/health', function (req, res) {
  res.json({ status: 'Backend running' });
});

app.get('/api/packing-entries', async function (req, res) {
  try {
    const token = await getAccessToken();

    const response = await axios.get(
      DATAVERSE_URL + '/api/data/v9.2/cr581_masters',
      {
        headers: {
          Authorization: 'Bearer ' + token,
          Accept: 'application/json'
        }
      }
    );

    res.json(response.data.value);
  } catch (error) {
    console.error(error.response?.data || error.message);
    res.status(500).json({
      error: 'Dataverse fetch failed'
    });
  }
});

app.listen(PORT, function () {
  console.log('Backend running at http://localhost:' + PORT);
});
