const express = require('express');
const compression = require('compression');
const axios = require('axios');
const cors = require('cors');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 10000;

/* ===========================
   ENV VALIDATION
=========================== */
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
const ENTITY_NAME = 'cr581_masters';
const SCOPE = DATAVERSE_URL + '/.default';

app.use(compression()); 
app.use(cors());
app.use(express.json());

/* ===========================
   ACCESS TOKEN
=========================== */
async function getAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
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

app.get('/health', (req, res) => {
  res.json({ status: 'OK', message: 'Backend running' });
});

/* ===========================
   INFINITE SCROLL & DELTA SYNC API
=========================== */
app.get('/api/packing-entries', async (req, res) => {
  try {
    const token = await getAccessToken();
    
    // 1. Look for the pagination link sent by the React frontend
    const nextLinkParam = req.query.nextLink || null;
    const newerThan = req.query.newerThan || null;

    let url;

    if (nextLinkParam) {
      // 2. If the frontend gave us a nextLink, use it EXACTLY as Dataverse provided it
      url = nextLinkParam;
    } else {
      // 3. Otherwise, build the base query for the very first 5,000 records
      const filters = [];
      
      // Support Delta Fetching (Only get brand new records)
      if (newerThan) filters.push(`createdon gt ${newerThan}`);

      const filterQuery = filters.length > 0 ? `&$filter=${filters.join(' and ')}` : '';

      const selectFields = [
        'cr581_masterid', 'cr581_press', 'cr581_packingdate', 'cr581_shift',
        'cr581_orderno', 'cr581_uniqueid', 'cr581_sectionno', 'cr581_dieno',
        'cr581_sectionname', 'cr581_sectionsize', 'cr581_cutlength', 'cr581_units',
        'cr581_bundleno', 'cr581_pcs', 'cr581_bundleweight', 'cr581_range',
        'cr581_holdstatus', 'cr581_holdto', 'cr581_status', 'cr581_dispatchedto',
        'cr581_dispatchdate', 'cr581_preparedfor', 'cr581_productiondate',
        'cr581_productionshift', 'cr581_productionsupervisor', 'cr581_sectiongroup',
        'cr581_productionpress', 'cr581_hardness', 'cr581_category', 'createdon'
      ].join(',');

      // Base URL grabs newest first (DESC)
      url = `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}?$select=${selectFields}&$orderby=createdon desc&$count=true${filterQuery}`;
    }

    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.maxpagesize=5000, odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      }
    });

    // 4. Pass the data AND the raw Dataverse nextLink back to the React frontend
    res.json({
      data: response.data.value,
      total: response.data['@odata.count'] || 0,
      nextLink: response.data['@odata.nextLink'] || null 
    });

  } catch (error) {
    console.error("Dataverse API Error:", error.response?.data || error.message);
    res.status(500).json({ error: 'Dataverse fetch failed' });
  }
});

app.listen(PORT, () => console.log(`Backend running on port ${PORT}`));