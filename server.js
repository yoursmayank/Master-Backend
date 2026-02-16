const express = require('express');
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
const SCOPE = `${DATAVERSE_URL}/.default`;

app.use(cors());
app.use(express.json());

/* ===========================
   GET ACCESS TOKEN
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

/* ===========================
   HEALTH ROUTES
=========================== */

app.get('/health', function (req, res) {
  res.status(200).json({
    status: 'OK',
    message: 'Backend is running',
    timestamp: new Date()
  });
});

app.get('/api/health', function (req, res) {
  res.json({ status: 'Backend running' });
});

/* ===========================
   PAGINATED + FILTERED FETCH
=========================== */

app.get('/api/packing-entries', async function (req, res) {
  try {
    const token = await getAccessToken();

    // Pagination
    const page = parseInt(req.query.page) || 1;
    const pageSize = parseInt(req.query.pageSize) || 100;
    const skip = (page - 1) * pageSize;

    // Filtering
    let filters = [];

    if (req.query.press) {
      filters.push(`cr581_press eq '${req.query.press}'`);
    }

    if (req.query.shift) {
      filters.push(`cr581_shift eq '${req.query.shift}'`);
    }

    if (req.query.status) {
      filters.push(`cr581_status eq '${req.query.status}'`);
    }

    if (req.query.holdStatus) {
      filters.push(`cr581_holdstatus eq '${req.query.holdStatus}'`);
    }

    const filterQuery = filters.length > 0
      ? `$filter=${filters.join(' and ')}`
      : '';

    // Only required columns (IMPORTANT for performance)
    const selectFields = [
      'cr581_masterid',
      'cr581_press',
      'cr581_packingdate',
      'cr581_shift',
      'cr581_orderno',
      'cr581_uniqueid',
      'cr581_sectionno',
      'cr581_sectionname',
      'cr581_sectionsize',
      'cr581_cutlength',
      'cr581_units',
      'cr581_bundleno',
      'cr581_pcs',
      'cr581_bundleweight',
      'cr581_range',
      'cr581_holdstatus',
      'cr581_holdto',
      'cr581_status',
      'cr581_dispatchedto',
      'cr581_dispatchdate',
      'cr581_preparedfor',
      'cr581_productiondate',
      'cr581_productionshift',
      'cr581_productionsupervisor',
      'cr581_packingsupervisor',
      'cr581_sectiongroup',
      'cr581_productionpress',
      'cr581_hardness',
      'cr581_category',
      'createdon'
    ].join(',');

    const url = `
      ${DATAVERSE_URL}/api/data/v9.2/cr581_masters
      ?$select=${selectFields}
      &$count=true
      &$top=${pageSize}
      &$skip=${skip}
      ${filterQuery ? '&' + filterQuery : ''}
    `.replace(/\s+/g, '');

    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'Prefer': 'odata.include-annotations="*"'
      }
    });

    res.json({
      data: response.data.value,
      total: response.data['@odata.count'],
      page,
      pageSize
    });

  } catch (error) {
    console.error('Dataverse Error:', error.response?.data || error.message);

    res.status(500).json({
      error: 'Dataverse fetch failed',
      details: error.response?.data || error.message
    });
  }
});

/* ===========================
   START SERVER
=========================== */

app.listen(PORT, function () {
  console.log(`Backend running on port ${PORT}`);
});
