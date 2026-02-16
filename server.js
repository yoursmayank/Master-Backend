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
const ENTITY_NAME = 'cr581_masters';
const SCOPE = DATAVERSE_URL + '/.default';

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

/* ===========================
   INFINITE SCROLL API
=========================== */
app.get('/api/packing-entries', async (req, res) => {
  try {
    const token = await getAccessToken();
    const pagingCookie = req.query.pagingCookie || null;

    /* FILTERS */
    const filters = [];
    if (req.query.press) filters.push(`cr581_press eq ${req.query.press}`);
    if (req.query.shift) filters.push(`cr581_shift eq ${req.query.shift}`);
    if (req.query.status) filters.push(`cr581_status eq ${req.query.status}`);
    if (req.query.holdStatus) filters.push(`cr581_holdstatus eq ${req.query.holdStatus}`);

    const filterQuery = filters.length > 0 ? `&$filter=${filters.join(' and ')}` : '';

    /* SELECT FIELDS */
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

    /* BUILD URL (REMOVED $top BUG) */
    let url = `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}?$select=${selectFields}&$count=true${filterQuery}`;

    if (pagingCookie) {
      url += `&$skiptoken=${encodeURIComponent(pagingCookie)}`;
    }

    /* CALL DATAVERSE */
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        // This is the correct way to set page size without killing pagination
        Prefer: 'odata.maxpagesize=5000, odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      }
    });

    const nextLink = response.data['@odata.nextLink'] || null;
    let nextPagingCookie = null;

    if (nextLink) {
      const parsed = new URL(nextLink);
      nextPagingCookie = parsed.searchParams.get('$skiptoken');
    }

    res.json({
      data: response.data.value,
      total: response.data['@odata.count'] || 0,
      nextPagingCookie: nextPagingCookie
    });

  } catch (error) {
    console.error(error.response?.data || error.message);
    res.status(500).json({ error: 'Dataverse fetch failed' });
  }
});

app.listen(PORT, () => console.log(`Backend running on port ${PORT}`));