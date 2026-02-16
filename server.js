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
   HEALTH
=========================== */

app.get('/health', (req, res) => {
  res.json({ status: 'OK', message: 'Backend running' });
});

/* ===========================
   FETCH ALL ROWS API
=========================== */

app.get('/api/packing-entries', async (req, res) => {
  try {
    const token = await getAccessToken();

    // Dataverse maxes out at 5000 per chunk. We force it to 5000 to speed up the loop.
    const pageSize = 5000;

    /* ======================
       FILTERS
    ====================== */

    const filters = [];

    if (req.query.press)
      filters.push(`cr581_press eq ${req.query.press}`);

    if (req.query.shift)
      filters.push(`cr581_shift eq ${req.query.shift}`);

    if (req.query.status)
      filters.push(`cr581_status eq ${req.query.status}`);

    if (req.query.holdStatus)
      filters.push(`cr581_holdstatus eq ${req.query.holdStatus}`);

    const filterQuery =
      filters.length > 0 ? `&$filter=${filters.join(' and ')}` : '';

    /* ======================
       SELECT ONLY REQUIRED FIELDS
    ====================== */

    const selectFields = [
      'cr581_masterid',
      'cr581_press',
      'cr581_packingdate',
      'cr581_shift',
      'cr581_orderno',
      'cr581_uniqueid',
      'cr581_sectionno',
      'cr581_dieno',
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
      'cr581_sectiongroup',
      'cr581_productionpress',
      'cr581_hardness',
      'cr581_category',
      'createdon'
    ].join(',');

    /* ======================
       BUILD INITIAL URL
    ====================== */

    let url =
      `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}` +
      `?$select=${selectFields}` +
      `&$top=${pageSize}` +
      `&$count=true` +
      filterQuery;

    /* ======================
       CALL DATAVERSE (LOOP UNTIL ALL DATA IS FETCHED)
    ====================== */

    let allRecords = [];
    let totalCount = 0;

    // This loop continues fetching the next 5000 rows as long as Dataverse provides a nextLink
    while (url) {
      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
        }
      });

      // Add the new chunk of records to our master array
      const batch = response.data.value || [];
      allRecords = allRecords.concat(batch);

      // Save the total count reported by Dataverse on the first page
      if (!totalCount && response.data['@odata.count']) {
        totalCount = response.data['@odata.count'];
      }

      // Grab the exact URL for the next chunk (null if we have reached the end)
      url = response.data['@odata.nextLink'] || null;
    }

    // Send the complete, combined dataset back to the React portal
    res.json({
      data: allRecords,
      total: totalCount || allRecords.length
    });

  } catch (error) {
    console.error(error.response?.data || error.message);

    res.status(500).json({
      error: 'Dataverse fetch failed',
      details: error.response?.data || error.message
    });
  }
});

/* ===========================
   START SERVER
=========================== */

app.listen(PORT, () => {
  console.log(`Backend running on port ${PORT}`);
});