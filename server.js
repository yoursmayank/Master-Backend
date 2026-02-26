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
   INFINITE SCROLL API
=========================== */
app.get('/api/packing-entries', async (req, res) => {
  try {
    const token = await getAccessToken();
    
    // Frontend sends ?nextLink=<skiptoken> on page 2+, or ?pagingCookie=<token>
    // and ?newerThan=<date> for delta sync
    const skiptokenFromQuery = req.query.nextLink || req.query.skiptoken || null;
    const pagingCookie        = req.query.pagingCookie || null;
    const newerThan           = req.query.newerThan || null;

    // Define the base parameters
    const queryParams = {
        $select: [
          'cr581_masterid', 'cr581_press', 'cr581_packingdate', 'cr581_shift',
          'cr581_orderno', 'cr581_uniqueid', 'cr581_sectionno', 'cr581_dieno',
          'cr581_sectionname', 'cr581_sectionsize', 'cr581_cutlength', 'cr581_units',
          'cr581_bundleno', 'cr581_pcs', 'cr581_bundleweight', 'cr581_range',
          'cr581_holdstatus', 'cr581_holdto', 'cr581_status', 'cr581_dispatchedto',
          'cr581_dispatchdate', 'cr581_preparedfor', 'cr581_productiondate',
          'cr581_productionshift', 'cr581_productionsupervisor', 'cr581_sectiongroup',
          'cr581_productionpress', 'cr581_hardness', 'cr581_category', 'createdon'
        ].join(','),
        $orderby: 'createdon desc',
        $count: 'true'
    };

    // Apply delta filter, or pagination token from whichever source
    if (newerThan) {
        queryParams.$filter = `createdon gt ${newerThan}`;
    }
    if (skiptokenFromQuery) {
        queryParams.$skiptoken = skiptokenFromQuery;
    } else if (pagingCookie) {
        queryParams.$skiptoken = pagingCookie;
    }

    const url = `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}`;

    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.maxpagesize=5000, odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      },
      params: queryParams
    });

    // Extract the raw $skiptoken string from the nextLink URL Dataverse returns
    let nextPageToken = null;
    if (response.data['@odata.nextLink']) {
        try {
            const parsedUrl = new URL(response.data['@odata.nextLink']);
            nextPageToken = parsedUrl.searchParams.get('$skiptoken');
        } catch (e) {
            console.warn("Failed to parse nextLink");
        }
    }

    res.json({
      data: response.data.value,
      total: response.data['@odata.count'] || 0,
      // Return as 'nextLink' — this is what the frontend pagination loop reads
      nextLink: nextPageToken || null,
    });

  } catch (error) {
    console.error("Dataverse API Error:", error.response ? JSON.stringify(error.response.data) : error.message);
    res.status(500).json({ error: 'Dataverse fetch failed', details: error.message });
  }
});

/* ===========================
   PATCH SINGLE RECORD
=========================== */
app.patch('/api/packing-entries/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const updateData = req.body;

    if (!id || Object.keys(updateData).length === 0) {
      return res.status(400).json({ error: 'Missing id or update data' });
    }

    const token = await getAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}(${id})`;

    await axios.patch(url, updateData, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        'If-Match': '*'
      }
    });

    res.status(204).send();
  } catch (error) {
    console.error("PATCH error:", error.response ? JSON.stringify(error.response.data) : error.message);
    res.status(500).json({ error: 'Update failed', details: error.message });
  }
});

/* ===========================
   BATCH HOLD UPDATE
   Body: { ids: string[], holdStatus: number, holdTo: string }
=========================== */
app.post('/api/packing-entries/batch-hold', async (req, res) => {
  try {
    const { ids, holdStatus, holdTo } = req.body;

    if (!ids || !Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'ids array is required' });
    }

    const token = await getAccessToken();

    const updateData = {};
    if (holdStatus !== undefined) updateData.cr581_holdstatus = holdStatus;
    if (holdTo !== undefined) updateData.cr581_holdto = holdTo;

    const results = await Promise.allSettled(
      ids.map(id =>
        axios.patch(
          `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}(${id})`,
          updateData,
          {
            headers: {
              Authorization: `Bearer ${token}`,
              'Content-Type': 'application/json',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0',
              'If-Match': '*'
            }
          }
        )
      )
    );

    const failed = results.filter(r => r.status === 'rejected').map((r, i) => ({
      id: ids[i],
      reason: r.reason?.message || 'unknown'
    }));

    if (failed.length > 0) {
      console.warn(`[batch-hold] ${failed.length}/${ids.length} failed:`, failed);
    }

    res.json({
      success: ids.length - failed.length,
      failed: failed.length,
      errors: failed
    });
  } catch (error) {
    console.error("batch-hold error:", error.response ? JSON.stringify(error.response.data) : error.message);
    res.status(500).json({ error: 'Batch hold update failed', details: error.message });
  }
});

/* ===========================
   BATCH DELETE
   Body: { ids: string[] }
=========================== */
app.post('/api/packing-entries/batch-delete', async (req, res) => {
  try {
    const { ids } = req.body;

    if (!ids || !Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'ids array is required' });
    }

    const token = await getAccessToken();

    const results = await Promise.allSettled(
      ids.map(id =>
        axios.delete(
          `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}(${id})`,
          {
            headers: {
              Authorization: `Bearer ${token}`,
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0'
            }
          }
        )
      )
    );

    const failed = results.filter(r => r.status === 'rejected').map((r, i) => ({
      id: ids[i],
      reason: r.reason?.message || 'unknown'
    }));

    res.json({
      success: ids.length - failed.length,
      failed: failed.length,
      errors: failed
    });
  } catch (error) {
    console.error("batch-delete error:", error.response ? JSON.stringify(error.response.data) : error.message);
    res.status(500).json({ error: 'Batch delete failed', details: error.message });
  }
});

app.listen(PORT, () => console.log(`Backend running on port ${PORT}`));