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

// Updated to your new table
const ENTITY_NAME = 'cr7e4_inventory_records'; 

// Generated orders table (kept as is)
const ORDERS_ENTITY_NAME = 'cr581_generatedorderses';

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
   PACKING ENTRIES API (cr7e4_inventory_records)
=========================== */
app.get('/api/packing-entries', async (req, res) => {
  try {
    const token = await getAccessToken();
    
    const skiptokenFromQuery = req.query.nextLink || req.query.skiptoken || null;
    const pagingCookie        = req.query.pagingCookie || null;
    const newerThan           = req.query.newerThan || null;

    // Updated $select to use cr7e4_ columns from your JSON
    const queryParams = {
        $select: [
          'cr7e4_inventory_recordid',
          'cr7e4_uniqueid', 
          'cr7e4_packingdate', 
          'cr7e4_packingshift',
          'cr7e4_packingpress', 
          'cr7e4_bundlenumber', 
          'cr7e4_pcs', 
          'cr7e4_bundleweight', 
          'cr7e4_range', 
          'cr7e4_holdstatus', 
          '_cr7e4_holdto_value', 
          'cr7e4_status', 
          '_cr7e4_dispatchedto_value', 
          'cr7e4_dispatchdate', 
          '_cr7e4_ordernumber_value', 
          '_cr7e4_packingsupervisor_value', 
          '_cr7e4_agingnumber_value',
          'createdon',
          'modifiedon',
          'statecode',
          'statuscode'
        ].join(','),
        $orderby: 'createdon desc',
        $count: 'true'
    };

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
      nextLink: nextPageToken || null,
    });

  } catch (error) {
    console.error("Dataverse API Error:", error.response ? JSON.stringify(error.response.data) : error.message);
    res.status(500).json({ error: 'Dataverse fetch failed', details: error.message });
  }
});

/* ===========================
   GENERATED ORDERS API
=========================== */
app.get('/api/generated-orders', async (req, res) => {
  try {
    const token = await getAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/${ORDERS_ENTITY_NAME}`;
    const headers = {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Prefer: 'odata.maxpagesize=5000, odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
    };

    const pageMode = (req.query.page || 'all').toString().toLowerCase();

    const firstParams = {
      $orderby: 'cr581_productiondate desc',
      $count: 'true'
    };

    const firstResponse = await axios.get(url, { headers, params: firstParams });
    const firstPage = firstResponse.data.value || [];
    const totalCount = firstResponse.data['@odata.count'] || firstPage.length;
    let nextUrl = firstResponse.data['@odata.nextLink'] || null;

    console.log(`[generated-orders] Page 1: ${firstPage.length} records (total in Dataverse: ${totalCount})`);

    if (pageMode === '1' || pageMode === 'first') {
      return res.json({
        data: firstPage,
        total: totalCount,
        hasMore: !!nextUrl,
        nextLink: null,
      });
    }

    const allRecords = [...firstPage];
    let page = 1;
    while (nextUrl) {
      page++;
      const nextResponse = await axios.get(nextUrl, { headers });
      const pageRecords = nextResponse.data.value || [];
      allRecords.push(...pageRecords);
      nextUrl = nextResponse.data['@odata.nextLink'] || null;
      console.log(`[generated-orders] Page ${page}: ${pageRecords.length} records (total so far: ${allRecords.length})`);
    }

    console.log(`[generated-orders] Done – ${allRecords.length} total records across ${page} page(s)`);

    res.json({
      data: allRecords,
      total: totalCount,
      hasMore: false,
      nextLink: null,
    });

  } catch (error) {
    console.error("Generated Orders API Error:", error.response ? JSON.stringify(error.response.data) : error.message);
    res.status(500).json({ error: 'Dataverse fetch failed for generated orders', details: error.message });
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
=========================== */
app.post('/api/packing-entries/batch-hold', async (req, res) => {
  try {
    const { ids, holdStatus, holdTo } = req.body;

    if (!ids || !Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'ids array is required' });
    }

    const token = await getAccessToken();

    // Updated to use the new cr7e4_ column prefixes
    const updateData = {};
    if (holdStatus !== undefined) updateData.cr7e4_holdstatus = holdStatus;
    
    // Note: Depending on Dataverse configuration, updating a lookup field via PATCH 
    // might require the @odata.bind syntax (e.g., 'cr7e4_HoldTo@odata.bind': `/systemusers(${holdTo})`)
    // We are keeping it as 'cr7e4_holdto' assuming you are sending standard text/value.
    if (holdTo !== undefined) updateData.cr7e4_holdto = holdTo; 

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