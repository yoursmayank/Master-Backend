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
const requiredEnv = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET', 'ORG_URL'];
const missingEnv = requiredEnv.filter(key => !process.env[key]);

if (missingEnv.length > 0) {
  console.error('Missing environment variables:', missingEnv.join(', '));
  process.exit(1);
}

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const DATAVERSE_URL = process.env.ORG_URL;

// Your table
const ENTITY_NAME = 'cr7e4_inventory_records';

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
   PACKING ENTRIES API
=========================== */
app.get('/api/packing-entries', async (req, res) => {
  try {
    const token = await getAccessToken();

    const skiptokenFromQuery = req.query.nextLink || req.query.skiptoken || null;
    const pagingCookie = req.query.pagingCookie || null;
    const newerThan = req.query.newerThan || null;

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
      // Expand the production order relationship to pull section/order fields
      $expand: 'cr7e4_ordernumber',
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
    res.status(500).json({ error: 'Dataverse fetch failed', details: error.response?.data || error.message });
  }
});

/* ===========================
   DEBUG: INSPECT PRODUCTION ORDER FIELDS
   Call after deploy: GET /api/debug-order
   Shows all fields on a sample production order + its section master
=========================== */
app.get('/api/debug-order', async (req, res) => {
  try {
    const token = await getAccessToken();
    // First fetch one packing record to get a real order GUID
    const packingRes = await axios.get(`${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}`, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' },
      params: { $select: '_cr7e4_ordernumber_value', $top: 1, $filter: '_cr7e4_ordernumber_value ne null' }
    });
    const orderGuid = packingRes.data.value?.[0]?._cr7e4_ordernumber_value;
    if (!orderGuid) return res.json({ message: 'No packing record found with an order number linked' });

    // Fetch that production order record with all its fields
    const orderRes = await axios.get(`${DATAVERSE_URL}/api/data/v9.2/cr7e4_production_orderses(${orderGuid})`, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      }
    });
    res.json({ orderGuid, fields: orderRes.data });
  } catch (error) {
    res.status(500).json({ error: 'Debug order fetch failed', details: error.response?.data || error.message });
  }
});

/* ===========================
   DEBUG: NAVIGATION PROPERTIES ON PRODUCTION ORDERS TABLE
   Call after deploy: GET /api/debug-order-navprops
   Shows section master lookup nav prop name
=========================== */
app.get('/api/debug-order-navprops', async (req, res) => {
  try {
    const token = await getAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/EntityDefinitions(LogicalName='cr7e4_production_orders')/ManyToOneRelationships?$select=SchemaName,ReferencingAttribute,ReferencedEntity,ReferencingEntityNavigationPropertyName`;
    const response = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
    });
    res.json(response.data.value);
  } catch (error) {
    res.status(500).json({ error: 'Metadata fetch failed', details: error.response?.data || error.message });
  }
});

/* ===========================
   DEBUG: NAV PROPS ON INVENTORY RECORDS TABLE
   Call after deploy: GET /api/debug-navprops
=========================== */
app.patch('/api/packing-entries/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const updateData = req.body;

    if (!id || Object.keys(updateData).length === 0) {
      return res.status(400).json({ error: 'Missing id or update data' });
    }

    if (updateData.holdTo) {
      updateData["cr7e4_holdto@odata.bind"] = `/systemusers(${updateData.holdTo})`;
      delete updateData.holdTo;
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
    res.status(500).json({ error: 'Update failed', details: error.response?.data || error.message });
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

    const updateData = {};
    if (holdStatus !== undefined) updateData.cr7e4_holdstatus = holdStatus;
    if (holdTo !== undefined) updateData["cr7e4_holdto@odata.bind"] = `/systemusers(${holdTo})`;

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

    res.json({
      success: ids.length - failed.length,
      failed: failed.length,
      errors: failed
    });
  } catch (error) {
    res.status(500).json({ error: 'Batch hold update failed', details: error.response?.data || error.message });
  }
});

/* ===========================
   DEBUG: FIND NAVIGATION PROPERTY NAMES
=========================== */
app.get('/api/debug-navprops', async (req, res) => {
  try {
    const token = await getAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/EntityDefinitions(LogicalName='cr7e4_inventory_records')/ManyToOneRelationships?$select=SchemaName,ReferencingAttribute,ReferencedEntity,ReferencingEntityNavigationPropertyName`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0'
      }
    });
    res.json(response.data.value);
  } catch (error) {
    res.status(500).json({ error: 'Metadata fetch failed', details: error.response?.data || error.message });
  }
});

app.listen(PORT, () => console.log(`Backend running on port ${PORT}`));