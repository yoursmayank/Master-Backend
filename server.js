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

// Packing entries entity logical name (Dataverse entity set)
const ENTITY_NAME = process.env.PACKING_ENTITY || 'cr7e4_inventory_records';

// Customer table credentials
const CUSTOMER_TENANT_ID = process.env.TENANT_ID_CUSTOMERS || TENANT_ID;
const CUSTOMER_CLIENT_ID = process.env.CLIEND_ID_CUSTOMERS || CLIENT_ID;
const CUSTOMER_CLIENT_SECRET = process.env.CLIEND_SECRET_CUSTOMERS || CLIENT_SECRET;

// Section Masters credentials
const SECTION_CLIENT_ID = process.env.CLIENT_ID_SECTIONS || CLIENT_ID;
const SECTION_CLIENT_SECRET = process.env.CLIENT_SECRET_SECTIONS || CLIENT_SECRET;

// Die Masters credentials
const DIE_CLIENT_ID = process.env.CLIENT_ID_DIE || CLIENT_ID;
const DIE_CLIENT_SECRET = process.env.CLIENT_SECRET_DIE || CLIENT_SECRET;

// Orders Headers credentials
const ORDERS_HEADERS_CLIENT_ID = process.env.CLIENT_ID_ORDERS_HEADERS || CLIENT_ID;
const ORDERS_HEADERS_CLIENT_SECRET = process.env.CLIENT_SECRET_ORDERS_HEADERS || CLIENT_SECRET;

// Orders Lines credentials
const ORDERS_LINES_CLIENT_ID = process.env.CLIENT_ID_ORDERS_LINES || CLIENT_ID;
const ORDERS_LINES_CLIENT_SECRET = process.env.CLIENT_SECRET_ORDERS_LINES || CLIENT_SECRET;

// Foundry Headers credentials
const FOUNDRY_HEADERS_CLIENT_ID = process.env.CLIENT_ID_FOUNDRY_HEADERS || CLIENT_ID;
const FOUNDRY_HEADERS_CLIENT_SECRET = process.env.CLIENT_SECRET_FOUNDRY_HEADERS || CLIENT_SECRET;

// Foundry Lines credentials
const FOUNDRY_LINES_CLIENT_ID = process.env.CLIENT_ID_FOUNDRY_LINES || CLIENT_ID;
const FOUNDRY_LINES_CLIENT_SECRET = process.env.CLIENT_SECRET_FOUNDRY_LINES || CLIENT_SECRET;

// Production Orders credentials
const PRODUCTION_ORDERS_CLIENT_ID = process.env.CLIENT_ID_PRODUCTIONORDERS || CLIENT_ID;
const PRODUCTION_ORDERS_CLIENT_SECRET = process.env.CLIENT_SECRET_PRODUCTIONORDERS || CLIENT_SECRET;

// Breakdown Reports credentials
const BREAKDOWNS_CLIENT_ID = process.env.CLIENT_ID_BREAKDOWNS || CLIENT_ID;
const BREAKDOWNS_CLIENT_SECRET = process.env.CLIENT_SECRET_BREAKDOWNS || CLIENT_SECRET;

// Aging Records Header credentials
const AGING_HEADER_CLIENT_ID = process.env.CLIENT_ID_AGINGHEADER || CLIENT_ID;
const AGING_HEADER_CLIENT_SECRET = process.env.CLIENT_SECRET_AGINGHEADER || CLIENT_SECRET;

// Aging Records Lines credentials
const AGING_LINE_CLIENT_ID = process.env.CLIENT_ID_AGINGLINE || CLIENT_ID;
const AGING_LINE_CLIENT_SECRET = process.env.CLIENT_SECRET_AGINGLINE || CLIENT_SECRET;

// Reference Data credentials
const REFDATA_CLIENT_ID = process.env.CLIENT_ID_REFDATA || CLIENT_ID;
const REFDATA_CLIENT_SECRET = process.env.CLIENT_SECRET_REFDATA || CLIENT_SECRET;

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

async function getCustomerAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${CUSTOMER_TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', CUSTOMER_CLIENT_ID);
  params.append('client_secret', CUSTOMER_CLIENT_SECRET);
  params.append('scope', SCOPE);

  const response = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
}

async function getSectionAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', SECTION_CLIENT_ID);
  params.append('client_secret', SECTION_CLIENT_SECRET);
  params.append('scope', SCOPE);

  const response = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
}

async function getDieAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', DIE_CLIENT_ID);
  params.append('client_secret', DIE_CLIENT_SECRET);
  params.append('scope', SCOPE);

  const response = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
}

async function getOrdersHeadersAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', ORDERS_HEADERS_CLIENT_ID);
  params.append('client_secret', ORDERS_HEADERS_CLIENT_SECRET);
  params.append('scope', SCOPE);

  const response = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
}

async function getOrdersLinesAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', ORDERS_LINES_CLIENT_ID);
  params.append('client_secret', ORDERS_LINES_CLIENT_SECRET);
  params.append('scope', SCOPE);

  const response = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
}

async function getFoundryHeadersAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', FOUNDRY_HEADERS_CLIENT_ID);
  params.append('client_secret', FOUNDRY_HEADERS_CLIENT_SECRET);
  params.append('scope', SCOPE);

  const response = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
}

async function getFoundryLinesAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', FOUNDRY_LINES_CLIENT_ID);
  params.append('client_secret', FOUNDRY_LINES_CLIENT_SECRET);
  params.append('scope', SCOPE);

  const response = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
}

async function getProductionOrdersAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', PRODUCTION_ORDERS_CLIENT_ID);
  params.append('client_secret', PRODUCTION_ORDERS_CLIENT_SECRET);
  params.append('scope', SCOPE);

  const response = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
}

async function getBreakdownsAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', BREAKDOWNS_CLIENT_ID);
  params.append('client_secret', BREAKDOWNS_CLIENT_SECRET);
  params.append('scope', SCOPE);

  const response = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
}

async function getAgingHeaderAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', AGING_HEADER_CLIENT_ID);
  params.append('client_secret', AGING_HEADER_CLIENT_SECRET);
  params.append('scope', SCOPE);

  const response = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
}

async function getAgingLineAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', AGING_LINE_CLIENT_ID);
  params.append('client_secret', AGING_LINE_CLIENT_SECRET);
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
   RESOLVE CUSTOMER GUID BY NAME
   Looks up cr7e4_customers table by cr7e4_name to find the GUID.
   Returns the GUID string or null if not found.
=========================== */
async function resolveCustomerGuid(customerName) {
  try {
    const token = await getCustomerAccessToken();
    const resp = await axios.get(`${DATAVERSE_URL}/api/data/v9.2/cr7e4_customers`, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
      },
      params: {
        $select: 'cr7e4_customerid',
        $filter: `cr7e4_companyname eq '${customerName.replace(/'/g, "''")}'`,
        $top: 1
      }
    });
    const record = resp.data.value?.[0];
    return record?.cr7e4_customerid || null;
  } catch (e) {
    console.error('[resolveCustomerGuid] Failed:', e.response?.data || e.message);
    return null;
  }
}

/* ===========================
   PACKING ENTRIES API
=========================== */
app.get('/api/packing-entries', async (req, res) => {
  try {
    const token = await getAccessToken();

    const skiptokenFromQuery = req.query.nextLink || req.query.skiptoken || null;
    const pagingCookie = req.query.pagingCookie || null;
    const newerThan = req.query.newerThan || null;
    const from = req.query.from || null;
    const to = req.query.to || null;
    const pageSizeRaw = parseInt(req.query.pageSize, 10);
    const pageSize = Number.isFinite(pageSizeRaw) && pageSizeRaw > 0 ? Math.min(pageSizeRaw, 5000) : 5000;

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
        'cr7e4_category',
        'cr7e4_cutlength',
        'cr7e4_clunit',
        '_cr7e4_dispatchitemid_value',
        'createdon',
        'modifiedon',
        'statecode',
        'statuscode'
      ].join(','),
      // Expand production order → and within it expand section master
      // Expand aging record for tensile strength (hardness)
      $expand: 'cr7e4_OrderNumber($select=cr7e4_ordernumber,cr7e4_productiondate,cr7e4_productionshift,cr7e4_productionpress,_cr7e4_dienumber_value,_cr7e4_customer_value,_cr7e4_productionsupervisor_value;$expand=cr7e4_POID($select=cr7e4_wtrangefromc,cr7e4_wtrangetoc;$expand=cr7e4_SectionNumber($select=cr7e4_sectionnumber,cr7e4_sectionname,cr7e4_sectionsize,cr7e4_sectiongroup))),cr7e4_AgingNumber($select=cr7e4_tensilestrength)',
      $orderby: 'createdon desc',
      $count: 'true'
    };

    const addFilter = (expr) => {
      if (!expr) return;
      queryParams.$filter = queryParams.$filter ? `(${queryParams.$filter}) and (${expr})` : expr;
    };
    const toOdataDateTime = (ymd, endOfDay) => {
      if (!ymd) return null;
      if (String(ymd).includes('T')) return ymd;
      return `${ymd}T${endOfDay ? '23:59:59Z' : '00:00:00Z'}`;
    };

    if (newerThan) {
      addFilter(`createdon gt ${newerThan}`);
    }
    if (from || to) {
      const fromIso = toOdataDateTime(from, false);
      const toIso = toOdataDateTime(to, true);
      if (fromIso) addFilter(`cr7e4_packingdate ge ${fromIso}`);
      if (toIso) addFilter(`cr7e4_packingdate le ${toIso}`);
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
        Prefer: `odata.maxpagesize=${pageSize}, odata.include-annotations="OData.Community.Display.V1.FormattedValue"`
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
   DEBUG: INSPECT SECTION MASTER FIELDS
   GET /api/debug-section
=========================== */
app.get('/api/debug-section', async (req, res) => {
  try {
    const token = await getAccessToken();
    const headers = {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
    };
    // Use the section master GUID we already know from debug-order
    const sectionGuid = '6cf0bad4-de22-f111-8341-7c1e523cad1f';
    const resp = await axios.get(`${DATAVERSE_URL}/api/data/v9.2/cr7e4_section_masters(${sectionGuid})`, { headers });
    res.json(resp.data);
  } catch (error) {
    res.status(500).json({ error: error.response?.data || error.message });
  }
});

/* ===========================
   DEBUG: TEST EXPAND NAV PROP NAMES
   GET /api/debug-expand
=========================== */
app.get('/api/debug-expand', async (req, res) => {
  try {
    const token = await getAccessToken();
    const headers = {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
    };

    // Try each candidate nav prop name - first one that works wins
    const candidates = [
      'cr7e4_OrderNumber',
      'cr7e4_ordernumber',
      'cr7e4_ProductionOrder',
      'cr7e4_productionorder',
      'cr7e4_Order',
      'cr7e4_order',
    ];

    const results = {};
    for (const navProp of candidates) {
      try {
        const url = `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}?$top=1&$select=cr7e4_inventory_recordid&$expand=${navProp}`;
        const resp = await axios.get(url, { headers });
        const record = resp.data.value?.[0];
        results[navProp] = record?.[navProp] ? 'SUCCESS - has data' : 'SUCCESS - but null/empty';
      } catch (e) {
        results[navProp] = `FAIL: ${e.response?.data?.error?.message || e.message}`;
      }
    }
    res.json(results);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/* ===========================
   DEBUG: INSPECT RAW INVENTORY RECORD (AGING + CATEGORY)
   GET /api/debug-aging
   Shows raw fields on a sample inventory record including aging expand + category
=========================== */
app.get('/api/debug-aging', async (req, res) => {
  try {
    const token = await getAccessToken();
    const headers = {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
    };

    // Step 1: get one record where aging number is linked
    const findRes = await axios.get(`${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}`, {
      headers,
      params: { $select: 'cr7e4_inventory_recordid,_cr7e4_agingnumber_value,cr7e4_category', $top: 1, $filter: '_cr7e4_agingnumber_value ne null' }
    });
    const record = findRes.data.value?.[0];
    if (!record) return res.json({ message: 'No record found with aging number linked' });

    const agingGuid = record._cr7e4_agingnumber_value;
    const recordId = record.cr7e4_inventory_recordid;

    // Step 2: try expanding the aging record via candidate nav prop names
    const agingCandidates = ['cr7e4_AgingNumber', 'cr7e4_agingnumber', 'cr7e4_Agingnumber'];
    const expandResults = {};
    for (const navProp of agingCandidates) {
      try {
        const r = await axios.get(`${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}(${recordId})`, {
          headers,
          params: { $select: 'cr7e4_inventory_recordid,cr7e4_category', $expand: navProp }
        });
        expandResults[navProp] = r.data[navProp] !== undefined ? { result: r.data[navProp] } : 'key_missing';
      } catch (e) {
        expandResults[navProp] = `FAIL: ${e.response?.data?.error?.message || e.message}`;
      }
    }

    // Step 3: fetch the aging record directly
    let agingDirect = null;
    try {
      const ar = await axios.get(`${DATAVERSE_URL}/api/data/v9.2/cr7e4_aging_records_lineses(${agingGuid})`, { headers });
      agingDirect = ar.data;
    } catch (e) {
      agingDirect = `FAIL: ${e.response?.data?.error?.message || e.message}`;
    }

    res.json({
      inventoryRecord: record,
      agingGuid,
      expandResults,
      agingDirectFetch: agingDirect
    });
  } catch (error) {
    res.status(500).json({ error: error.response?.data || error.message });
  }
});

/* ===========================
   DEBUG: TEST CUSTOMER RESOLUTION + HOLD BIND
   GET /api/debug-customer?name=Universal Material House
=========================== */
app.get('/api/debug-customer', async (req, res) => {
  try {
    const token = await getAccessToken();
    const name = req.query.name || '';
    const headers = {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
    };

    // Step 1: try to list cr7e4_customers with top 3
    let customerList = null;
    try {
      const r = await axios.get(`${DATAVERSE_URL}/api/data/v9.2/cr7e4_customers`, {
        headers,
        params: { $top: 3 }
      });
      customerList = r.data.value;
    } catch (e) {
      customerList = `FAIL: ${e.response?.data?.error?.message || e.message}`;
    }

    // Step 2: resolve by name
    let resolved = null;
    if (name) {
      resolved = await resolveCustomerGuid(name);
    }

    res.json({ customerList, resolvedGuid: resolved, searchedName: name });
  } catch (error) {
    res.status(500).json({ error: error.response?.data || error.message });
  }
});

/* ===========================
   DISPATCHED QTY BY LINE IDS
   GET /api/packing-entries/dispatched-by-lines?lineIds=guid1,guid2,...
   Returns { [lineId]: { pcs: number, kg: number } }
=========================== */
app.get('/api/packing-entries/dispatched-by-lines', async (req, res) => {
  try {
    const raw = (req.query.lineIds || '').trim();
    if (!raw) return res.json({});

    const lineIds = raw.split(',').map(s => s.trim().toLowerCase()).filter(s => /^[0-9a-f-]{36}$/.test(s));
    if (lineIds.length === 0) return res.json({});

    const token = await getAccessToken();
    const result = {};

    // Dataverse has a URL length limit; process in chunks of 15
    const CHUNK = 15;
    for (let i = 0; i < lineIds.length; i += CHUNK) {
      const chunk = lineIds.slice(i, i + CHUNK);
      const filterParts = chunk.map(id => `_cr7e4_dispatchitemid_value eq ${id}`);
      const filter = filterParts.join(' or ');
      const url = `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}`;
      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          Prefer: 'odata.maxpagesize=5000',
        },
        params: {
          $select: 'cr7e4_pcs,cr7e4_bundleweight,_cr7e4_dispatchitemid_value',
          $filter: filter,
          $top: 5000,
        }
      });
      const records = response.data.value || [];
      for (const rec of records) {
        const lineId = String(rec._cr7e4_dispatchitemid_value || '').toLowerCase();
        if (!lineId) continue;
        const pcs = Number(rec.cr7e4_pcs) || 0;
        const kg  = Number(rec.cr7e4_bundleweight) || 0;
        if (!result[lineId]) result[lineId] = { pcs: 0, kg: 0 };
        result[lineId].pcs += pcs;
        result[lineId].kg  += kg;
      }
    }

    res.json(result);
  } catch (error) {
    console.error('[dispatched-by-lines]', error.response?.data || error.message);
    res.status(500).json({ error: error.response?.data || error.message });
  }
});

/* ===========================
   DISPATCHED QTY BY PRODUCTION ORDER IDS
   GET /api/packing-entries/dispatched-by-prod-orders?prodOrderIds=guid1,guid2,...
   Returns { [prodOrderId]: { pcs: number, kg: number, customers: string[] } }
   — finds inventory records linked to those production orders that were dispatched
=========================== */
app.get('/api/packing-entries/dispatched-by-prod-orders', async (req, res) => {
  try {
    const raw = (req.query.prodOrderIds || '').trim();
    if (!raw) return res.json({});

    const prodOrderIds = raw.split(',').map(s => s.trim().toLowerCase()).filter(s => /^[0-9a-f-]{36}$/.test(s));
    if (prodOrderIds.length === 0) return res.json({});

    const token = await getAccessToken();
    const result = {};

    // Process in chunks of 10 to stay within Dataverse URL limits
    const CHUNK = 10;
    for (let i = 0; i < prodOrderIds.length; i += CHUNK) {
      const chunk = prodOrderIds.slice(i, i + CHUNK);
      const filterParts = chunk.map(id => `_cr7e4_ordernumber_value eq ${id}`);
      // Use _cr7e4_dispatchitemid_value ne null as the dispatch signal — more reliable than status code
      const filter = `(${filterParts.join(' or ')}) and _cr7e4_dispatchitemid_value ne null`;
      const url = `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}`;
      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          Prefer: 'odata.maxpagesize=5000, odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
        },
        params: {
          $select: 'cr7e4_pcs,cr7e4_bundleweight,_cr7e4_ordernumber_value,_cr7e4_dispatchedto_value,_cr7e4_dispatchitemid_value,cr7e4_status',
          $filter: filter,
          $top: 5000,
        }
      });
      const records = response.data.value || [];
      console.log(`[dispatched-by-prod-orders] chunk ${i}-${i+chunk.length}: filter="${filter}" → ${records.length} records`);
      if (records.length > 0) console.log('[dispatched-by-prod-orders] sample:', JSON.stringify(records[0]));
      for (const rec of records) {
        const prodOrderId = String(rec._cr7e4_ordernumber_value || '').toLowerCase();
        if (!prodOrderId) continue;
        const pcs = Number(rec.cr7e4_pcs) || 0;
        const kg  = Number(rec.cr7e4_bundleweight) || 0;
        // Customer display name from formatted value annotation
        const customerName = rec['_cr7e4_dispatchedto_value@OData.Community.Display.V1.FormattedValue']
          || String(rec._cr7e4_dispatchedto_value || '');
        if (!result[prodOrderId]) result[prodOrderId] = { pcs: 0, kg: 0, customers: [] };
        result[prodOrderId].pcs += pcs;
        result[prodOrderId].kg  += kg;
        if (customerName && !result[prodOrderId].customers.includes(customerName)) {
          result[prodOrderId].customers.push(customerName);
        }
      }
    }

    res.json(result);
  } catch (error) {
    console.error('[dispatched-by-prod-orders]', error.response?.data || error.message);
    res.status(500).json({ error: error.response?.data || error.message });
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

    const token = await getAccessToken();

    if (updateData.holdTo) {
      const custGuid = await resolveCustomerGuid(updateData.holdTo);
      if (custGuid) {
        updateData["cr7e4_HoldTo@odata.bind"] = `/cr7e4_customers(${custGuid})`;
      }
      delete updateData.holdTo;
    }

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
    if (holdTo) {
      const custGuid = await resolveCustomerGuid(holdTo);
      if (custGuid) {
        updateData["cr7e4_HoldTo@odata.bind"] = `/cr7e4_customers(${custGuid})`;
      }
    } else if (holdTo === '') {
      // Unhold: clear the lookup by setting it to null
      updateData.cr7e4_HoldTo = null;
    }

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
      reason: r.reason?.response?.data?.error?.message || r.reason?.message || 'unknown'
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
   BATCH DISPATCH UPDATE
   Uses Dataverse OData $batch for max throughput (up to 1000 ops per batch request)
=========================== */
app.post('/api/packing-entries/batch-dispatch', async (req, res) => {
  try {
    const { ids, dispatchTo, dispatchDate, category, lineIdMap } = req.body;

    if (!ids || !Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'ids array is required' });
    }

    const token = await getAccessToken();

    const baseData = {};
    baseData.cr7e4_status = 285960001;       // DISPATCHED
    baseData.cr7e4_holdstatus = 285960002;   // N/A

    if (dispatchDate) baseData.cr7e4_dispatchdate = dispatchDate;
    if (category !== undefined) baseData.cr7e4_category = category;

    if (dispatchTo) {
      const custGuid = await resolveCustomerGuid(dispatchTo);
      if (custGuid) {
        baseData["cr7e4_DispatchedTo@odata.bind"] = `/cr7e4_customers(${custGuid})`;
      }
    }

    // lineIdMap: { [srn]: ordersLineGuid } — used in Step 2 to patch DispatchItemID lookup
    const hasLineMap = lineIdMap && typeof lineIdMap === 'object' && Object.keys(lineIdMap).length > 0;
    console.log(`[batch-dispatch] ids=${ids.length}, hasLineMap=${hasLineMap}, mapKeys=${hasLineMap ? Object.keys(lineIdMap).length : 0}`);
    if (hasLineMap && ids.length > 0) {
      console.log(`[batch-dispatch] sample: id="${ids[0]}" -> lineId="${lineIdMap[ids[0]]}"`); 
    }

    const BATCH_LIMIT = 1000; // Dataverse max per $batch
    const allFailed = [];
    let totalSuccess = 0;

    // Split into $batch requests of up to 1000
    for (let start = 0; start < ids.length; start += BATCH_LIMIT) {
      const chunk = ids.slice(start, start + BATCH_LIMIT);
      const batchId = `batch_dispatch_${Date.now()}_${start}`;
      const changesetId = `changeset_dispatch_${Date.now()}_${start}`;

      // Build multipart $batch body
      let batchBody = '';
      batchBody += `--${batchId}\r\n`;
      batchBody += `Content-Type: multipart/mixed; boundary=${changesetId}\r\n\r\n`;

      chunk.forEach((id, idx) => {
        // Base fields only — DispatchItemID is patched separately in Step 2
        const itemBody = JSON.stringify(baseData);

        batchBody += `--${changesetId}\r\n`;
        batchBody += `Content-Type: application/http\r\n`;
        batchBody += `Content-Transfer-Encoding: binary\r\n`;
        batchBody += `Content-ID: ${idx + 1}\r\n\r\n`;
        batchBody += `PATCH ${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}(${id}) HTTP/1.1\r\n`;
        batchBody += `Content-Type: application/json\r\n`;
        batchBody += `If-Match: *\r\n\r\n`;
        batchBody += `${itemBody}\r\n`;
      });

      batchBody += `--${changesetId}--\r\n`;
      batchBody += `--${batchId}--\r\n`;

      try {
        const batchRes = await axios.post(
          `${DATAVERSE_URL}/api/data/v9.2/$batch`,
          batchBody,
          {
            headers: {
              Authorization: `Bearer ${token}`,
              'Content-Type': `multipart/mixed; boundary=${batchId}`,
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0',
              Accept: 'application/json',
            },
            maxContentLength: Infinity,
            maxBodyLength: Infinity,
          }
        );

        // Parse response: if the changeset succeeds, all ops succeed (atomic).
        // If it fails, Dataverse returns 4xx for the changeset — all ops in it fail.
        const resText = typeof batchRes.data === 'string' ? batchRes.data : JSON.stringify(batchRes.data);
        const hasError = resText.includes('"error"') || resText.includes('HTTP/1.1 4') || resText.includes('HTTP/1.1 5');

        if (hasError) {
          // Changeset failed — fall back to individual PATCHes for this chunk
          const fallbackResults = await Promise.allSettled(
            chunk.map(id => axios.patch(
              `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}(${id})`,
              baseData,
              {
                headers: {
                  Authorization: `Bearer ${token}`,
                  'Content-Type': 'application/json',
                  'OData-MaxVersion': '4.0',
                  'OData-Version': '4.0',
                  'If-Match': '*'
                }
              }
            ))
          );
          fallbackResults.forEach((r, i) => {
            if (r.status === 'rejected') {
              allFailed.push({ id: chunk[i], reason: r.reason?.response?.data?.error?.message || r.reason?.message || 'unknown' });
            } else {
              totalSuccess++;
            }
          });
        } else {
          totalSuccess += chunk.length;
        }
      } catch (batchErr) {
        // $batch request itself failed — fall back to individual PATCHes
        const fallbackResults = await Promise.allSettled(
          chunk.map(id => axios.patch(
            `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}(${id})`,
            baseData,
            {
              headers: {
                Authorization: `Bearer ${token}`,
                'Content-Type': 'application/json',
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'If-Match': '*'
              }
            }
          ))
        );
        fallbackResults.forEach((r, i) => {
          if (r.status === 'rejected') {
            allFailed.push({ id: chunk[i], reason: r.reason?.response?.data?.error?.message || r.reason?.message || 'unknown' });
          } else {
            totalSuccess++;
          }
        });
      }
    }

    // --- STEP 2: Patch DispatchItemID lookup separately (individual PATCHes) ---
    // Kept separate from the $batch so the dispatch never fails due to lookup errors.
    let lookupPatched = 0;
    const lookupErrors = [];
    if (hasLineMap) {
      const mappedIds = ids.filter(id => lineIdMap[id]);
      console.log(`[batch-dispatch] patching DispatchItemID for ${mappedIds.length} items`);
      const lookupResults = await Promise.allSettled(
        mappedIds.map(id =>
          axios.patch(
            `${DATAVERSE_URL}/api/data/v9.2/${ENTITY_NAME}(${id})`,
            { "cr7e4_DispatchItemID@odata.bind": `/cr7e4_orderses(${lineIdMap[id]})` },
            {
              headers: {
                Authorization: `Bearer ${token}`,
                'Content-Type': 'application/json',
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'If-Match': '*',
                Prefer: 'return=minimal',
              }
            }
          )
        )
      );
      lookupResults.forEach((r, i) => {
        if (r.status === 'rejected') {
          const errMsg = r.reason?.response?.data?.error?.message || r.reason?.message || 'unknown';
          lookupErrors.push(errMsg);
          if (i === 0) console.error(`[batch-dispatch] DispatchItemID error sample:`, errMsg, `| id=${mappedIds[i]} | guid=${lineIdMap[mappedIds[i]]}`);
        } else {
          lookupPatched++;
        }
      });
      console.log(`[batch-dispatch] DispatchItemID done: ${lookupPatched} ok, ${lookupErrors.length} errors`);
    }

    res.json({
      success: totalSuccess,
      failed: allFailed.length,
      errors: allFailed,
      lookupPatched,
      lookupErrors: lookupErrors.slice(0, 5),
    });
  } catch (error) {
    res.status(500).json({ error: 'Batch dispatch update failed', details: error.response?.data || error.message });
  }
});

/* ===========================
   SECTION MASTERS API
=========================== */

// Auto-discover section entity set name (Dataverse pluralization varies)
let sectionEntitySet = null;

async function discoverSectionEntitySet(token) {
  const candidates = ['cr7e4_section_masterses', 'cr7e4_section_masters', 'cr7e4_section_masterss'];
  for (const name of candidates) {
    try {
      const resp = await axios.get(`${DATAVERSE_URL}/api/data/v9.2/${name}?$top=1`, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0'
        }
      });
      if (resp.status === 200) {
        console.log(`[sections] Discovered entity set: ${name}`);
        sectionEntitySet = name;
        return name;
      }
    } catch (e) {
      // Try next candidate
    }
  }
  throw new Error('Could not discover section entity set name. Tried: ' + candidates.join(', '));
}

// GET all sections
app.get('/api/sections', async (req, res) => {
  try {
    const token = await getSectionAccessToken();

    if (!sectionEntitySet) {
      await discoverSectionEntitySet(token);
    }

    const url = `${DATAVERSE_URL}/api/data/v9.2/${sectionEntitySet}`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      },
      params: {
        $orderby: 'createdon desc'
      }
    });
    res.json({ data: response.data.value });
  } catch (error) {
    console.error('[GET /api/sections] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch sections', details: error.response?.data || error.message });
  }
});

// GET section entity set name (used by frontend to build OData bind strings directly)
app.get('/api/sections-entityset', async (req, res) => {
  try {
    const token = await getSectionAccessToken();
    if (!sectionEntitySet) {
      await discoverSectionEntitySet(token);
    }
    res.json({ entitySet: sectionEntitySet });
  } catch (error) {
    console.error('[GET /api/sections-entityset] Error:', error.message);
    res.status(500).json({ error: 'Failed to discover section entity set', details: error.message });
  }
});

// POST create a new section
app.post('/api/sections', async (req, res) => {
  try {
    const token = await getSectionAccessToken();
    if (!sectionEntitySet) {
      await discoverSectionEntitySet(token);
    }
    const url = `${DATAVERSE_URL}/api/data/v9.2/${sectionEntitySet}`;
    const response = await axios.post(url, req.body, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'return=representation'
      }
    });
    res.status(201).json(response.data);
  } catch (error) {
    console.error('[POST /api/sections] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to create section', details: error.response?.data || error.message });
  }
});

// PATCH update a section
app.patch('/api/sections/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getSectionAccessToken();
    if (!sectionEntitySet) {
      await discoverSectionEntitySet(token);
    }
    const url = `${DATAVERSE_URL}/api/data/v9.2/${sectionEntitySet}(${id})`;
    await axios.patch(url, req.body, {
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
    console.error('[PATCH /api/sections] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to update section', details: error.response?.data || error.message });
  }
});

// DELETE a section
app.delete('/api/sections/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getSectionAccessToken();
    if (!sectionEntitySet) {
      await discoverSectionEntitySet(token);
    }
    const url = `${DATAVERSE_URL}/api/data/v9.2/${sectionEntitySet}(${id})`;
    await axios.delete(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0'
      }
    });
    res.status(204).send();
  } catch (error) {
    console.error('[DELETE /api/sections] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to delete section', details: error.response?.data || error.message });
  }
});

// POST bulk create sections
// Body: { rows: [ { cr7e4_sectionnumber, cr7e4_sectionname, ... }, ... ] }
app.post('/api/sections/bulk', async (req, res) => {
  const rows = req.body?.rows;
  if (!Array.isArray(rows) || rows.length === 0) {
    return res.status(400).json({ error: 'rows array is required and must not be empty' });
  }
  try {
    const token = await getSectionAccessToken();
    if (!sectionEntitySet) {
      await discoverSectionEntitySet(token);
    }
    const url = `${DATAVERSE_URL}/api/data/v9.2/${sectionEntitySet}`;
    const headers = {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
    };
    const numericFields = ['cr7e4_widthmm', 'cr7e4_heightmm', 'cr7e4_thicknessmm', 'cr7e4_rangefrommm', 'cr7e4_rangetomm'];
    const results = [];
    for (const row of rows) {
      const payload = { ...row };
      numericFields.forEach(f => {
        if (payload[f] !== undefined && payload[f] !== '') {
          const n = parseFloat(payload[f]);
          if (!isNaN(n)) payload[f] = n;
          else delete payload[f];
        } else {
          delete payload[f];
        }
      });
      try {
        const resp = await axios.post(url, payload, { headers: { ...headers, Prefer: 'return=minimal' } });
        results.push({ success: true, sectionnumber: row.cr7e4_sectionnumber });
      } catch (err) {
        results.push({
          success: false,
          sectionnumber: row.cr7e4_sectionnumber,
          error: err.response?.data?.error?.message || err.message
        });
      }
    }
    const failed = results.filter(r => !r.success);
    res.json({ total: rows.length, succeeded: results.length - failed.length, failed: failed.length, results });
  } catch (error) {
    console.error('[POST /api/sections/bulk] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Bulk create failed', details: error.response?.data || error.message });
  }
});

/* ===========================
   DIE MASTERS API
=========================== */

// Auto-discover die entity set name
let dieEntitySet = null;

async function discoverDieEntitySet(token) {
  const candidates = ['cr7e4_die_masterses', 'cr7e4_die_masters', 'cr7e4_die_masterss'];
  for (const name of candidates) {
    try {
      const resp = await axios.get(`${DATAVERSE_URL}/api/data/v9.2/${name}?$top=1`, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0'
        }
      });
      if (resp.status === 200) {
        console.log(`[dies] Discovered entity set: ${name}`);
        dieEntitySet = name;
        return name;
      }
    } catch (e) {
      // Try next candidate
    }
  }
  throw new Error('Could not discover die entity set name. Tried: ' + candidates.join(', '));
}

// GET all dies
app.get('/api/dies', async (req, res) => {
  try {
    const token = await getDieAccessToken();

    if (!dieEntitySet) {
      await discoverDieEntitySet(token);
    }

    const url = `${DATAVERSE_URL}/api/data/v9.2/${dieEntitySet}`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      },
      params: {
        $orderby: 'createdon desc'
      }
    });
    res.json({ data: response.data.value });
  } catch (error) {
    console.error('[GET /api/dies] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch dies', details: error.response?.data || error.message });
  }
});

// POST create a new die
app.post('/api/dies', async (req, res) => {
  try {
    const token = await getDieAccessToken();
    if (!dieEntitySet) {
      await discoverDieEntitySet(token);
    }
    // Ensure section entity set is discovered before building the bind URL
    if (!sectionEntitySet) {
      const sectionToken = await getSectionAccessToken();
      await discoverSectionEntitySet(sectionToken);
    }
    const body = { ...req.body };
    // Transform plain GUID to OData bind navigation property (PascalCase nav prop name)
    if (body.cr7e4_sectionnumber) {
      body['cr7e4_SectionNumber@odata.bind'] = `/${sectionEntitySet}(${body.cr7e4_sectionnumber})`;
      delete body.cr7e4_sectionnumber;
    }
    const url = `${DATAVERSE_URL}/api/data/v9.2/${dieEntitySet}`;
    const response = await axios.post(url, body, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'return=representation'
      }
    });
    res.status(201).json(response.data);
  } catch (error) {
    console.error('[POST /api/dies] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to create die', details: error.response?.data || error.message });
  }
});

// PATCH update a die
app.patch('/api/dies/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getDieAccessToken();
    if (!dieEntitySet) {
      await discoverDieEntitySet(token);
    }
    if (!sectionEntitySet) {
      const sectionToken = await getSectionAccessToken();
      await discoverSectionEntitySet(sectionToken);
    }
    const body = { ...req.body };
    if (body.cr7e4_sectionnumber) {
      body['cr7e4_SectionNumber@odata.bind'] = `/${sectionEntitySet}(${body.cr7e4_sectionnumber})`;
      delete body.cr7e4_sectionnumber;
    }
    const url = `${DATAVERSE_URL}/api/data/v9.2/${dieEntitySet}(${id})`;
    await axios.patch(url, body, {
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
    console.error('[PATCH /api/dies] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to update die', details: error.response?.data || error.message });
  }
});

// DELETE a die
app.delete('/api/dies/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getDieAccessToken();
    if (!dieEntitySet) {
      await discoverDieEntitySet(token);
    }
    const url = `${DATAVERSE_URL}/api/data/v9.2/${dieEntitySet}(${id})`;
    await axios.delete(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0'
      }
    });
    res.status(204).send();
  } catch (error) {
    console.error('[DELETE /api/dies] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to delete die', details: error.response?.data || error.message });
  }
});

/* ===========================
   CUSTOMER MASTER CRUD
=========================== */
const CUSTOMER_ENTITY = 'cr7e4_customers';

// GET all customers
app.get('/api/customers', async (req, res) => {
  try {
    const token = await getCustomerAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/${CUSTOMER_ENTITY}`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      },
      params: {
        $select: 'cr7e4_customerid,cr7e4_companyname,cr7e4_customercode,cr7e4_contactperson,cr7e4_phonenumber,cr7e4_alternatephonenumber,cr7e4_alternatecontactperson,cr7e4_email,cr7e4_addressline1,cr7e4_addressline2,cr7e4_city,cr7e4_district,cr7e4_state,cr7e4_country,cr7e4_pincode,createdon,modifiedon',
        $orderby: 'cr7e4_companyname asc'
      }
    });
    res.json({ data: response.data.value });
  } catch (error) {
    console.error('[GET /api/customers] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch customers', details: error.response?.data || error.message });
  }
});

// POST create a new customer
app.post('/api/customers', async (req, res) => {
  try {
    const token = await getCustomerAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/${CUSTOMER_ENTITY}`;
    const body = req.body;

    const response = await axios.post(url, body, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'return=representation'
      }
    });
    res.status(201).json(response.data);
  } catch (error) {
    console.error('[POST /api/customers] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to create customer', details: error.response?.data || error.message });
  }
});

// PATCH update a customer
app.patch('/api/customers/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getCustomerAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/${CUSTOMER_ENTITY}(${id})`;

    await axios.patch(url, req.body, {
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
    console.error('[PATCH /api/customers] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to update customer', details: error.response?.data || error.message });
  }
});

// DELETE a customer
app.delete('/api/customers/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getCustomerAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/${CUSTOMER_ENTITY}(${id})`;

    await axios.delete(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0'
      }
    });
    res.status(204).send();
  } catch (error) {
    console.error('[DELETE /api/customers] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to delete customer', details: error.response?.data || error.message });
  }
});

/* ===========================
   DEBUG: FIND NAVIGATION PROPERTY NAMES
=========================== */
app.get('/api/debug-navprops', async (req, res) => {
  try {
    const token = await getAccessToken();
    // Try both singular and plural logical names
    const results = {};
    for (const logicalName of ['cr7e4_inventory_record', 'cr7e4_inventory_records']) {
      try {
        const url = `${DATAVERSE_URL}/api/data/v9.2/EntityDefinitions(LogicalName='${logicalName}')/ManyToOneRelationships?$select=SchemaName,ReferencingAttribute,ReferencedEntity,ReferencingEntityNavigationPropertyName`;
        const response = await axios.get(url, {
          headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
        });
        results[logicalName] = response.data.value;
      } catch (e) {
        results[logicalName] = `Error: ${e.response?.data?.error?.message || e.message}`;
      }
    }
    res.json(results);
  } catch (error) {
    res.status(500).json({ error: 'Metadata fetch failed', details: error.response?.data || error.message });
  }
});

/* ===========================
   ORDERS HEADERS API
   Entity: cr7e4_orders_headers
=========================== */

// GET all order headers
app.get('/api/orders-headers', async (req, res) => {
  try {
    const token = await getOrdersHeadersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orders_headers`;
    const queryParams = {
      $orderby: 'createdon desc',
      $count: 'true'
    };
    // Optional: filter by order number
    if (req.query.orderNo) {
      const safe = String(req.query.orderNo).replace(/'/g, "''");
      queryParams.$filter = `cr7e4_ordernumber eq '${safe}'`;
    }
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=5000'
      },
      params: queryParams
    });
    res.json({ data: response.data.value, total: response.data['@odata.count'] || 0 });
  } catch (error) {
    console.error('[GET /api/orders-headers] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch order headers', details: error.response?.data || error.message });
  }
});

// GET single order header by ID
app.get('/api/orders-headers/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getOrdersHeadersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orders_headers(${id})`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      }
    });
    res.json(response.data);
  } catch (error) {
    console.error('[GET /api/orders-headers/:id] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch order header', details: error.response?.data || error.message });
  }
});

// POST create order header
app.post('/api/orders-headers', async (req, res) => {
  try {
    const token = await getOrdersHeadersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orders_headers`;
    const body = { ...req.body };
    // Convert customerId to OData bind
    if (body.customerId) {
      body['cr7e4_Customer@odata.bind'] = `/cr7e4_customers(${body.customerId})`;
      delete body.customerId;
    }
    const response = await axios.post(url, body, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'return=representation'
      }
    });
    res.status(201).json(response.data);
  } catch (error) {
    console.error('[POST /api/orders-headers] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to create order header', details: error.response?.data || error.message });
  }
});

// PATCH update order header
app.patch('/api/orders-headers/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getOrdersHeadersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orders_headers(${id})`;
    await axios.patch(url, req.body, {
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
    console.error('[PATCH /api/orders-headers/:id] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to update order header', details: error.response?.data || error.message });
  }
});

// DELETE order header
app.delete('/api/orders-headers/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getOrdersHeadersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orders_headers(${id})`;
    await axios.delete(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0'
      }
    });
    res.status(204).send();
  } catch (error) {
    console.error('[DELETE /api/orders-headers/:id] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to delete order header', details: error.response?.data || error.message });
  }
});

// DEBUG: Discover all columns on orders_headers
app.get('/api/debug-orders-headers-columns', async (req, res) => {
  try {
    const token = await getOrdersHeadersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orders_headers?$top=1`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      }
    });
    const record = response.data.value?.[0];
    if (!record) return res.json({ message: 'No records found in cr7e4_orders_headers' });
    const columns = Object.keys(record).sort();
    res.json({ totalColumns: columns.length, columns, sampleRecord: record });
  } catch (error) {
    console.error('[DEBUG orders-headers-columns] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch order header columns', details: error.response?.data || error.message });
  }
});

// DEBUG: Navigation properties on orders_headers
app.get('/api/debug-orders-headers-navprops', async (req, res) => {
  try {
    const token = await getOrdersHeadersAccessToken();
    const results = {};
    for (const logicalName of ['cr7e4_orders_header', 'cr7e4_orders_headers']) {
      try {
        const url = `${DATAVERSE_URL}/api/data/v9.2/EntityDefinitions(LogicalName='${logicalName}')/ManyToOneRelationships?$select=SchemaName,ReferencingAttribute,ReferencedEntity,ReferencingEntityNavigationPropertyName`;
        const response = await axios.get(url, {
          headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
        });
        results[logicalName] = response.data.value;
      } catch (e) {
        results[logicalName] = `Error: ${e.response?.data?.error?.message || e.message}`;
      }
    }
    res.json(results);
  } catch (error) {
    res.status(500).json({ error: 'Metadata fetch failed', details: error.response?.data || error.message });
  }
});

/* ===========================
   ORDERS LINES API
   Entity: cr7e4_orderses
=========================== */

// Helper: convert friendly IDs to Dataverse @odata.bind
async function buildOrderLineBody(line, token) {
  const body = { ...line };
  if (body.headerId) {
    body['cr7e4_OrderNumberL@odata.bind'] = `/cr7e4_orders_headers(${body.headerId})`;
    delete body.headerId;
  }
  if (body.sectionId) {
    // Ensure section entity set is discovered
    if (!sectionEntitySet) {
      try { await discoverSectionEntitySet(token); } catch (e) { console.warn('[buildOrderLineBody] Section discovery failed:', e.message); }
    }
    const setName = sectionEntitySet || 'cr7e4_section_masterses';
    body['cr7e4_SectionNumber@odata.bind'] = `/${setName}(${body.sectionId})`;
    delete body.sectionId;
  }
  return body;
}

// GET all order lines
app.get('/api/orders-lines', async (req, res) => {
  try {
    const token = await getOrdersLinesAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orderses`;
    const queryParams = {
      $orderby: 'createdon desc',
      $count: 'true'
    };
    // Optional: filter by order header ID
    if (req.query.headerId) {
      queryParams.$filter = `_cr7e4_ordernumberl_value eq ${req.query.headerId}`;
    }
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=5000'
      },
      params: queryParams
    });
    res.json({ data: response.data.value, total: response.data['@odata.count'] || 0 });
  } catch (error) {
    console.error('[GET /api/orders-lines] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch order lines', details: error.response?.data || error.message });
  }
});

// GET single order line by ID
app.get('/api/orders-lines/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getOrdersLinesAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orderses(${id})`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      }
    });
    res.json(response.data);
  } catch (error) {
    console.error('[GET /api/orders-lines/:id] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch order line', details: error.response?.data || error.message });
  }
});

// POST create order line
app.post('/api/orders-lines', async (req, res) => {
  try {
    const token = await getOrdersLinesAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orderses`;
    const body = await buildOrderLineBody(req.body, token);
    const response = await axios.post(url, body, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'return=representation'
      }
    });
    res.status(201).json(response.data);
  } catch (error) {
    console.error('[POST /api/orders-lines] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to create order line', details: error.response?.data || error.message });
  }
});

// POST batch create order lines (for bulk order submission)
// PUT upsert order lines ? PATCH lines with an existing id, POST new ones (no deletes, preserves POID)
app.put('/api/orders-lines/batch-upsert', async (req, res) => {
  try {
    const { lines } = req.body;
    if (!lines || !Array.isArray(lines) || lines.length === 0) {
      return res.status(400).json({ error: 'lines array is required' });
    }
    const token = await getOrdersLinesAccessToken();
    const baseUrl = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orderses`;

    // Split into updates (have id) and creates (no id)
    const toUpdate = lines.filter(l => l.id);
    const toCreate = lines.filter(l => !l.id);

    // Also collect IDs that are present in the upsert ? lines missing from this list should be deleted
    const incomingIds = new Set(toUpdate.map(l => l.id));
    const headerId = lines[0]?.headerId || null;

    // Delete lines that exist in Dataverse but are NOT in the incoming list
    if (headerId) {
      try {
        const listRes = await axios.get(baseUrl, {
          headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' },
          params: { $select: 'cr7e4_ordersid', $filter: `_cr7e4_ordernumberl_value eq ${headerId}` }
        });
        const existing = listRes.data.value || [];
        const toDelete = existing.filter(e => !incomingIds.has(e.cr7e4_ordersid));
        await Promise.allSettled(toDelete.map(e =>
          axios.delete(`${baseUrl}(${e.cr7e4_ordersid})`, {
            headers: { Authorization: `Bearer ${token}`, 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
          })
        ));
      } catch (e) { console.warn('[batch-upsert] delete-stale failed:', e.message); }
    }

    // PATCH existing lines
    const updateResults = await Promise.allSettled(
      toUpdate.map(async l => {
        const { id, ...rest } = l;
        const body = await buildOrderLineBody(rest, token);
        await axios.patch(`${baseUrl}(${id})`, body, {
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', 'If-Match': '*' }
        });
        return id;
      })
    );

    // POST new lines
    const builtNew = await Promise.all(toCreate.map(l => buildOrderLineBody(l, token)));
    const createResults = await Promise.allSettled(
      builtNew.map(line =>
        axios.post(baseUrl, line, {
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', Prefer: 'return=representation' }
        })
      )
    );

    const successCount = updateResults.filter(r => r.status === 'fulfilled').length + createResults.filter(r => r.status === 'fulfilled').length;
    const failedCount  = updateResults.filter(r => r.status === 'rejected').length  + createResults.filter(r => r.status === 'rejected').length;
    res.json({ success: successCount, failed: failedCount });
  } catch (error) {
    console.error('[PUT /api/orders-lines/batch-upsert] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Batch upsert failed', details: error.response?.data || error.message });
  }
});

// PUT upsert order lines ? PATCH lines with an existing id, POST new ones (no deletes, preserves POID)
app.put('/api/orders-lines/batch-upsert', async (req, res) => {
  try {
    const { lines } = req.body;
    if (!lines || !Array.isArray(lines) || lines.length === 0) {
      return res.status(400).json({ error: 'lines array is required' });
    }
    const token = await getOrdersLinesAccessToken();
    const baseUrl = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orderses`;

    // Split into updates (have id) and creates (no id)
    const toUpdate = lines.filter(l => l.id);
    const toCreate = lines.filter(l => !l.id);

    // Also collect IDs that are present in the upsert ? lines missing from this list should be deleted
    const incomingIds = new Set(toUpdate.map(l => l.id));
    const headerId = lines[0]?.headerId || null;

    // Delete lines that exist in Dataverse but are NOT in the incoming list
    if (headerId) {
      try {
        const listRes = await axios.get(baseUrl, {
          headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' },
          params: { $select: 'cr7e4_ordersid', $filter: `_cr7e4_ordernumberl_value eq ${headerId}` }
        });
        const existing = listRes.data.value || [];
        const toDelete = existing.filter(e => !incomingIds.has(e.cr7e4_ordersid));
        await Promise.allSettled(toDelete.map(e =>
          axios.delete(`${baseUrl}(${e.cr7e4_ordersid})`, {
            headers: { Authorization: `Bearer ${token}`, 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
          })
        ));
      } catch (e) { console.warn('[batch-upsert] delete-stale failed:', e.message); }
    }

    // PATCH existing lines
    const updateResults = await Promise.allSettled(
      toUpdate.map(async l => {
        const { id, ...rest } = l;
        const body = await buildOrderLineBody(rest, token);
        await axios.patch(`${baseUrl}(${id})`, body, {
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', 'If-Match': '*' }
        });
        return id;
      })
    );

    // POST new lines
    const builtNew = await Promise.all(toCreate.map(l => buildOrderLineBody(l, token)));
    const createResults = await Promise.allSettled(
      builtNew.map(line =>
        axios.post(baseUrl, line, {
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', Prefer: 'return=representation' }
        })
      )
    );

    const successCount = updateResults.filter(r => r.status === 'fulfilled').length + createResults.filter(r => r.status === 'fulfilled').length;
    const failedCount  = updateResults.filter(r => r.status === 'rejected').length  + createResults.filter(r => r.status === 'rejected').length;
    res.json({ success: successCount, failed: failedCount });
  } catch (error) {
    console.error('[PUT /api/orders-lines/batch-upsert] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Batch upsert failed', details: error.response?.data || error.message });
  }
});

app.post('/api/orders-lines/batch', async (req, res) => {
  try {
    const { lines } = req.body;
    if (!lines || !Array.isArray(lines) || lines.length === 0) {
      return res.status(400).json({ error: 'lines array is required' });
    }
    const token = await getOrdersLinesAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orderses`;
    // Build all line bodies (converts IDs to binds)
    const builtLines = await Promise.all(lines.map(l => buildOrderLineBody(l, token)));
    const results = await Promise.allSettled(
      builtLines.map(line =>
        axios.post(url, line, {
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            Prefer: 'return=representation'
          }
        })
      )
    );
    const created = results.filter(r => r.status === 'fulfilled').map(r => r.value.data);
    const failed = results.filter(r => r.status === 'rejected').map((r, i) => ({
      index: i,
      reason: r.reason?.response?.data?.error?.message || r.reason?.message || 'unknown'
    }));
    res.json({ success: created.length, failed: failed.length, created, errors: failed });
  } catch (error) {
    console.error('[POST /api/orders-lines/batch] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Batch create failed', details: error.response?.data || error.message });
  }
});

// PATCH update order line
app.patch('/api/orders-lines/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getOrdersLinesAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orderses(${id})`;
    await axios.patch(url, req.body, {
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
    console.error('[PATCH /api/orders-lines/:id] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to update order line', details: error.response?.data || error.message });
  }
});

// DELETE order line
app.delete('/api/orders-lines/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getOrdersLinesAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orderses(${id})`;
    await axios.delete(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0'
      }
    });
    res.status(204).send();
  } catch (error) {
    console.error('[DELETE /api/orders-lines/:id] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to delete order line', details: error.response?.data || error.message });
  }
});

// DELETE all lines for an order (by header ID, with orderNo fallback)
app.delete('/api/orders-lines/by-header/:headerId', async (req, res) => {
  try {
    const { headerId } = req.params;
    const orderNo = req.query.orderNo ? String(req.query.orderNo) : null;
    const token = await getOrdersLinesAccessToken();
    const baseUrl = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orderses`;

    // Build OData filter: match on the lookup field AND/OR the plain orderNo string
    let filter = `_cr7e4_ordernumberl_value eq ${headerId}`;
    if (orderNo) {
      const safe = orderNo.replace(/'/g, "''");
      filter = `(${filter}) or (cr7e4_ordernumber eq '${safe}')`;
    }

    const fetchRes = await axios.get(baseUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
      },
      params: { $select: 'cr7e4_ordersid', $filter: filter }
    });

    const lines = fetchRes.data.value || [];
    if (lines.length === 0) {
      return res.json({ deleted: 0 });
    }

    // De-duplicate in case both filter arms matched the same record
    const uniqueIds = [...new Set(lines.map(l => l.cr7e4_ordersid))];

    await Promise.all(uniqueIds.map(id =>
      axios.delete(`${baseUrl}(${id})`, {
        headers: {
          Authorization: `Bearer ${token}`,
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0'
        }
      })
    ));

    res.json({ deleted: uniqueIds.length });
  } catch (error) {
    console.error('[DELETE /api/orders-lines/by-header] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to delete order lines', details: error.response?.data || error.message });
  }
});

// DEBUG: Discover all columns on cr7e4_orderses (order lines)
app.get('/api/debug-orders-lines-columns', async (req, res) => {
  try {
    const token = await getOrdersLinesAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_orderses?$top=1`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      }
    });
    const record = response.data.value?.[0];
    if (!record) return res.json({ message: 'No records found in cr7e4_orderses' });
    const columns = Object.keys(record).sort();
    res.json({ totalColumns: columns.length, columns, sampleRecord: record });
  } catch (error) {
    console.error('[DEBUG orders-lines-columns] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch order lines columns', details: error.response?.data || error.message });
  }
});

// DEBUG: Navigation properties on cr7e4_orderses
app.get('/api/debug-orders-lines-navprops', async (req, res) => {
  try {
    const token = await getOrdersLinesAccessToken();
    const results = {};
    for (const logicalName of ['cr7e4_orders', 'cr7e4_orderses']) {
      try {
        const url = `${DATAVERSE_URL}/api/data/v9.2/EntityDefinitions(LogicalName='${logicalName}')/ManyToOneRelationships?$select=SchemaName,ReferencingAttribute,ReferencedEntity,ReferencingEntityNavigationPropertyName`;
        const response = await axios.get(url, {
          headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
        });
        results[logicalName] = response.data.value;
      } catch (e) {
        results[logicalName] = `Error: ${e.response?.data?.error?.message || e.message}`;
      }
    }
    res.json(results);
  } catch (error) {
    res.status(500).json({ error: 'Metadata fetch failed', details: error.response?.data || error.message });
  }
});

/* ===========================
   FOUNDRY RECORDS HEADERS API
   Entity: cr7e4_foundry_records_headers
=========================== */

// GET all foundry record headers
app.get('/api/foundry-headers', async (req, res) => {
  try {
    const token = await getFoundryHeadersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_foundry_records_headers`;
    const queryParams = {
      $orderby: 'createdon desc',
      $count: 'true'
    };
    if (req.query.top) queryParams.$top = parseInt(req.query.top, 10);
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=5000'
      },
      params: queryParams
    });
    res.json({ data: response.data.value, total: response.data['@odata.count'] || 0 });
  } catch (error) {
    console.error('[GET /api/foundry-headers] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch foundry headers', details: error.response?.data || error.message });
  }
});

// GET single foundry header by ID
app.get('/api/foundry-headers/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getFoundryHeadersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_foundry_records_headers(${id})`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      }
    });
    res.json(response.data);
  } catch (error) {
    console.error('[GET /api/foundry-headers/:id] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch foundry header', details: error.response?.data || error.message });
  }
});

// DEBUG: Discover all columns on cr7e4_foundry_records_headers
app.get('/api/debug-foundry-headers-columns', async (req, res) => {
  try {
    const token = await getFoundryHeadersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_foundry_records_headers?$top=1`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      }
    });
    const record = response.data.value?.[0];
    if (!record) return res.json({ message: 'No records found in cr7e4_foundry_records_headers' });
    const columns = Object.keys(record).sort();
    res.json({ totalColumns: columns.length, columns, sampleRecord: record });
  } catch (error) {
    console.error('[DEBUG foundry-headers-columns] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch foundry headers columns', details: error.response?.data || error.message });
  }
});

/* ===========================
   FOUNDRY RECORDS LINES API
   Entity: cr7e4_foundry_records_lineses
=========================== */

// GET all foundry record lines
app.get('/api/foundry-lines', async (req, res) => {
  try {
    const token = await getFoundryLinesAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_foundry_records_lineses`;
    const queryParams = {
      $orderby: 'createdon desc',
      $count: 'true'
    };
    if (req.query.top) queryParams.$top = parseInt(req.query.top, 10);
    // Optional: filter by header ID — tries common lookup field naming patterns
    if (req.query.headerId) {
      queryParams.$filter = `_cr7e4_header_value eq ${req.query.headerId}`;
    }
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=5000'
      },
      params: queryParams
    });
    res.json({ data: response.data.value, total: response.data['@odata.count'] || 0 });
  } catch (error) {
    console.error('[GET /api/foundry-lines] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch foundry lines', details: error.response?.data || error.message });
  }
});

// GET single foundry line by ID
app.get('/api/foundry-lines/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getFoundryLinesAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_foundry_records_lineses(${id})`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      }
    });
    res.json(response.data);
  } catch (error) {
    console.error('[GET /api/foundry-lines/:id] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch foundry line', details: error.response?.data || error.message });
  }
});

// DEBUG: Discover all columns on cr7e4_foundry_records_lineses
app.get('/api/debug-foundry-lines-columns', async (req, res) => {
  try {
    const token = await getFoundryLinesAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/cr7e4_foundry_records_lineses?$top=1`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
      }
    });
    const record = response.data.value?.[0];
    if (!record) return res.json({ message: 'No records found in cr7e4_foundry_records_lineses' });
    const columns = Object.keys(record).sort();
    res.json({ totalColumns: columns.length, columns, sampleRecord: record });
  } catch (error) {
    console.error('[DEBUG foundry-lines-columns] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch foundry lines columns', details: error.response?.data || error.message });
  }
});

/* DEBUG: list all attributes on production orders entity */
app.get('/api/debug-production-columns', async (req, res) => {
  try {
    const token = await getProductionOrdersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/EntityDefinitions(LogicalName='cr7e4_production_orders')/Attributes?$select=LogicalName,AttributeType&$filter=AttributeType ne 'Virtual'&$orderby=LogicalName`;
    const response = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
    });
    res.json(response.data.value.map(a => ({ name: a.LogicalName, type: a.AttributeType })));
  } catch (error) {
    res.status(500).json({ error: 'Debug production columns failed', details: error.response?.data || error.message });
  }
});

/* ===========================
   PRODUCTION ORDERS API
   Entity set: cr7e4_production_orderses
   Key fields:
     cr7e4_production_ordersid  – primary key
     cr7e4_ordernumber          – string
     cr7e4_productiondate       – DateTimeOffset
     cr7e4_productionshift      – OptionSet (int)
     cr7e4_productionpress      – string
     cr7e4_cutlength            – Decimal
     cr7e4_clunit               – string / OptionSet
     cr7e4_rangefrom            – Decimal (kg/mm)
     cr7e4_rangeto              – Decimal (kg/mm)
     cr7e4_orderstatus          – OptionSet (int)
     cr7e4_planneddate          – DateTimeOffset
     cr7e4_duedate              – DateTimeOffset
     cr7e4_totalweight          – Decimal
     cr7e4_plnumber             – string
     _cr7e4_sectionnumber_value – lookup GUID
     _cr7e4_dienumber_value     – lookup GUID
     _cr7e4_customer_value      – lookup GUID
     _cr7e4_productionsupervisor_value – lookup GUID
     createdon / modifiedon
=========================== */
const PRODUCTION_ORDERS_ENTITY = 'cr7e4_production_orderses';

const PRODUCTION_ORDERS_SELECT = [
  'cr7e4_production_ordersid',
  'cr7e4_ordernumber',
  'cr7e4_parentordernumber',
  'cr7e4_productiondate',
  'cr7e4_orderdate',
  'cr7e4_dispatchdate',
  'cr7e4_productionshift',
  'cr7e4_productionpress',
  'cr7e4_productionstatus',
  'cr7e4_completedkgs',
  'cr7e4_completedpcs',
  'cr7e4_productionquantity',
  '_cr7e4_poid_value',
  '_cr7e4_dienumber_value',
  '_cr7e4_customer_value',
  '_cr7e4_productionsupervisor_value',
  'statecode',
  'statuscode',
  'createdon',
  'modifiedon',
  'cr7e4_diein',
  'cr7e4_dieout',
  'cr7e4_runtime',
  'cr7e4_billetcount',
  'cr7e4_billetlength',
  'cr7e4_input',
  'cr7e4_weightperpiece',
  'cr7e4_productionremarks',
].join(',');

// GET all production orders
app.get('/api/production-orders', async (req, res) => {
  try {
    const token = await getProductionOrdersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/${PRODUCTION_ORDERS_ENTITY}`;

    const queryParams = {
      $select: PRODUCTION_ORDERS_SELECT,
      $expand: "cr7e4_POID($select=_cr7e4_sectionnumber_value,cr7e4_cutlength,cr7e4_clunit,cr7e4_wtrangefromc,cr7e4_wtrangetoc;$expand=cr7e4_SectionNumber($select=cr7e4_sectionnumber,cr7e4_sectionname,cr7e4_sectionsize))",
      $orderby: 'createdon desc',
      $count: 'true',
    };

    // Optional top-level filters forwarded from frontend query string
    const filters = [];
    const { from, to } = req.query;
    const toOdataDateTime = (ymd, endOfDay) => {
      if (!ymd) return null;
      if (String(ymd).includes('T')) return ymd;
      return `${ymd}T${endOfDay ? '23:59:59Z' : '00:00:00Z'}`;
    };
    if (req.query.orderNumber) {
      const safe = String(req.query.orderNumber).replace(/'/g, "''");
      filters.push(`cr7e4_ordernumber eq '${safe}'`);
    }
    if (req.query.status) {
      const safe = String(req.query.status).replace(/'/g, "''");
      filters.push(`cr7e4_productionstatus eq '${safe}'`);
    }
    if (req.query.shift) {
      filters.push(`cr7e4_productionshift eq ${parseInt(req.query.shift)}`);
    }
    if (from || to) {
      const fromIso = toOdataDateTime(from, false);
      const toIso = toOdataDateTime(to, true);
      if (fromIso) filters.push(`cr7e4_productiondate ge ${fromIso}`);
      if (toIso) filters.push(`cr7e4_productiondate le ${toIso}`);
    }
    if (filters.length > 0) {
      queryParams.$filter = filters.join(' and ');
    }

    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=5000',
      },
      params: queryParams,
    });

    res.json({
      data: response.data.value,
      total: response.data['@odata.count'] || 0,
    });
  } catch (error) {
    console.error('[GET /api/production-orders] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch production orders', details: error.response?.data || error.message });
  }
});

// GET single production order by ID
app.get('/api/production-orders/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getProductionOrdersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/${PRODUCTION_ORDERS_ENTITY}(${id})`;
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
      },
      params: { $select: PRODUCTION_ORDERS_SELECT },
    });
    res.json(response.data);
  } catch (error) {
    console.error('[GET /api/production-orders/:id] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch production order', details: error.response?.data || error.message });
  }
});

// POST create a new production order
app.post('/api/production-orders', async (req, res) => {
  try {
    const token = await getProductionOrdersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/${PRODUCTION_ORDERS_ENTITY}`;
    const response = await axios.post(url, req.body, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'return=representation',
      },
    });
    res.status(201).json(response.data);
  } catch (error) {
    console.error('[POST /api/production-orders] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to create production order', details: error.response?.data || error.message });
  }
});

// PATCH update a production order
app.patch('/api/production-orders/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getProductionOrdersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/${PRODUCTION_ORDERS_ENTITY}(${id})`;
    await axios.patch(url, req.body, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        'If-Match': '*',
      },
    });
    res.status(204).send();
  } catch (error) {
    console.error('[PATCH /api/production-orders/:id] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to update production order', details: error.response?.data || error.message });
  }
});

// DELETE a production order
app.delete('/api/production-orders/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getProductionOrdersAccessToken();
    const url = `${DATAVERSE_URL}/api/data/v9.2/${PRODUCTION_ORDERS_ENTITY}(${id})`;
    await axios.delete(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
      },
    });
    res.status(204).send();
  } catch (error) {
    console.error('[DELETE /api/production-orders/:id] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to delete production order', details: error.response?.data || error.message });
  }
});

/* ===========================
   BREAKDOWN REPORTS API
   Entity: cr7e4_breakdown_reports
=========================== */
const BREAKDOWN_REPORTS_ENTITY = 'cr7e4_breakdown_reportses';

app.get('/api/breakdown-reports', async (req, res) => {
  try {
    const token = await getBreakdownsAccessToken();
    const queryParams = {
      $orderby: 'createdon desc',
      $count: 'true',
    };
    if (req.query.top) queryParams.$top = parseInt(req.query.top, 10);
    if (req.query.filter) queryParams.$filter = req.query.filter;

    const response = await axios.get(
      `${DATAVERSE_URL}/api/data/v9.2/${BREAKDOWN_REPORTS_ENTITY}`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=5000',
        },
        params: queryParams,
      }
    );
    res.json({
      data: response.data.value,
      total: response.data['@odata.count'] || 0,
    });
  } catch (error) {
    console.error('[GET /api/breakdown-reports] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch breakdown reports', details: error.response?.data || error.message });
  }
});

// DEBUG: Discover all columns on cr7e4_breakdown_reports
app.get('/api/debug-breakdown-columns', async (req, res) => {
  try {
    const token = await getBreakdownsAccessToken();
    const response = await axios.get(
      `${DATAVERSE_URL}/api/data/v9.2/${BREAKDOWN_REPORTS_ENTITY}?$top=1`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
        },
      }
    );
    const record = response.data.value?.[0];
    if (!record) return res.json({ message: 'No records found in cr7e4_breakdown_reports' });
    const columns = Object.keys(record).sort();
    res.json({ totalColumns: columns.length, columns, sampleRecord: record });
  } catch (error) {
    console.error('[DEBUG breakdown-columns] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch breakdown columns', details: error.response?.data || error.message });
  }
});

/* ===========================
   AGING RECORDS API
   Headers: cr7e4_aging_records
   Lines:   cr7e4_aging_records_lineses
=========================== */
const AGING_HEADERS_ENTITY = 'cr7e4_aging_recordses';
const AGING_LINES_ENTITY   = 'cr7e4_aging_records_lineses';

// GET all aging headers — newest first
app.get('/api/aging-records', async (req, res) => {
  try {
    const token = await getAgingHeaderAccessToken();
    const queryParams = {
      $orderby: 'createdon desc',
      $count: 'true',
    };
    if (req.query.top) queryParams.$top = parseInt(req.query.top, 10);
    if (req.query.filter) queryParams.$filter = req.query.filter;

    const response = await axios.get(
      `${DATAVERSE_URL}/api/data/v9.2/${AGING_HEADERS_ENTITY}`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=5000',
        },
        params: queryParams,
      }
    );
    res.json({ data: response.data.value, total: response.data['@odata.count'] || 0 });
  } catch (error) {
    console.error('[GET /api/aging-records] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch aging records', details: error.response?.data || error.message });
  }
});

// GET lines for a specific aging header ID
app.get('/api/aging-records/:id/lines', async (req, res) => {
  try {
    const { id } = req.params;
    const token = await getAgingLineAccessToken();

    // Try the most common lookup field name patterns for the parent header
    const filterCandidates = [
      `_cr7e4_agingnumber_value eq ${id}`,
      `_cr7e4_aging_record_value eq ${id}`,
      `_cr7e4_header_value eq ${id}`,
      `_cr7e4_agingid_value eq ${id}`,
    ];

    let data = [];
    let lastError = null;
    for (const filterExpr of filterCandidates) {
      try {
        const response = await axios.get(
          `${DATAVERSE_URL}/api/data/v9.2/${AGING_LINES_ENTITY}`,
          {
            headers: {
              Authorization: `Bearer ${token}`,
              Accept: 'application/json',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0',
              Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=5000',
            },
            params: {
              $filter: filterExpr,
              $orderby: 'createdon asc',
              $count: 'true',
              $expand: 'cr7e4_OrderNumber($select=cr7e4_ordernumber,_cr7e4_customer_value;$expand=cr7e4_POID($select=cr7e4_cutlength,cr7e4_clunit;$expand=cr7e4_SectionNumber($select=cr7e4_sectionnumber,cr7e4_sectionname,cr7e4_sectionsize)))',
            },
          }
        );
        data = response.data.value || [];
        if (data.length > 0) break; // found with this filter
        // Even if 0 results, treat as success and stop trying
        break;
      } catch (e) {
        lastError = e;
        // try next candidate
      }
    }

    if (data.length === 0 && lastError) {
      throw lastError;
    }

    res.json({ data, total: data.length });
  } catch (error) {
    console.error('[GET /api/aging-records/:id/lines] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch aging lines', details: error.response?.data || error.message });
  }
});

// DEBUG: columns on aging headers
app.get('/api/debug-aging-header-columns', async (req, res) => {
  try {
    const token = await getAgingHeaderAccessToken();
    const response = await axios.get(
      `${DATAVERSE_URL}/api/data/v9.2/${AGING_HEADERS_ENTITY}?$top=1`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
        },
      }
    );
    const record = response.data.value?.[0];
    if (!record) return res.json({ message: 'No records found in aging headers' });
    res.json({ totalColumns: Object.keys(record).length, columns: Object.keys(record).sort(), sampleRecord: record });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch aging header columns', details: error.response?.data || error.message });
  }
});

// DEBUG: columns on aging lines
app.get('/api/debug-aging-line-columns', async (req, res) => {
  try {
    const token = await getAgingLineAccessToken();
    const response = await axios.get(
      `${DATAVERSE_URL}/api/data/v9.2/${AGING_LINES_ENTITY}?$top=1`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
        },
      }
    );
    const record = response.data.value?.[0];
    if (!record) return res.json({ message: 'No records found in aging lines' });
    res.json({ totalColumns: Object.keys(record).length, columns: Object.keys(record).sort(), sampleRecord: record });
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch aging line columns', details: error.response?.data || error.message });
  }
});

/* ===========================
   REFERENCE DATA (Plant Report)
=========================== */
async function getRefDataAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');
  params.append('client_id', REFDATA_CLIENT_ID);
  params.append('client_secret', REFDATA_CLIENT_SECRET);
  params.append('scope', SCOPE);
  const response = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });
  return response.data.access_token;
}

app.get('/api/reference-data', async (req, res) => {
  try {
    const token = await getRefDataAccessToken();
    const queryParams = {
      $orderby: 'cr7e4_date asc',
      $count: 'true',
    };
    // Optional date-range filter: ?from=YYYY-MM-DD&to=YYYY-MM-DD
    const { from, to } = req.query;
    const filters = [];
    if (from) filters.push(`cr7e4_date ge ${from}`);
    if (to)   filters.push(`cr7e4_date le ${to}`);
    if (filters.length) queryParams.$filter = filters.join(' and ');

    const response = await axios.get(
      `${DATAVERSE_URL}/api/data/v9.2/cr7e4_reference_datas`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue", odata.maxpagesize=5000',
        },
        params: queryParams,
      }
    );
    res.json({ data: response.data.value, total: response.data['@odata.count'] || 0 });
  } catch (error) {
    console.error('[GET /api/reference-data] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch reference data', details: error.response?.data || error.message });
  }
});

// DEBUG: discover all columns on cr7e4_reference_datas
app.get('/api/debug-refdata-columns', async (req, res) => {
  try {
    const token = await getRefDataAccessToken();
    const response = await axios.get(
      `${DATAVERSE_URL}/api/data/v9.2/cr7e4_reference_datas?$top=1`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
        },
      }
    );
    const record = response.data.value?.[0];
    if (!record) return res.json({ message: 'No records found in cr7e4_reference_datas' });
    res.json({ totalColumns: Object.keys(record).length, columns: Object.keys(record).sort(), sampleRecord: record });
  } catch (error) {
    console.error('[DEBUG refdata-columns] Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch refdata columns', details: error.response?.data || error.message });
  }
});

app.listen(PORT, () => console.log(`Backend running on port ${PORT}`));