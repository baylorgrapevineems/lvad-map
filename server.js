const express = require('express');
const crypto  = require('crypto');
const path    = require('path');
const fs      = require('fs');

const app  = express();
const PORT = process.env.PORT || 3001;

// ─── STORAGE MODE ─────────────────────────────────────────────────────────────
// Priority: SharePoint > Upstash Redis > local JSON file
const SP = {
  tenantId:     process.env.AZURE_TENANT_ID,
  clientId:     process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  siteUrl:      process.env.SHAREPOINT_SITE_URL,   // e.g. https://bswh.sharepoint.com/sites/EMS
  listName:     process.env.SP_PATIENTS_LIST || 'LVAD_Patients',
};
const USE_SP    = !!(SP.tenantId && SP.clientId && SP.clientSecret && SP.siteUrl);
const USE_KV    = !USE_SP && !!(process.env.UPSTASH_REDIS_REST_URL && process.env.UPSTASH_REDIS_REST_TOKEN);
const DATA_FILE = path.join(__dirname, 'lvad_patients.json');
const HOSP_FILE = path.join(__dirname, 'lvad_hospitals.json');

// ─── HOSPITALS (static reference data) ───────────────────────────────────────
const HOSPITALS = [
  { id: 'dallas-bumc',  name: 'Baylor Univ. Medical Center – Dallas',       lat: 32.7834, lng: -96.7806 },
  { id: 'dallas-hvh',   name: 'BSW Heart & Vascular Hospital – Dallas',     lat: 32.7894, lng: -96.7761 },
  { id: 'grapevine',    name: 'BSW Medical Center – Grapevine',             lat: 32.9337, lng: -97.0817 },
  { id: 'irving',       name: 'BSW Medical Center – Irving',                lat: 32.8648, lng: -97.0108 },
  { id: 'plano',        name: 'BSW Medical Center – Plano',                 lat: 33.0707, lng: -96.8237 },
  { id: 'mckinney',     name: 'BSW Medical Center – McKinney',              lat: 33.1975, lng: -96.7033 },
  { id: 'allsaints-fw', name: 'BSW All Saints – Fort Worth',                lat: 32.7310, lng: -97.3264 },
  { id: 'waxahachie',   name: 'BSW Medical Center – Waxahachie',            lat: 32.4241, lng: -96.8499 },
  { id: 'carrollton',   name: 'BSW Medical Center – Carrollton',            lat: 32.9879, lng: -96.9049 },
  { id: 'sunnyvale',    name: 'BSW Medical Center – Sunnyvale',             lat: 32.7903, lng: -96.5535 },
];

app.use(express.json());
app.use(express.static(__dirname));

// ─── MICROSOFT GRAPH API ──────────────────────────────────────────────────────
let _tokenCache = null;

async function getToken() {
  if (_tokenCache && _tokenCache.expires > Date.now() + 60_000) return _tokenCache.value;
  const res = await fetch(
    `https://login.microsoftonline.com/${SP.tenantId}/oauth2/v2.0/token`,
    {
      method:  'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body:    new URLSearchParams({
        grant_type:    'client_credentials',
        client_id:     SP.clientId,
        client_secret: SP.clientSecret,
        scope:         'https://graph.microsoft.com/.default',
      }).toString(),
    }
  );
  const d = await res.json();
  if (!res.ok) throw new Error(`Azure auth failed: ${d.error_description || d.error}`);
  _tokenCache = { value: d.access_token, expires: Date.now() + d.expires_in * 1000 };
  return _tokenCache.value;
}

async function graph(endpoint, opts = {}) {
  const token = await getToken();
  const url   = endpoint.startsWith('https://') ? endpoint : `https://graph.microsoft.com/v1.0${endpoint}`;
  const res   = await fetch(url, {
    ...opts,
    headers: {
      Authorization:  `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...opts.headers,
    },
  });
  if (res.status === 204) return null;
  const body = await res.json();
  if (!res.ok) throw new Error(`Graph ${res.status}: ${JSON.stringify(body?.error || body)}`);
  return body;
}

// Cache resolved IDs across requests
let _siteId = null;
let _listId = null;

async function getSiteId() {
  if (_siteId) return _siteId;
  const u    = new URL(SP.siteUrl);
  const data = await graph(`/sites/${u.hostname}:${u.pathname}`);
  _siteId    = data.id;
  return _siteId;
}

async function getListId() {
  if (_listId) return _listId;
  const sid  = await getSiteId();
  const data = await graph(`/sites/${sid}/lists/${encodeURIComponent(SP.listName)}`);
  _listId    = data.id;
  return _listId;
}

async function spAllItems() {
  const sid  = await getSiteId();
  const lid  = await getListId();
  const data = await graph(`/sites/${sid}/lists/${lid}/items?$expand=fields&$select=id,fields,lastModifiedDateTime,lastModifiedBy&$top=999`);
  let items  = data.value || [];
  let next   = data['@odata.nextLink'];
  while (next) {
    const more = await graph(next);
    items = items.concat(more.value || []);
    next  = more['@odata.nextLink'];
  }
  return items;
}

async function spFindItem(patientUid) {
  const items = await spAllItems();
  return items.find(i => i.fields?.PatientUID === patientUid) ?? null;
}

// Map SharePoint fields ↔ app patient object
function spToApp(item) {
  const f = item.fields || {};
  return {
    id:                 f.PatientUID || String(item.id),
    name:               f.Title || '',
    phone:              f.Phone || '',
    emergencyContact:   f.EmergencyContact || '',
    lvadDevice:         f.LVADDevice || '',
    lvadSettings:       f.LVADSettings || '',
    hospital:           f.HospitalId || '',
    address:            f.HomeAddress || '',
    lat:                Number(f.Latitude)  || 0,
    lng:                Number(f.Longitude) || 0,
    notes:              f.Notes || '',
    addedAt:            f.AddedAt || f.Created || new Date().toISOString(),
    implantCenter:      f.ImplantCenter || '',
    implantCenterPhone: f.ImplantCenterPhone || '',
    lvadBagLocation:    f.LVADBagLocation || '',
    lastEditedAt:       item.lastModifiedDateTime || '',
    lastEditedBy:       item.lastModifiedBy?.user?.displayName || '',
  };
}

function appToSP(p) {
  return {
    Title:              p.name,
    PatientUID:         p.id,
    Phone:              p.phone || '',
    EmergencyContact:   p.emergencyContact || '',
    LVADDevice:         p.lvadDevice || '',
    LVADSettings:       p.lvadSettings || '',
    HospitalId:         p.hospital || '',
    HomeAddress:        p.address || '',
    Latitude:           p.lat,
    Longitude:          p.lng,
    Notes:              p.notes || '',
    AddedAt:            p.addedAt,
    ImplantCenter:      p.implantCenter || '',
    ImplantCenterPhone: p.implantCenterPhone || '',
    LVADBagLocation:    p.lvadBagLocation || '',
  };
}

// ─── LOCAL STORAGE FALLBACK ───────────────────────────────────────────────────
async function localRead() {
  if (USE_KV) {
    const { Redis } = await import('@upstash/redis');
    const redis = new Redis({ url: process.env.UPSTASH_REDIS_REST_URL, token: process.env.UPSTASH_REDIS_REST_TOKEN });
    return (await redis.get('lvad:patients')) || [];
  }
  if (!fs.existsSync(DATA_FILE)) return [];
  return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
}

async function localWrite(data) {
  if (USE_KV) {
    const { Redis } = await import('@upstash/redis');
    const redis = new Redis({ url: process.env.UPSTASH_REDIS_REST_URL, token: process.env.UPSTASH_REDIS_REST_TOKEN });
    await redis.set('lvad:patients', data);
    return;
  }
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));
}

async function hospExtrasRead() {
  if (USE_KV) {
    const { Redis } = await import('@upstash/redis');
    const redis = new Redis({ url: process.env.UPSTASH_REDIS_REST_URL, token: process.env.UPSTASH_REDIS_REST_TOKEN });
    return (await redis.get('lvad:hospitals')) || {};
  }
  if (!fs.existsSync(HOSP_FILE)) return {};
  return JSON.parse(fs.readFileSync(HOSP_FILE, 'utf8'));
}

async function hospExtrasWrite(data) {
  if (USE_KV) {
    const { Redis } = await import('@upstash/redis');
    const redis = new Redis({ url: process.env.UPSTASH_REDIS_REST_URL, token: process.env.UPSTASH_REDIS_REST_TOKEN });
    await redis.set('lvad:hospitals', data);
    return;
  }
  fs.writeFileSync(HOSP_FILE, JSON.stringify(data, null, 2));
}

// ─── ROUTES ───────────────────────────────────────────────────────────────────
app.get('/api/hospitals', async (_req, res) => {
  try {
    const extras = await hospExtrasRead();
    res.json(HOSPITALS.map(h => ({
      ...h,
      address:      (extras[h.id] || {}).address      || '',
      phone:        (extras[h.id] || {}).phone        || '',
      capabilities: (extras[h.id] || {}).capabilities || '',
    })));
  } catch {
    res.json(HOSPITALS);
  }
});

app.put('/api/hospitals/:id', async (req, res) => {
  try {
    if (!HOSPITALS.find(h => h.id === req.params.id))
      return res.status(404).json({ error: 'Hospital not found' });
    const { address, phone, capabilities } = req.body;
    const extras = await hospExtrasRead();
    extras[req.params.id] = {
      address:      address      || '',
      phone:        phone        || '',
      capabilities: capabilities || '',
    };
    await hospExtrasWrite(extras);
    res.json({ success: true });
  } catch (e) {
    console.error('[PUT /api/hospitals]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// GET all patients
app.get('/api/patients', async (_req, res) => {
  try {
    if (USE_SP) {
      const items = await spAllItems();
      return res.json(items.map(spToApp));
    }
    res.json(await localRead());
  } catch (e) {
    console.error('[GET /api/patients]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// CREATE patient
app.post('/api/patients', async (req, res) => {
  try {
    const { name, phone, emergencyContact, lvadDevice, lvadSettings, hospital, address, lat, lng, notes } = req.body;
    if (!name || lat == null || lng == null || !hospital)
      return res.status(400).json({ error: 'Missing required fields: name, lat, lng, hospital' });

    const { implantCenter, implantCenterPhone, lvadBagLocation } = req.body;
    const patient = {
      id:                 crypto.randomUUID(),
      name,
      phone:              phone              || '',
      emergencyContact:   emergencyContact   || '',
      lvadDevice:         lvadDevice         || '',
      lvadSettings:       lvadSettings       || '',
      hospital,
      address:            address            || '',
      lat:                Number(lat),
      lng:                Number(lng),
      notes:              notes              || '',
      addedAt:            new Date().toISOString(),
      implantCenter:      implantCenter      || '',
      implantCenterPhone: implantCenterPhone || '',
      lvadBagLocation:    lvadBagLocation    || '',
      lastEditedAt:       new Date().toISOString(),
      lastEditedBy:       'Admin',
    };

    if (USE_SP) {
      const sid = await getSiteId();
      const lid = await getListId();
      await graph(`/sites/${sid}/lists/${lid}/items`, {
        method: 'POST',
        body:   JSON.stringify({ fields: appToSP(patient) }),
      });
      return res.json({ success: true, patient });
    }

    const patients = await localRead();
    patients.push(patient);
    await localWrite(patients);
    res.json({ success: true, patient });
  } catch (e) {
    console.error('[POST /api/patients]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// UPDATE patient
app.put('/api/patients/:id', async (req, res) => {
  try {
    if (USE_SP) {
      const item = await spFindItem(req.params.id);
      if (!item) return res.status(404).json({ error: 'Patient not found' });
      const original = spToApp(item);
      const updated  = { ...original, ...req.body, id: req.params.id, addedAt: original.addedAt };
      const sid = await getSiteId();
      const lid = await getListId();
      await graph(`/sites/${sid}/lists/${lid}/items/${item.id}/fields`, {
        method: 'PATCH',
        body:   JSON.stringify(appToSP(updated)),
      });
      return res.json({ success: true, patient: updated });
    }

    const patients = await localRead();
    const idx      = patients.findIndex(p => p.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'Patient not found' });
    patients[idx] = { ...patients[idx], ...req.body, id: patients[idx].id, addedAt: patients[idx].addedAt };
    await localWrite(patients);
    res.json({ success: true, patient: patients[idx] });
  } catch (e) {
    console.error('[PUT /api/patients]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// DELETE patient
app.delete('/api/patients/:id', async (req, res) => {
  try {
    if (USE_SP) {
      const item = await spFindItem(req.params.id);
      if (!item) return res.status(404).json({ error: 'Patient not found' });
      const sid = await getSiteId();
      const lid = await getListId();
      await graph(`/sites/${sid}/lists/${lid}/items/${item.id}`, { method: 'DELETE' });
      return res.json({ success: true });
    }

    const patients = await localRead();
    const idx      = patients.findIndex(p => p.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'Patient not found' });
    patients.splice(idx, 1);
    await localWrite(patients);
    res.json({ success: true, total: patients.length });
  } catch (e) {
    console.error('[DELETE /api/patients]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── START ────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  const mode = USE_SP ? 'SharePoint Online (Microsoft Graph API)' : USE_KV ? 'Upstash Redis' : 'Local JSON file';
  console.log(`\n  LVAD Patient Monitor  →  http://localhost:${PORT}`);
  console.log(`  Admin                 →  http://localhost:${PORT}/admin.html`);
  console.log(`  Storage               →  ${mode}\n`);
  if (!USE_SP) {
    console.log('  SharePoint not configured — running with local storage.');
    console.log('  See .env.example for required environment variables.\n');
  }
});
