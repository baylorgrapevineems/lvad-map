const express = require('express');
const fs      = require('fs');
const path    = require('path');
const crypto  = require('crypto');

const app  = express();
const PORT = process.env.PORT || 3001;

const USE_KV    = !!(process.env.UPSTASH_REDIS_REST_URL && process.env.UPSTASH_REDIS_REST_TOKEN);
const DATA_FILE = path.join(__dirname, 'lvad_patients.json');
const KV_KEY    = 'lvad:patients';

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

async function readPatients() {
  if (USE_KV) {
    const { Redis } = await import('@upstash/redis');
    const redis = new Redis({ url: process.env.UPSTASH_REDIS_REST_URL, token: process.env.UPSTASH_REDIS_REST_TOKEN });
    return (await redis.get(KV_KEY)) || [];
  }
  if (!fs.existsSync(DATA_FILE)) return [];
  return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
}

async function writePatients(data) {
  if (USE_KV) {
    const { Redis } = await import('@upstash/redis');
    const redis = new Redis({ url: process.env.UPSTASH_REDIS_REST_URL, token: process.env.UPSTASH_REDIS_REST_TOKEN });
    await redis.set(KV_KEY, data);
    return;
  }
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));
}

app.get('/api/hospitals', (_req, res) => res.json(HOSPITALS));

app.get('/api/patients', async (_req, res) => {
  try   { res.json(await readPatients()); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/patients', async (req, res) => {
  try {
    const { name, phone, emergencyContact, lvadDevice, lvadSettings, hospital, address, lat, lng, notes } = req.body;
    if (!name || !lat || !lng || !hospital)
      return res.status(400).json({ error: 'Missing required fields: name, lat, lng, hospital' });
    const patients = await readPatients();
    const patient = {
      id: crypto.randomUUID(),
      name, phone: phone || '', emergencyContact: emergencyContact || '',
      lvadDevice: lvadDevice || '', lvadSettings: lvadSettings || '',
      hospital, address: address || '',
      lat: Number(lat), lng: Number(lng),
      notes: notes || '',
      addedAt: new Date().toISOString()
    };
    patients.push(patient);
    await writePatients(patients);
    res.json({ success: true, patient });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/patients/:id', async (req, res) => {
  try {
    const patients = await readPatients();
    const idx = patients.findIndex(p => p.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'Patient not found' });
    patients[idx] = { ...patients[idx], ...req.body, id: patients[idx].id, addedAt: patients[idx].addedAt };
    await writePatients(patients);
    res.json({ success: true, patient: patients[idx] });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/patients/:id', async (req, res) => {
  try {
    const patients = await readPatients();
    const idx = patients.findIndex(p => p.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'Patient not found' });
    patients.splice(idx, 1);
    await writePatients(patients);
    res.json({ success: true, total: patients.length });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.listen(PORT, () => {
  console.log(`\n  LVAD Patient Monitor  →  http://localhost:${PORT}`);
  console.log(`  Admin                 →  http://localhost:${PORT}/admin.html\n`);
  console.log(`  Storage: ${USE_KV ? 'Upstash Redis' : 'Local file (lvad_patients.json)'}\n`);
  if (!USE_KV) {
    console.log('  NOTE: This application handles sensitive patient data (PHI).');
    console.log('  Ensure it is only accessible on an internal/VPN-protected network.\n');
  }
});
