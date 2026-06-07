const express = require('express');
const fs = require('fs');
const path = require('path');
const app = express();

app.use(express.json({ limit: '20mb' }));
app.use(express.static(__dirname));

/* ============================================================
   ARMAZENAMENTO
   - Se a variável de ambiente DATABASE_URL existir  -> PostgreSQL
     (durável e compartilhado entre TODOS os computadores).
   - Senão -> arquivo data.json em DATA_DIR (ou na pasta do app).
     Use DATA_DIR para apontar a um disco persistente (NAS/Render Disk).
   ============================================================ */
const DATA_DIR  = process.env.DATA_DIR || __dirname;
const DATA_FILE = path.join(DATA_DIR, 'data.json');

let pool = null;

async function initDb() {
  if (!process.env.DATABASE_URL) return;            // sem banco -> usa arquivo
  const { Pool } = require('pg');
  pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.PGSSL === 'disable' ? false : { rejectUnauthorized: false }
  });
  await pool.query('CREATE TABLE IF NOT EXISTS app_state (id INT PRIMARY KEY, data JSONB, updated_at TIMESTAMPTZ DEFAULT now())');
}

async function readData() {
  if (pool) {
    const r = await pool.query('SELECT data FROM app_state WHERE id = 1');
    return r.rows.length ? r.rows[0].data : null;
  }
  if (fs.existsSync(DATA_FILE)) return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
  return null;
}

async function writeData(obj) {
  if (pool) {
    await pool.query(
      `INSERT INTO app_state (id, data, updated_at) VALUES (1, $1::jsonb, now())
       ON CONFLICT (id) DO UPDATE SET data = $1::jsonb, updated_at = now()`,
      [JSON.stringify(obj)]
    );
    return;
  }
  fs.writeFileSync(DATA_FILE, JSON.stringify(obj));
}

app.get('/api/health', (req, res) =>
  res.json({ ok: true, storage: pool ? 'postgres' : 'file', version: 'v8.4' })
);

app.get('/api/data', async (req, res) => {
  try { res.json(await readData()); }
  catch (e) { console.error('GET /api/data', e); res.status(500).json({ error: e.message }); }
});

app.post('/api/data', async (req, res) => {
  try { await writeData(req.body); res.json({ ok: true }); }
  catch (e) { console.error('POST /api/data', e); res.status(500).json({ error: e.message }); }
});

app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

const PORT = process.env.PORT || 3000;
initDb()
  .catch(e => { console.error('Falha ao iniciar o banco, usando arquivo:', e.message); pool = null; })
  .finally(() => app.listen(PORT, () =>
    console.log(`Genesis I na porta ${PORT} — armazenamento: ${pool ? 'PostgreSQL' : ('arquivo (' + DATA_FILE + ')')}`)
  ));
