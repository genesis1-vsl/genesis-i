const express = require('express');
const fs = require('fs');
const path = require('path');
const app = express();

app.use(express.json({ limit: '20mb' }));
app.use(express.static(__dirname));

const DATA_FILE = path.join(__dirname, 'data.json');

app.get('/api/data', (req, res) => {
  try {
    if (fs.existsSync(DATA_FILE)) {
      res.json(JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')));
    } else {
      res.json(null);
    }
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/data', (req, res) => {
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(req.body));
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Genesis I running on port ${PORT}`));
