const express = require('express');
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
        BorderStyle, ShadingType, AlignmentType, VerticalAlign, HeightRule, PageBreak,
        Header, ImageRun, PageOrientation } = require('docx');
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

/* ============================================================
   GANGWAY VISITORS RECORD BOOK (.docx) — gerado a partir do
   template OSM Thome (A4 paisagem, cabeçalho azul repetido,
   10 nomes por página, "Mandatory Check" a cada 25, Date/Time On
   em branco, mínimo de 10 páginas).
   ============================================================ */
const OSM_LOGO_B64 = "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";
const logoBuf = Buffer.from(OSM_LOGO_B64, 'base64');
const COLW = [2971,2412,991,1560,1560,2976,1560,1699];
const TABLEW = COLW.reduce((a,b)=>a+b,0);
const HDR = ["Visitor\u2019s Name","Visitor\u2019s Company buisness onboard","Visitor`s Pass No.","Date / Time On","Personal belongings checked Yes/No","Visitor\u2019s Signature","Date / Time Off","Duty Crew Signature"];
const BRD = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const CB  = { top:BRD, bottom:BRD, left:BRD, right:BRD };
const NB  = { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} };
const grun = (t,o={}) => new TextRun({ text:t, font:o.font||"Times New Roman", size:o.size||24, bold:!!o.bold });
const gpar = (t,o={}) => new Paragraph({ alignment:o.align||AlignmentType.LEFT, spacing:{after:0}, children: t===""?[]:[grun(t,o)] });

function gHeaderRow() {
  return new TableRow({ tableHeader:true, height:{value:1319,rule:HeightRule.ATLEAST},
    children: COLW.map((w,i)=> new TableCell({ width:{size:w,type:WidthType.DXA}, borders:CB, verticalAlign:VerticalAlign.CENTER,
      shading:{type:ShadingType.CLEAR,fill:"8EAADB",color:"auto"}, margins:{top:20,bottom:20,left:60,right:60},
      children:[ new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:0}, children:[grun(HDR[i],{font:"Arial",size:20,bold:true})] }) ] })) });
}
function gDataRow(person,num) {
  const name = person ? (person.name||"") : "";
  const company = person ? (person.company||"") : "";
  const mand = (num % 25 === 0) ? "Mandatory Check" : "";
  const C = (i,children,valign) => new TableCell({ width:{size:COLW[i],type:WidthType.DXA}, borders:CB,
    verticalAlign:valign||VerticalAlign.CENTER, margins:{top:30,bottom:30,left:60,right:60}, children });
  return new TableRow({ height:{value:671,rule:HeightRule.ATLEAST}, children:[
    C(0,[gpar(name)],VerticalAlign.TOP),
    C(1,[gpar(company,{align:AlignmentType.CENTER})]),
    C(2,[gpar("")]),
    C(3,[gpar("____________",{size:20}), gpar("Time:",{size:20})]),
    C(4,[ mand ? new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:0},children:[grun(mand,{font:"Arial",size:20,bold:true})]}) : gpar("") ]),
    C(5,[gpar("")]),
    C(6,[gpar("")]),
    C(7,[gpar("")]),
  ]});
}
function gTable(people,pageIndex) {
  const PER=10, rows=[gHeaderRow()];
  for (let r=0;r<PER;r++){ const idx=pageIndex*PER+r; rows.push(gDataRow(people[idx],idx+1)); }
  return new Table({ width:{size:TABLEW,type:WidthType.DXA}, columnWidths:COLW.slice(), rows });
}
function gSectionHeader() {
  return new Header({ children:[
    new Table({ width:{size:TABLEW,type:WidthType.DXA}, columnWidths:[3200,TABLEW-3200], borders:NB, rows:[
      new TableRow({ children:[
        new TableCell({ width:{size:3200,type:WidthType.DXA}, borders:NB, verticalAlign:VerticalAlign.CENTER,
          children:[ new Paragraph({ spacing:{after:0}, children:[ new ImageRun({ type:"jpg", data:logoBuf, transformation:{width:165,height:40} }) ] }) ] }),
        new TableCell({ width:{size:TABLEW-3200,type:WidthType.DXA}, borders:NB, verticalAlign:VerticalAlign.CENTER, children:[
          new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:0}, children:[grun("Gangway Visitors Record Book",{font:"Arial",size:34,bold:true})] }),
          new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:0}, children:[grun("Marsec level 1",{font:"Arial",size:24})] }),
        ]}),
      ]})
    ]})
  ]});
}
function buildGangwayDocx(people) {
  people = Array.isArray(people) ? people : [];
  const PER=10, nPages=Math.max(10, Math.ceil(people.length/PER));
  const body=[];
  for (let p=0;p<nPages;p++){
    if (p>0) body.push(new Paragraph({ spacing:{after:0}, children:[new PageBreak()] }));
    body.push(gTable(people,p));
  }
  return new Document({
    sections:[{
      properties:{ page:{ size:{ width:11906, height:16838, orientation:PageOrientation.LANDSCAPE },
        margin:{ top:1380, right:283, bottom:280, left:708, header:242, footer:0 } } },
      headers:{ default:gSectionHeader() },
      children:body
    }]
  });
}

app.get('/api/health', (req, res) =>
  res.json({ ok: true, storage: pool ? 'postgres' : 'file', version: 'v8.5' })
);

app.get('/api/data', async (req, res) => {
  try { res.json(await readData()); }
  catch (e) { console.error('GET /api/data', e); res.status(500).json({ error: e.message }); }
});

app.post('/api/data', async (req, res) => {
  try { await writeData(req.body); res.json({ ok: true }); }
  catch (e) { console.error('POST /api/data', e); res.status(500).json({ error: e.message }); }
});

app.post('/api/gangway', async (req, res) => {
  try {
    const people = (req.body && Array.isArray(req.body.people)) ? req.body.people : [];
    const buffer = await Packer.toBuffer(buildGangwayDocx(people));
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="Gangway_Visitors_Genesis.docx"');
    res.send(buffer);
  } catch (e) { console.error('POST /api/gangway', e); res.status(500).json({ error: e.message }); }
});

app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

const PORT = process.env.PORT || 3000;
initDb()
  .catch(e => { console.error('Falha ao iniciar o banco, usando arquivo:', e.message); pool = null; })
  .finally(() => app.listen(PORT, () =>
    console.log(`Genesis I na porta ${PORT} — armazenamento: ${pool ? 'PostgreSQL' : ('arquivo (' + DATA_FILE + ')')}`)
  ));
