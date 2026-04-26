const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.static('public'));

function buildRowImageMap() {
  const relsPath = path.join(__dirname, 'xlsx_extracted/contents/xl/drawings/_rels/drawing1.xml.rels');
  const drawingPath = path.join(__dirname, 'xlsx_extracted/contents/xl/drawings/drawing1.xml');
  if (!fs.existsSync(relsPath) || !fs.existsSync(drawingPath)) return {};

  const relsXml = fs.readFileSync(relsPath, 'utf8');
  const relsMap = {};
  for (const m of relsXml.matchAll(/Id="(rId\d+)"[^>]+Target="\.\.\/media\/([^"]+)"/g)) {
    relsMap[m[1]] = m[2];
  }

  const drawingXml = fs.readFileSync(drawingPath, 'utf8');
  const rowToImage = {};
  for (const anchor of drawingXml.split('<xdr:oneCellAnchor>').slice(1)) {
    const rowM = anchor.match(/<xdr:row>(\d+)<\/xdr:row>/);
    const ridM = anchor.match(/r:embed="(rId\d+)"/);
    if (rowM && ridM && relsMap[ridM[1]]) {
      rowToImage[parseInt(rowM[1])] = relsMap[ridM[1]];
    }
  }
  return rowToImage;
}

function loadCandidates() {
  const wb = XLSX.readFile(path.join(__dirname, 'candidates.xlsx'));
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const rowToImage = buildRowImageMap();

  const candidates = [];
  let id = 3;

  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    const name = row[1];
    if (!name || typeof name !== 'string') continue;

    const age = typeof row[4] === 'number' && row[4] < 120 ? row[4] : null;
    const height = typeof row[6] === 'number' ? row[6] : null;
    // XML row index = i (0-indexed from top of sheet, row 0 = sheet row 1)
    const xmlRow = i; // data[0]=row1, data[2]=row3=first candidate, xmlRow=2
    const imageFile = rowToImage[xmlRow] || null;

    candidates.push({
      id: id++,
      image: imageFile ? `/images/${imageFile}` : null,
      name: name.trim(),
      dob: row[3] || null,
      age,
      gender: row[5] || null,
      height,
      education: row[7] || null,
      profession: row[8] || null,
      income: row[9] || null,
      religiosity: row[10] || null,
      location: row[11] || null,
      hobbies: row[12] || null,
      ethnicity: row[13] || null,
      minAge: typeof row[14] === 'number' ? row[14] : null,
      maxAge: typeof row[15] === 'number' ? row[15] : null,
      status: row[16] || null,
      hasKids: row[17] || null,
      sect: row[18] || null,
      contact: row[19] || null,
      matchHistory: row[20] || null,
      comments: row[21] || null,
    });
  }

  return candidates;
}

app.get('/api/matched-pairs', (req, res) => {
  delete require.cache[require.resolve('./matchedPairs')];
  const pairs = require('./matchedPairs');
  res.json({ pairs });
});

app.get('/api/candidates', (req, res) => {
  try {
    res.json(loadCandidates());
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

const PORT = 3000;
app.listen(PORT, () => console.log(`Candidate browser running at http://localhost:${PORT}`));
