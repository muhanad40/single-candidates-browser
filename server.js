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

  // Find the header row (first row that contains "Full Name")
  let headerRowIdx = data.findIndex(row => row.includes('Full Name'));
  if (headerRowIdx === -1) throw new Error('Could not find header row in spreadsheet');

  // Build a column name → index map (normalise whitespace/case)
  const headers = data[headerRowIdx];
  const col = {};
  headers.forEach((h, i) => { if (h) col[h.toString().trim().toLowerCase()] = i; });

  const c = name => col[name] ?? -1;

  const candidates = [];
  let id = 3;

  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const row = data[i];
    const name = row[c('full name')];
    if (!name || typeof name !== 'string' || !name.trim()) continue;

    const rawAge = row[c('age')];
    const age = typeof rawAge === 'number' && rawAge > 0 && rawAge < 120 ? rawAge : null;
    const rawHeight = row[c('height (cm)')];
    const height = typeof rawHeight === 'number' && rawHeight > 0 ? rawHeight : null;
    const imageFile = rowToImage[i] || null;

    candidates.push({
      id: id++,
      image: imageFile ? `/images/${imageFile}` : null,
      name: name.trim(),
      dob:          row[c('date of birth')] || null,
      age,
      gender:       row[c('gender')] || null,
      height,
      education:    row[c('education')] || null,
      profession:   row[c('profession')] || null,
      income:       row[c('income status')] || null,
      religiosity:  row[c('religiousity level')] || row[c('religiosity level')] || null,
      location:     row[c('location')] || null,
      hobbies:      row[c('interests/ hobbies')] || row[c('interests/hobbies')] || null,
      ethnicity:    row[c('ethinicity')] || row[c('ethnicity')] || null,
      minAge:       typeof row[c('min age')] === 'number' ? row[c('min age')] : null,
      maxAge:       typeof row[c('max age')] === 'number' ? row[c('max age')] : null,
      status:       row[c('single/divorced?')] || null,
      hasKids:      row[c('do they have kids?')] || null,
      sect:         row[c('shia/ sunni')] || row[c('shia/sunni')] || null,
      contact:      row[c('contact details')] || null,
      matchHistory: row[c('match history')] || null,
      pointOfContact: row[c('point of contact')] || null,
      comments:     row[c('comments')] || null,
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
