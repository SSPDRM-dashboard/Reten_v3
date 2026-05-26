import fetch from 'node-fetch';
import Papa from 'papaparse';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";
const sheetNames = [
  '22022', '22023', '22024', '22025',
  '2021', '2022', '2023', '2024', '2025', '2026'
];

async function run() {
  for (const name of sheetNames) {
    const url = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(name)}`;
    const res = await fetch(url);
    const text = await res.text();
    const rows = Papa.parse(text).data;
    
    // Sample years in this sheet
    const yearsSet = new Set();
    const dateIdx = 15; // TARIKH
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const dateStr = row[dateIdx];
      if (dateStr) {
        const dateOnly = String(dateStr).split(' ')[0];
        const parts = dateOnly.split(/[\/\-]/);
        if (parts.length === 3) {
          let y = parseInt(parts[2], 10);
          if (y < 100) y += 2000;
          if (y) yearsSet.add(y);
        }
      }
    }
    console.log(`Tab name "${name}" in Gviz: rows = ${rows.length}. Years found in TARIKH Col 15:`, Array.from(yearsSet));
  }
}

run();
