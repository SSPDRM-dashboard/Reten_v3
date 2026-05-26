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
    
    const yearCounts = {};
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
          if (y) {
            yearCounts[y] = (yearCounts[y] || 0) + 1;
          }
        }
      }
    }
    console.log(`Gviz Tab "${name}":`);
    console.log(`  Total rows: ${rows.length}`);
    console.log(`  Row counts per year:`, JSON.stringify(yearCounts));
  }
}

run();
