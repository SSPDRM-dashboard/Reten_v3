import fetch from 'node-fetch';
import Papa from 'papaparse';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";
const sheetNames = [
  '22022', '22023', '22024', '22025',
  '2021', '2022', '2023', '2024', '2025', '2026', '2027', '2028',
  'Users'
];

async function run() {
  for (const name of sheetNames) {
    const url = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(name)}`;
    const res = await fetch(url);
    const text = await res.text();
    const rows = Papa.parse(text).data;
    console.log(`Tab name "${name}": rows = ${rows.length}, cols = ${rows[0] ? rows[0].length : 0}`);
  }
}

run();
