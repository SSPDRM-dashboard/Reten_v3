import fetch from 'node-fetch';
import Papa from 'papaparse';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";
const years = [2021, 2022, 2023, 2024, 2025, 2026];

async function check() {
  console.log("=== METHOD 1: gviz/tq (current code) ===");
  for (const y of years) {
    const url = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=${y}`;
    const res = await fetch(url);
    const text = await res.text();
    const rows = Papa.parse(text).data;
    console.log(`Year ${y}: cols = ${rows[0] ? rows[0].length : 0}, rows = ${rows.length}`);
  }
  
  console.log("\n=== METHOD 2: export?format=csv ===");
  for (const y of years) {
    // We encode the sheet name to ensure numeric values or special names are correctly treated as sheet titles
    const url = `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&sheet=${encodeURIComponent(y)}`;
    const res = await fetch(url);
    const text = await res.text();
    const rows = Papa.parse(text).data;
    console.log(`Year ${y}: cols = ${rows[0] ? rows[0].length : 0}, rows = ${rows.length}`);
  }
}

check();
