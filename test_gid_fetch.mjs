import fetch from 'node-fetch';
import Papa from 'papaparse';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";

async function run() {
  // Let's try gid=9 (which was sheetId 9 for "2025")
  const url = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&gid=9`;
  const res = await fetch(url);
  const text = await res.text();
  
  const parsed = Papa.parse(text).data;
  console.log(`Fetched gid=9 (Year 2025 tab):`);
  console.log(`  Total rows: ${parsed.length}`);
  if (parsed.length > 0) {
    console.log(`  Headers:`, JSON.stringify(parsed[0].slice(0, 10)));
    console.log(`  Row 1:`, JSON.stringify(parsed[1].slice(0, 10)));
  }
}

run();
