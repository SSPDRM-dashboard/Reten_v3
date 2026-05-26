import fetch from 'node-fetch';
import * as XLSX from 'xlsx';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";

async function run() {
  const url = `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`;
  console.log("Downloading spreadsheet XLSX...");
  const res = await fetch(url);
  const buffer = await res.arrayBuffer();
  
  console.log("Parsing XLSX...");
  const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
  
  console.log("Found Sheet Tab Names:");
  console.log(workbook.SheetNames);
}

run();
