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
  
  // XLSX library stores sheet metadata including ID in some internal structures, let's look at workbook.Workbook
  if (workbook.Workbook && workbook.Workbook.Sheets) {
    console.log("Sheet IDs in workbook.Workbook.Sheets:");
    console.log(JSON.stringify(workbook.Workbook.Sheets, null, 2));
  } else {
    console.log("workbook.Workbook.Sheets not found.");
  }
}

run();
