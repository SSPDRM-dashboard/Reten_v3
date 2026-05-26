import fetch from 'node-fetch';
import * as XLSX from 'xlsx';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";

async function run() {
  const url = `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`;
  const res = await fetch(url);
  const buffer = await res.arrayBuffer();
  
  const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
  const sheet = workbook.Sheets["2025"];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  
  console.log(`Sheet 2025 has ${rows.length} rows.`);
  console.log("Header columns in XLSX sheet 2025:");
  console.log(JSON.stringify(rows[0]));
  
  // Let's inspect some row dates
  console.log("\nRow index and date values in column 15:");
  let count = 0;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const dateVal = row[15];
    if (dateVal !== undefined && dateVal !== null && dateVal !== "") {
      count++;
      if (count <= 15) {
        console.log(`Row ${i}: Original Col 15 = "${dateVal}" (Type: ${typeof dateVal})`);
      }
    }
  }
  console.log(`Total non-empty dates in Col 15: ${count}`);
}

run();
