import fetch from 'node-fetch';
import * as XLSX from 'xlsx';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";

async function run() {
  const url = `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`;
  const res = await fetch(url);
  const buffer = await res.arrayBuffer();
  
  const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
  
  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    console.log(`Sheet "${name}":`);
    console.log(`  Total rows: ${data.length}`);
    if (data.length > 1) {
      // Find what year is in the rows
      const counts = {};
      const dateIdx = 15; // TARIKH
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const dateStr = row[dateIdx];
        if (dateStr) {
          const dateOnly = String(dateStr).split(' ')[0];
          const parts = dateOnly.split(/[\/\-]/);
          let rowYear = -1;
          if (parts.length === 3) {
            const p2 = parseInt(parts[2], 10);
            rowYear = p2;
            if (rowYear < 100) rowYear += 2000;
          }
          if (rowYear !== -1) {
            counts[rowYear] = (counts[rowYear] || 0) + 1;
          }
        }
      }
      console.log(`  Years in data:`, JSON.stringify(counts));
    }
  }
}

run();
