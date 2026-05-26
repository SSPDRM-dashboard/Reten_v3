import fetch from 'node-fetch';
import Papa from 'papaparse';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";

async function run() {
  const url = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv`;
  const res = await fetch(url);
  const text = await res.text();
  
  Papa.parse(text, {
    header: false,
    skipEmptyLines: true,
    complete: (parseResult) => {
      const rows = parseResult.data;
      console.log(`Default sheet has ${rows.length} rows.`);
      
      const yearsCount = {};
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const dateStr = row[15]; // Column index 15 is TARIKH
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
            yearsCount[rowYear] = (yearsCount[rowYear] || 0) + 1;
          }
        }
      }
      console.log("Years count in default sheet:", yearsCount);
    }
  });
}

run();
