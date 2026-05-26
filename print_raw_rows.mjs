import fetch from 'node-fetch';
import Papa from 'papaparse';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";

async function run() {
  const url = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=2025`;
  const res = await fetch(url);
  const text = await res.text();
  
  Papa.parse(text, {
    header: false,
    skipEmptyLines: true,
    complete: (parseResult) => {
      const rows = parseResult.data;
      console.log(`Parsed ${rows.length} rows.`);
      
      for (let i = 0; i < Math.min(20, rows.length); i++) {
        const row = rows[i];
        const nonValueCols = [];
        row.forEach((val, idx) => {
          if (val !== undefined && val !== null && val !== "") {
            nonValueCols.push(`${idx}: ${val}`);
          }
        });
        console.log(`Row ${i}:`, nonValueCols.join(' | '));
      }
    }
  });
}

run();
