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
      const names = new Set();
      
      const nameIndices = [3, 5, 7, 9];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        nameIndices.forEach(idx => {
          if (idx < row.length && row[idx]) {
            names.add(String(row[idx]).trim().toUpperCase());
          }
        });
      }
      
      console.log("Distinct names/badan in 2025:");
      console.log(Array.from(names).slice(0, 20));
    }
  });
}

run();
