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
      const counts = {};
      
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const dist = row[1];
        counts[dist] = (counts[dist] || 0) + 1;
      }
      
      console.log("District rows count in year 2025 sheet:");
      console.log(JSON.stringify(counts, null, 2));
    }
  });
}

run();
