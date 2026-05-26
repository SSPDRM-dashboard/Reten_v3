import fetch from 'node-fetch';
import Papa from 'papaparse';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";
const years = [2021, 2022, 2023, 2024, 2025, 2026];

async function run() {
  for (const y of years) {
    const url = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=${y}`;
    const res = await fetch(url);
    const text = await res.text();
    
    Papa.parse(text, {
      header: false,
      skipEmptyLines: true,
      complete: (parseResult) => {
        const rows = parseResult.data;
        if (rows.length > 0) {
          const row = rows[0];
          const dateIdx = row.findIndex(c => String(c).toUpperCase().includes('TARIKH'));
          const districtIdx = row.findIndex(c => String(c).toUpperCase().includes('DAERAH'));
          const taskIdx = row.findIndex(c => String(c).toUpperCase().includes('JENIS TUGASAN'));
          const hoursIdx = row.findIndex(c => String(c).toUpperCase().includes('JUMLAH JAM'));
          const rankIdx = row.findIndex(c => String(c).toUpperCase().includes('PANGKAT'));
          
          console.log(`Year ${y}:`);
          console.log(`  TARIKH (date) index: ${dateIdx}`);
          console.log(`  DAERAH (district) index: ${districtIdx}`);
          console.log(`  JENIS TUGASAN (task) index: ${taskIdx}`);
          console.log(`  JUMLAH JAM (hours) index: ${hoursIdx}`);
          console.log(`  PANGKAT (rank) index: ${rankIdx}`);
          console.log(`  Total row columns: ${row.length}`);
        }
      }
    });
  }
}

run();
