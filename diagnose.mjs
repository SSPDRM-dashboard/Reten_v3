import fetch from 'node-fetch';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";
const years = [2021, 2022, 2023, 2024, 2025, 2026];

async function check() {
  for (const y of years) {
    const url = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=${y}`;
    try {
      const res = await fetch(url);
      console.log(`Year ${y}: Status = ${res.status}`);
      if (res.ok) {
        const text = await res.text();
        console.log(`Year ${y}: First 200 chars = ${text.substring(0, 200).replace(/\r?\n/g, ' ')}`);
        console.log(`Year ${y}: Total length = ${text.length} chars`);
      } else {
        console.log(`Year ${y}: Error reading`);
      }
    } catch (err) {
      console.log(`Year ${y}: Catch error = ${err.message}`);
    }
  }
}

check();
