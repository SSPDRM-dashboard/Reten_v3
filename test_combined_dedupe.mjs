import fetch from 'node-fetch';
import Papa from 'papaparse';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";
const allTabsToFetch = [
  '2021', '2022', '2023', '2024', '2025', '2026',
  '22022', '22023', '22024', '22025'
];

async function run() {
  const fetchPromises = allTabsToFetch.map(async (name) => {
    const url = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(name)}`;
    try {
      const res = await fetch(url);
      if (res.ok) {
        const text = await res.text();
        const rows = Papa.parse(text).data;
        console.log(`Fetched tab "${name}": rows = ${rows.length}`);
        return { name, rows };
      }
    } catch (err) {
      console.warn(`Error on tab ${name}:`, err.message);
    }
    return null;
  });
  
  const results = (await Promise.all(fetchPromises)).filter(Boolean);
  
  // Combine and dedupe
  const uniqueRowsMap = new Map();
  let totalRawRows = 0;
  
  for (const item of results) {
    const { name, rows } = item;
    // skip header row
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      totalRawRows++;
      
      // Construct deduplication key
      // Use columns: Timestamp, Name (at 3, 5, 7, 9), Date, and Hours
      const timestamp = String(row[0]).trim();
      const dateVal = String(row[15]).trim();
      const hoursVal = String(row[18]).trim();
      let nameVal = '';
      [3, 5, 7, 9].forEach(idx => {
        if (idx < row.length && row[idx]) nameVal += String(row[idx]).trim();
      });
      
      const key = `${timestamp}|${nameVal}|${dateVal}|${hoursVal}`;
      
      if (!uniqueRowsMap.has(key)) {
        uniqueRowsMap.set(key, row);
      }
    }
  }
  
  console.log(`\nTotal raw rows processed across all tabs: ${totalRawRows}`);
  console.log(`Total unique rows after deduplication: ${uniqueRowsMap.size}`);
  
  // Count years of unique rows
  const yearCounts = {};
  for (const row of uniqueRowsMap.values()) {
    const dateStr = row[15]; // Column index 15 is TARIKH
    if (dateStr) {
      const dateOnly = String(dateStr).split(' ')[0];
      const parts = dateOnly.split(/[\/\-]/);
      let rowYear = -1;
      if (parts.length === 3) {
        let p2 = parseInt(parts[2], 10);
        if (p2 < 100) p2 += 2000;
        rowYear = p2;
      }
      if (rowYear !== -1) {
        yearCounts[rowYear] = (yearCounts[rowYear] || 0) + 1;
      }
    }
  }
  
  console.log(`\nYear distribution of deduplicated rows:`);
  console.log(JSON.stringify(yearCounts, null, 2));
}

run();
