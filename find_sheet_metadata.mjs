import fetch from 'node-fetch';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";

async function run() {
  const url = `https://docs.google.com/spreadsheets/d/${id}/edit`;
  const res = await fetch(url);
  const html = await res.text();
  
  // Search for the initial data payload in standard Google Sheets HTML
  // Usually housed inside bootstrap_data or similar variable
  const regex = /gspro_meta\s*=\s*(.*?);/s;
  const match = html.match(regex);
  if (match) {
    console.log("gspro_meta matched!");
  } else {
    // Look for more generic spreadsheet titles / names
    // Tab names are often placed inside script tags, let's look for known tabs like "2026"
    console.log("Searching for strings '2026' or '2021' list in HTML...");
    // Let's find patterns like: ["2021", ... ] or similar
    const p1 = html.match(/\[\s*"\d\d\d\d"\s*,\s*"\d\d\d\d"\s*\]/g);
    console.log("Numeric list match:", p1);
    
    // Let's search inside the HTML for the names of all tabs by searching the JSON model
    const jsonMatch = html.match(/_bootstrapData\s*=\s*(.*?);\s*\n/);
    if (jsonMatch) {
      console.log("_bootstrapData found!");
      const rawJson = jsonMatch[1];
      // Search for tab names (title property)
      const matches = [...rawJson.matchAll(/"title"\s*:\s*"([^"]+)"/g)].map(m => m[1]);
      console.log("Tabs in bootstrap:", [...new Set(matches)]);
    } else {
      // Let's do a simple fuzzy match for "title":"AnyTabName"
      const fuzzyTitles = [...html.matchAll(/"title"\s*:\s*"([^"]+)"/g)].map(m => m[1]);
      const yearsSet = new Set(fuzzyTitles.filter(t => /^[12]\d\d\d$/.test(t) || t.includes("202")));
      console.log("Year-like titles found:", Array.from(yearsSet));
    }
  }
}

run();
