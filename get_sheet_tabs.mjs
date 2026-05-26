import fetch from 'node-fetch';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";

async function run() {
  const url = `https://docs.google.com/spreadsheets/d/${id}/edit`;
  const res = await fetch(url);
  const html = await res.text();
  
  // Look for bootstrap data or tab names in the HTML page source
  const matches = html.matchAll(/"_gid":\s*"\d+",\s*"_name":\s*"([^"]+)"/g);
  const tabNamesSet = new Set();
  for (const match of matches) {
    tabNamesSet.add(match[1]);
  }
  
  console.log("Found raw tab names via _name pattern:");
  console.log(Array.from(tabNamesSet));
  
  // Try another common pattern in Google Sheets html
  const sheetNamesMatches = html.matchAll(/{"sheetId":\d+,"title":"([^"]+)"/g);
  const titleSet = new Set();
  for (const match of sheetNamesMatches) {
    titleSet.add(match[1]);
  }
  console.log("Found title tab names via title pattern:");
  console.log(Array.from(titleSet));
}

run();
