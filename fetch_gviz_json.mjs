import fetch from 'node-fetch';

const id = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";

async function run() {
  const url = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:json`;
  const res = await fetch(url);
  const text = await res.text();
  
  // Clean up the google visualization JSON response
  const jsonStr = text.replace('/*O_o*/\ngoogle.visualization.Query.setResponse(', '').replace(/\);$/, '');
  try {
    const data = JSON.parse(jsonStr);
    console.log("JSON response signature successfully parsed. Keys:");
    console.log(Object.keys(data));
    console.log("Status:", data.status);
    if (data.table) {
      console.log("Table columns count:", data.table.cols ? data.table.cols.length : 0);
      console.log("Table rows count:", data.table.rows ? data.table.rows.length : 0);
    }
  } catch (err) {
    console.log("Failed to parse JSON, first 1000 chars of raw text:");
    console.log(text.substring(0, 1000));
  }
}

run();
