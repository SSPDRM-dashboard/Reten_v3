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
      
      let parsedSuccessfully = 0;
      let failedToParse = 0;

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const dateStr = row[15]; // Column index 15 is TARIKH
        
        let rowMonth = -1, rowYear = -1, rowDay = -1;
        
        if (dateStr) {
          const dateOnly = String(dateStr).split(' ')[0];
          const parts = dateOnly.split(/[\/\-]/);
          
          if (parts.length === 3) {
            if (parts[0].length === 4) {
              rowYear = parseInt(parts[0], 10);
              rowMonth = parseInt(parts[1], 10) - 1;
              rowDay = parseInt(parts[2], 10);
            } else {
              const p0 = parseInt(parts[0], 10);
              const p1 = parseInt(parts[1], 10);
              const p2 = parseInt(parts[2], 10);
              
              if (p1 > 12) {
                rowMonth = p0 - 1;
                rowDay = p1;
                rowYear = p2;
              } else {
                rowDay = p0;
                rowMonth = p1 - 1;
                rowYear = p2;
              }
              if (rowYear < 100) rowYear += 2000;
            }
          } else {
            const d = new Date(dateStr);
            if (!isNaN(d.getTime())) {
              rowYear = d.getFullYear();
              rowMonth = d.getMonth();
              rowDay = d.getDate();
            }
          }
        }
        
        if (rowYear === 2025 && rowMonth >= 0 && rowMonth < 12) {
          parsedSuccessfully++;
          if (parsedSuccessfully <= 5) {
            console.log(`Row ${i} date = ${dateStr} => parsed to Year: ${rowYear}, Month: ${rowMonth}, Day: ${rowDay}`);
          }
        } else {
          failedToParse++;
          console.log(`Failed to parse row ${i} date = ${dateStr} => Year: ${rowYear}, Month: ${rowMonth}, Day: ${rowDay}`);
        }
      }
      
      console.log(`Results: ${parsedSuccessfully} parsed successfully, ${failedToParse} failed.`);
    }
  });
}

run();
