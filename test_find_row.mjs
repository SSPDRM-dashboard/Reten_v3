import Papa from 'papaparse';

fetch('https://docs.google.com/spreadsheets/d/140GxAx-6bU_PQipsPY97Aj1lkETvy3OmUnvbSYbHe1k/export?format=csv')
  .then(res => res.text())
  .then(text => {
    Papa.parse(text, {
      header: true,
      complete: (results) => {
        const row = results.data.find(r => (r['BUTIR-BUTIR MAKLUMAT'] || '').includes('minuman keras'));
        if (row) {
          console.log("Found row:");
          console.log("JENIS MAKLUMAT:", row['JENIS MAKLUMAT']);
          console.log("Raw row:", row);
        } else {
          console.log("Not found.");
        }
      }
    });
  });
