import Papa from 'papaparse';

fetch('https://docs.google.com/spreadsheets/d/140GxAx-6bU_PQipsPY97Aj1lkETvy3OmUnvbSYbHe1k/export?format=csv')
  .then(res => res.text())
  .then(text => {
    Papa.parse(text, {
      header: true,
      complete: (results) => {
        const row = results.data.find(r => (r['BUTIR-BUTIR MAKLUMAT'] || '').includes('minuman keras'));
        if (row) {
          const j = String(row['JENIS MAKLUMAT'] || '').toUpperCase().trim();
          console.log("JENIS MAKLUMAT:", j);
          console.log("Matches?", j.includes('JENAYAH'));
          let hex = '';
          for(let i=0; i<j.length; i++) { hex += j.charCodeAt(i).toString(16) + ' '; }
          console.log("hex:", hex);
        }
      }
    });
  });
