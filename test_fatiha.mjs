import Papa from 'papaparse';

fetch('https://docs.google.com/spreadsheets/d/140GxAx-6bU_PQipsPY97Aj1lkETvy3OmUnvbSYbHe1k/export?format=csv')
  .then(res => res.text())
  .then(text => {
    Papa.parse(text, {
      header: true,
      complete: (results) => {
        const row = results.data.find(r => (r['NAMA'] || '').includes('NOR FATIHA'));
        if (row) {
          const j = String(row['JENIS MAKLUMAT'] || '').toUpperCase().trim();
          console.log("JENIS MAKLUMAT:", j);
          console.log("KATEGORI MAKLUMAT:", row['KATEGORI MAKLUMAT']);
        } else {
            console.log("Not found");
        }
      }
    });
  });
