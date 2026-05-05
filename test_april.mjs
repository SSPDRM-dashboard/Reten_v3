import Papa from 'papaparse';

fetch('https://docs.google.com/spreadsheets/d/140GxAx-6bU_PQipsPY97Aj1lkETvy3OmUnvbSYbHe1k/export?format=csv')
  .then(res => res.text())
  .then(text => {
    Papa.parse(text, {
      header: true,
      complete: (results) => {
        results.data.forEach(r => {
           if(r['TARIKH MAKLUMAT']?.includes('04/2026')) {
               console.log(r['PEJAWATAN'], r['PEJAWATAN ANDA'], r['PANGKAT']);
           }
        })
      }
    });
  });
