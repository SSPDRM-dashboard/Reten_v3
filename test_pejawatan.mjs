import Papa from 'papaparse';

fetch('https://docs.google.com/spreadsheets/d/140GxAx-6bU_PQipsPY97Aj1lkETvy3OmUnvbSYbHe1k/export?format=csv')
  .then(res => res.text())
  .then(text => {
    Papa.parse(text, {
      header: true,
      complete: (results) => {
        const unique = new Set(results.data.map(r => r['PEJAWATAN']));
        console.log("PEJAWATAN:", unique);
        const uniquePenyalur = new Set(results.data.map(r => r['PEJAWATAN ANDA']));
        console.log("PEJAWATAN ANDA:", uniquePenyalur);
      }
    });
  });
