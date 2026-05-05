const Papa = require('papaparse');
const fs = require('fs');

const fetch = require('node-fetch');

fetch('https://docs.google.com/spreadsheets/d/140GxAx-6bU_PQipsPY97Aj1lkETvy3OmUnvbSYbHe1k/export?format=csv')
  .then(res => res.text())
  .then(text => {
    Papa.parse(text, {
      header: true,
      complete: (results) => {
        console.log("Headers:");
        console.log(results.meta.fields);
        console.log("First row:");
        console.log(results.data[0]);
      }
    });
  });
