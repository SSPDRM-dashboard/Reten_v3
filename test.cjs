fetch('https://docs.google.com/spreadsheets/d/140GxAx-6bU_PQipsPY97Aj1lkETvy3OmUnvbSYbHe1k/export?format=csv')
  .then(res => res.text())
  .then(text => {
    const lines = text.split('\n');
    console.log(lines[0].split(','));
    console.log(lines[1].split(','));
  });
