document.addEventListener('DOMContentLoaded', () => {
  const table = document.getElementById('research-table');
  if (table) {
    fetch('assets/research%20papers.xlsx')
      .then(response => response.arrayBuffer())
      .then(data => {
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const html = XLSX.utils.sheet_to_html(workbook.Sheets[sheetName]);
        table.innerHTML = html;
      })
      .catch(error => console.error('Error loading research papers:', error));
  }
});
