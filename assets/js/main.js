document.addEventListener('DOMContentLoaded', () => {
  loadSheet('research-table', 'assets/research%20papers.xlsx');
  loadSheet('projects-table', 'assets/projects.xlsx');
  loadSheet('teaching-exp-table', 'assets/teaching_exp.xlsx');
});

function loadSheet(elementId, url) {
  const table = document.getElementById(elementId);
  if (!table) return;
  fetch(url)
    .then(res => res.arrayBuffer())
    .then(data => {
      const wb = XLSX.read(data, { type: 'array' });
      const sheet = wb.SheetNames[0];
      table.innerHTML = XLSX.utils.sheet_to_html(wb.Sheets[sheet]);
    })
    .catch(err => console.error(`Error loading ${url}:`, err));
}

