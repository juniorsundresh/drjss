document.addEventListener('DOMContentLoaded', () => {
  loadSheet('research-table', 'assets/research%20papers.xlsx');
  loadSheet('projects-table', 'assets/projects.xlsx');
  loadSheet('teaching-exp-table', 'assets/teaching_exp.xlsx');
  loadSheet('tamil-articles-table', 'assets/tamil_articles.xlsx');
});

function loadSheet(elementId, url) {
  const container = document.getElementById(elementId);
  if (!container) return;
  fetch(url)
    .then(res => res.arrayBuffer())
    .then(data => {
      const wb = XLSX.read(data, { type: 'array' });
      const sheet = wb.SheetNames[0];
      const json = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1 });
      const maxCols = Math.max(0, ...json.map(row => row.length));
      const html = json
        .map(row =>
          `<div class="sheet-row">${Array.from({ length: maxCols }, (_, i) =>
            `<div class="sheet-cell">${row[i] !== undefined ? row[i] : ''}</div>`
          ).join('')}</div>`
        )
        .join('');
      container.innerHTML = `<div class="sheet">${html}</div>`;
    })
    .catch(err => console.error(`Error loading ${url}:`, err));
}

