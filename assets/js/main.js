document.addEventListener('DOMContentLoaded', () => {
  loadSheet('research-table', 'assets/research%20papers.xlsx');
  loadSheet('projects-table', 'assets/projects.xlsx');
  loadSheet('teaching-exp-table', 'assets/teaching_exp.xlsx');
  loadSheet('tamil-articles-table', 'assets/tamil_articles.xlsx');
  loadSheet('tamil-articles-table', 'assets/awards.xlsx');
});

function loadSheet(elementId, url) {
  const container = document.getElementById(elementId);
  if (!container) return;
  fetch(url)
    .then(res => res.arrayBuffer())
    .then(data => {
      const wb = XLSX.read(data, { type: 'array' });
      const sheet = wb.SheetNames[0];
      const worksheet = wb.Sheets[sheet];
      if (!worksheet || !worksheet['!ref']) {
        container.innerHTML = '<div class="sheet"></div>';
        return;
      }

      const range = XLSX.utils.decode_range(worksheet['!ref']);
      const rows = [];

      for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum += 1) {
        const cells = [];

        for (let colNum = range.s.c; colNum <= range.e.c; colNum += 1) {
          const address = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
          const cell = worksheet[address];
          let cellHtml = '';

          if (cell) {
            if (cell.l && cell.l.Target) {
              const linkText = cell.v || cell.l.Target;
              cellHtml = `<a href="${cell.l.Target}" target="_blank" rel="noopener">${linkText}</a>`;
            } else if (typeof cell.f === 'string' && cell.f.startsWith('HYPERLINK(')) {
              const hyperlinkMatch = cell.f.match(/^HYPERLINK\(\s*(["'])(.*?)\1\s*,\s*(["'])(.*?)\3\s*\)$/i);
              if (hyperlinkMatch) {
                const url = hyperlinkMatch[2].replace(/""/g, '"');
                const label = hyperlinkMatch[4].replace(/""/g, '"');
                cellHtml = `<a href="${url}" target="_blank" rel="noopener">${label}</a>`;
              } else if (cell.v != null) {
                cellHtml = cell.v;
              }
            } else if (cell.w != null) {
              cellHtml = cell.w;
            } else if (cell.v != null) {
              cellHtml = cell.v;
            }
          }

          cells.push(`<div class="sheet-cell">${cellHtml}</div>`);
        }

        rows.push(`<div class="sheet-row">${cells.join('')}</div>`);
      }

      container.innerHTML = `<div class="sheet">${rows.join('')}</div>`;
    })
    .catch(err => console.error(`Error loading ${url}:`, err));
}

