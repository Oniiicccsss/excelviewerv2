const viewport = document.getElementById('viewport');
const queryInput = document.getElementById('query');
const clearBtn = document.getElementById('clear');
const searchBox = document.getElementById('searchBox');
const loader = document.getElementById('loader');
const sheetTabsContainer = document.getElementById('sheetTabs');
let currentFile = null;

// Debounce utility to prevent search from running on every keystroke
function debounce(func, delay) {
  let timeoutId;
  return function (...args) {
    clearTimeout(timeoutId);
    timeoutId = setTimeout(() => func.apply(this, args), delay);
  };
}

async function loadFile(file, sheetName = null) {
  currentFile = file;
  const formData = new FormData();
  formData.append('file', file);
  if (sheetName) {
    formData.append('sheetName', sheetName);
  }
  loader.style.display = 'flex';

  try {
    // --- UPDATED: Target the new PHP backend file ---
    let res = await fetch('process_excel.php', { method: 'POST', body: formData });
    let data = await res.json();

    if (data.error) {
      viewport.innerHTML = '<p style="color:red">' + data.error + '</p>';
    } else {
      viewport.innerHTML = data.html;
      document.getElementById('colLabels').innerHTML = data.colLabels;
      document.getElementById('rowLabels').innerHTML = data.rowLabels;
      document.getElementById('sheetInfo').textContent = data.info;
      searchBox.style.display = 'flex';

      if (!sheetName) {
        sheetTabsContainer.innerHTML = '';
        data.sheets.forEach(sheet => {
          const tab = document.createElement('li');
          tab.className = 'sheet-tab';
          if (sheet === data.currentSheet) {
            tab.classList.add('active');
          }
          tab.textContent = sheet;
          tab.addEventListener('click', () => {
            document.querySelectorAll('.sheet-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            loadFile(currentFile, sheet);
          });
          sheetTabsContainer.appendChild(tab);
        });
      }
      adjustRowHeights();
    }
  } catch (err) {
    viewport.innerHTML = '<p style="color:red">Error loading file</p>';
  } finally {
    loader.style.display = 'none';
  }
}

function adjustRowHeights() {
  const cells = document.querySelectorAll('.cell.word-wrap, .cell');
  const rowsToUpdate = {};

  cells.forEach(cell => {
    const row = cell.getAttribute('data-row');
    const content = cell.querySelector('.cell-content');
    if (content) {
      const requiredHeight = content.scrollHeight;
      if (requiredHeight > (rowsToUpdate[row] || 0)) {
        rowsToUpdate[row] = requiredHeight;
      }
    }
  });

  for (const row in rowsToUpdate) {
    const newHeight = rowsToUpdate[row] + 16; // padding adjustment
    document.querySelectorAll(`.cell[data-row="${row}"]`).forEach(cell => {
      cell.style.height = `${newHeight}px`;
    });
    const rowLabel = document.querySelector(`.row-label[data-row="${row}"]`);
    if (rowLabel) {
      rowLabel.style.height = `${newHeight}px`; // sync height
      rowLabel.style.lineHeight = `${newHeight}px`; // vertical center
    }
  }
}

document.getElementById('file').addEventListener('change', function (e) {
  const file = e.target.files[0];
  if (!file) return;
  loadFile(file);
});

const performSearch = debounce(function () {
  const query = queryInput.value.trim().toLowerCase();
  const cells = viewport.querySelectorAll('.cell');
  let firstMatch = null;

  cells.forEach(cell => {
    let text = cell.textContent.toLowerCase();
    if (query && text.includes(query)) {
      if (!firstMatch) firstMatch = cell;
      const idx = text.indexOf(query);
      const original = cell.textContent;
      cell.innerHTML = '<span class="cell-content">' + original.substring(0, idx) +
        '<mark>' + original.substring(idx, idx + query.length) + '</mark>' +
        original.substring(idx + query.length) + '</span>';
    } else {
      cell.innerHTML = '<span class="cell-content">' + cell.textContent + '</span>';
    }
  });

  if (firstMatch) {
    firstMatch.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }
}, 300);

queryInput.addEventListener('input', performSearch);

clearBtn.addEventListener('click', () => {
  queryInput.value = '';
  performSearch();
});

const rowLabelsContainer = document.querySelector('.row-labels-container');
viewport.addEventListener('scroll', () => {
  rowLabelsContainer.scrollTop = viewport.scrollTop;
});