let rubricsData = [];

async function loadExcel() {
  const response = await fetch('./Рубрики_Россия_07.05.26_с_описанием.xlsx');
  const arrayBuffer = await response.arrayBuffer();

  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];

  rubricsData = XLSX.utils.sheet_to_json(worksheet);

  console.log('Данные загружены:', rubricsData);
}

function normalizeText(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function simpleSearch(query) {
  const normalizedQuery = normalizeText(query);
  const queryWords = normalizedQuery.split(' ').filter(Boolean);

  return rubricsData
    .map(row => {
      const description = normalizeText(row['Описание']);
      let score = 0;

      queryWords.forEach(word => {
        if (description.includes(word)) {
          score++;
        }
      });

      return {
        rubric: row['Рубрика'],
        code: row['Код рубрики'],
        score
      };
    })
    .filter(item => item.score > 0)
    .sort((a, b) => b.score - a.score);
}

function renderSimpleResults(results) {
  const container = document.getElementById('simpleResults');

  if (!results.length) {
    container.innerHTML = '<p>Ничего не найдено</p>';
    return;
  }

  container.innerHTML = results
    .map(item => `
      <div style="margin-bottom: 10px; padding: 10px; border: 1px solid #ccc;">
        <div><strong>Рубрика:</strong> ${item.rubric}</div>
        <div><strong>Код рубрики:</strong> ${item.code}</div>
        <div><strong>Совпадений:</strong> ${item.score}</div>
      </div>
    `)
    .join('');
}

document.getElementById('searchBtn').addEventListener('click', () => {
  const query = document.getElementById('query').value;

  const simpleResults = simpleSearch(query);
  renderSimpleResults(simpleResults);

  document.getElementById('smartResults').innerHTML = '<p>Пока не реализован</p>';
});

loadExcel();