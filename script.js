let rubricsData = [];
let selectedMap = new Map();

async function loadExcel() {
  try {
    const response = await fetch('./rubrics.xlsx');
    const arrayBuffer = await response.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    rubricsData = XLSX.utils.sheet_to_json(worksheet);

    console.log('Данные загружены:', rubricsData);
  } catch (error) {
    console.error('Ошибка загрузки Excel:', error);
    document.getElementById('simpleResults').innerHTML = '<div class="empty">Ошибка загрузки Excel-файла</div>';
    document.getElementById('smartResults').innerHTML = '<div class="empty">Ошибка загрузки Excel-файла</div>';
  }
}

function normalizeText(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function getWords(text) {
  return normalizeText(text)
    .split(' ')
    .filter(Boolean);
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function highlightText(text, queryWords) {
  let result = String(text || '');

  queryWords
    .filter(word => word.length > 1)
    .sort((a, b) => b.length - a.length)
    .forEach(word => {
      const regex = new RegExp(`(${escapeRegExp(word)})`, 'gi');
      result = result.replace(regex, '<mark>$1</mark>');
    });

  return result;
}

function simpleSearch(query) {
  const queryWords = getWords(query);

  return rubricsData
    .map((row, index) => {
      const keywords = normalizeText(row['Ключевые слова']);
      const description = normalizeText(row['Описание']);

      let keywordMatches = 0;
      let descriptionMatches = 0;

      queryWords.forEach(word => {
        if (keywords.includes(word)) {
          keywordMatches += 1;
        }
        if (description.includes(word)) {
          descriptionMatches += 1;
        }
      });

      const totalMatches = keywordMatches + descriptionMatches;
      const weightedScore = keywordMatches * 2 + descriptionMatches;

      return {
        id: `simple-${index}`,
        originalId: index,
        rubric: row['Рубрика'] || '—',
        code: row['Код рубрики'] || '—',
        keywordsText: row['Ключевые слова'] || '',
        descriptionText: row['Описание'] || '',
        keywordMatches,
        descriptionMatches,
        totalMatches,
        score: weightedScore
      };
    })
    .filter(item => item.totalMatches > 0)
    .sort((a, b) => {
      if (b.totalMatches !== a.totalMatches) {
        return b.totalMatches - a.totalMatches;
      }
      if (b.keywordMatches !== a.keywordMatches) {
        return b.keywordMatches - a.keywordMatches;
      }
      if (b.descriptionMatches !== a.descriptionMatches) {
        return b.descriptionMatches - a.descriptionMatches;
      }
      return b.score - a.score;
    });
}

function smartSearch(query) {
  const queryWords = getWords(query);

  return rubricsData
    .map((row, index) => {
      const keywordsWords = getWords(row['Ключевые слова']);
      const descriptionWords = getWords(row['Описание']);

      let keywordScore = 0;
      let descriptionScore = 0;
      let partialScore = 0;
      let exactMatches = 0;

      queryWords.forEach(queryWord => {
        keywordsWords.forEach(word => {
          if (word === queryWord) {
            keywordScore += 3;
            exactMatches += 1;
          } else if (
            queryWord.length > 3 &&
            word.length > 3 &&
            (word.includes(queryWord) || queryWord.includes(word))
          ) {
            partialScore += 1.5;
          }
        });

        descriptionWords.forEach(word => {
          if (word === queryWord) {
            descriptionScore += 2;
            exactMatches += 1;
          } else if (
            queryWord.length > 3 &&
            word.length > 3 &&
            (word.includes(queryWord) || queryWord.includes(word))
          ) {
            partialScore += 1;
          }
        });
      });

      const totalScore = Number((keywordScore + descriptionScore + partialScore).toFixed(1));

      return {
        id: `smart-${index}`,
        originalId: index,
        rubric: row['Рубрика'] || '—',
        code: row['Код рубрики'] || '—',
        keywordsText: row['Ключевые слова'] || '',
        descriptionText: row['Описание'] || '',
        keywordScore,
        descriptionScore,
        partialScore,
        exactMatches,
        score: totalScore
      };
    })
    .filter(item => item.score > 0 && (item.exactMatches > 0 || item.score >= 2))
    .sort((a, b) => {
      if (b.score !== a.score) {
        return b.score - a.score;
      }
      return b.exactMatches - a.exactMatches;
    });
}

function isChecked(code, rubric) {
  const key = `${code}|||${rubric}`;
  return selectedMap.has(key);
}

function updateSelection(code, rubric, checked) {
  const key = `${code}|||${rubric}`;

  if (checked) {
    selectedMap.set(key, { code, rubric });
  } else {
    selectedMap.delete(key);
  }

  syncCheckboxes();
  renderSelectedPanel();
}

function syncCheckboxes() {
  const checkboxes = document.querySelectorAll('.result-checkbox');

  checkboxes.forEach(checkbox => {
    const code = checkbox.dataset.code;
    const rubric = checkbox.dataset.rubric;
    checkbox.checked = isChecked(code, rubric);
  });
}

function renderSelectedPanel() {
  const selectedCodesEl = document.getElementById('selectedCodes');
  const selectedItemsEl = document.getElementById('selectedItems');

  const items = Array.from(selectedMap.values());

  if (!items.length) {
    selectedCodesEl.textContent = 'Пока ничего не выбрано';
    selectedItemsEl.innerHTML = '';
    return;
  }

  const uniqueCodes = [...new Set(items.map(item => String(item.code)))];
  selectedCodesEl.textContent = uniqueCodes.join('\n');

  selectedItemsEl.innerHTML = items
    .map(item => `<li>${item.rubric} — <strong>${item.code}</strong></li>`)
    .join('');
}

function createResultCard(item, queryWords, mode) {
  const checked = isChecked(String(item.code), item.rubric);

  const scoreBlock = mode === 'simple'
    ? `
      <div class="meta"><strong>Всего совпадений:</strong> ${item.totalMatches}</div>
      <div class="meta"><strong>Совпадения в ключевых словах:</strong> ${item.keywordMatches}</div>
      <div class="meta"><strong>Совпадения в описании:</strong> ${item.descriptionMatches}</div>
    `
    : `
      <div class="meta"><strong>Общий балл:</strong> ${item.score}</div>
      <div class="meta"><strong>Точные совпадения:</strong> ${item.exactMatches}</div>
      <div class="meta"><strong>Совпадения в ключевых словах:</strong> ${item.keywordScore}</div>
      <div class="meta"><strong>Совпадения в описании:</strong> ${item.descriptionScore}</div>
      <div class="meta"><strong>Частичные совпадения:</strong> ${item.partialScore}</div>
    `;

  return `
    <div class="result-card">
      <div class="result-top">
        <input
          type="checkbox"
          class="result-checkbox"
          data-code="${String(item.code).replace(/"/g, '&quot;')}"
          data-rubric="${String(item.rubric).replace(/"/g, '&quot;')}"
          ${checked ? 'checked' : ''}
        />
        <div class="result-main">
          <div><strong>Рубрика:</strong> ${highlightText(item.rubric, queryWords)}</div>
          <div><strong>Код рубрики:</strong> ${item.code}</div>
          ${scoreBlock}
          <div class="meta"><strong>Ключевые слова:</strong> ${highlightText(item.keywordsText, queryWords)}</div>
          <div class="meta"><strong>Описание:</strong> ${highlightText(item.descriptionText, queryWords)}</div>
        </div>
      </div>
    </div>
  `;
}

function renderSimpleResults(results, query) {
  const container = document.getElementById('simpleResults');
  const queryWords = getWords(query);

  if (!results.length) {
    container.innerHTML = '<div class="empty">Нет совпадений</div>';
    return;
  }

  container.innerHTML = results
    .map(item => createResultCard(item, queryWords, 'simple'))
    .join('');

  attachCheckboxHandlers();
}

function renderSmartResults(results, query) {
  const container = document.getElementById('smartResults');
  const queryWords = getWords(query);

  if (!results.length) {
    container.innerHTML = '<div class="empty">Нет совпадений</div>';
    return;
  }

  container.innerHTML = results
    .map(item => createResultCard(item, queryWords, 'smart'))
    .join('');

  attachCheckboxHandlers();
}

function attachCheckboxHandlers() {
  const checkboxes = document.querySelectorAll('.result-checkbox');

  checkboxes.forEach(checkbox => {
    checkbox.removeEventListener('change', checkbox._handler);

    const handler = function () {
      const code = this.dataset.code;
      const rubric = this.dataset.rubric;
      updateSelection(code, rubric, this.checked);
    };

    checkbox._handler = handler;
    checkbox.addEventListener('change', handler);
  });
}

function runSearch() {
  const query = document.getElementById('query').value.trim();

  if (!query) {
    document.getElementById('simpleResults').innerHTML = '<div class="empty">Введите запрос</div>';
    document.getElementById('smartResults').innerHTML = '<div class="empty">Введите запрос</div>';
    return;
  }

  const simpleResults = simpleSearch(query);
  const smartResults = smartSearch(query);

  renderSimpleResults(simpleResults, query);
  renderSmartResults(smartResults, query);
  syncCheckboxes();
}

document.getElementById('searchBtn').addEventListener('click', runSearch);

document.getElementById('query').addEventListener('keydown', function(event) {
  if (event.key === 'Enter') {
    runSearch();
  }
});

document.getElementById('copyCodesBtn').addEventListener('click', async function () {
  const text = document.getElementById('selectedCodes').textContent;

  if (!text || text === 'Пока ничего не выбрано') {
    alert('Нет выбранных кодов для копирования');
    return;
  }

  try {
    await navigator.clipboard.writeText(text);
    alert('Коды скопированы');
  } catch (error) {
    alert('Не удалось скопировать коды');
  }
});

document.getElementById('clearSelectionBtn').addEventListener('click', function () {
  selectedMap.clear();
  syncCheckboxes();
  renderSelectedPanel();
});

loadExcel();
renderSelectedPanel();
