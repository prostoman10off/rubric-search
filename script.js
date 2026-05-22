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

function escapeHtml(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getMatchData(queryWords, textWords) {
  const exactMatches = new Set();
  const partialMatches = new Set();

  queryWords.forEach(queryWord => {
    textWords.forEach(word => {
      if (word === queryWord) {
        exactMatches.add(word);
      } else if (
        queryWord.length >= 4 &&
        word.length >= 4 &&
        (word.includes(queryWord) || queryWord.includes(word))
      ) {
        if (word.includes(queryWord)) {
          partialMatches.add(queryWord);
        } else if (queryWord.includes(word)) {
          partialMatches.add(word);
        }
      }
    });
  });

  return {
    exactMatches: Array.from(exactMatches),
    partialMatches: Array.from(partialMatches)
  };
}

function highlightTextByMatchType(text, exactFragments = [], partialFragments = []) {
  let result = escapeHtml(String(text || ''));

  const allFragments = [
    ...exactFragments.map(value => ({ value, type: 'exact' })),
    ...partialFragments.map(value => ({ value, type: 'partial' }))
  ]
    .filter(item => item.value && item.value.length > 1)
    .sort((a, b) => b.value.length - a.value.length);

  allFragments.forEach(fragment => {
    const className = fragment.type === 'exact' ? 'hl-exact' : 'hl-partial';
    const regex = new RegExp(`(${escapeRegExp(fragment.value)})`, 'gi');
    result = result.replace(regex, `<span class="${className}">$1</span>`);
  });

  return result;
}

function countExactMatches(queryWords, textWords) {
  let count = 0;
  const usedIndexes = new Set();

  queryWords.forEach(queryWord => {
    for (let i = 0; i < textWords.length; i++) {
      if (!usedIndexes.has(i) && textWords[i] === queryWord) {
        count++;
        usedIndexes.add(i);
        break;
      }
    }
  });

  return count;
}

function countPartialMatches(queryWords, textWords) {
  let count = 0;

  queryWords.forEach(queryWord => {
    if (queryWord.length < 4) return;

    for (let i = 0; i < textWords.length; i++) {
      const word = textWords[i];
      if (
        word.length >= 4 &&
        word !== queryWord &&
        (word.includes(queryWord) || queryWord.includes(word))
      ) {
        count++;
        break;
      }
    }
  });

  return count;
}

function simpleSearch(query) {
  const queryWords = getWords(query);

  return rubricsData
    .map((row, index) => {
      const keywordWords = getWords(row['Ключевые слова']);
      const descriptionWords = getWords(row['Описание']);

      const keywordMatches = countExactMatches(queryWords, keywordWords);
      const descriptionMatches = countExactMatches(queryWords, descriptionWords);
      const totalMatches = keywordMatches + descriptionMatches;

      const keywordMatchData = getMatchData(queryWords, keywordWords);
      const descriptionMatchData = getMatchData(queryWords, descriptionWords);

      return {
        id: `simple-${index}`,
        rubric: row['Рубрика'] || '—',
        code: row['Код рубрики'] || '—',
        keywordsText: row['Ключевые слова'] || '',
        descriptionText: row['Описание'] || '',
        keywordMatches,
        descriptionMatches,
        totalMatches,
        keywordExactFragments: keywordMatchData.exactMatches,
        keywordPartialFragments: [],
        descriptionExactFragments: descriptionMatchData.exactMatches,
        descriptionPartialFragments: []
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
      return b.descriptionMatches - a.descriptionMatches;
    });
}

function smartSearch(query) {
  const queryWords = getWords(query);

  return rubricsData
    .map((row, index) => {
      const keywordWords = getWords(row['Ключевые слова']);
      const descriptionWords = getWords(row['Описание']);

      const keywordExact = countExactMatches(queryWords, keywordWords);
      const descriptionExact = countExactMatches(queryWords, descriptionWords);

      const keywordPartial = countPartialMatches(queryWords, keywordWords);
      const descriptionPartial = countPartialMatches(queryWords, descriptionWords);

      const keywordMatchData = getMatchData(queryWords, keywordWords);
      const descriptionMatchData = getMatchData(queryWords, descriptionWords);

      const score =
        keywordExact * 3 +
        descriptionExact * 2 +
        keywordPartial * 1.5 +
        descriptionPartial * 1;

      return {
        id: `smart-${index}`,
        rubric: row['Рубрика'] || '—',
        code: row['Код рубрики'] || '—',
        keywordsText: row['Ключевые слова'] || '',
        descriptionText: row['Описание'] || '',
        keywordExact,
        descriptionExact,
        keywordPartial,
        descriptionPartial,
        score: Number(score.toFixed(1)),
        keywordExactFragments: keywordMatchData.exactMatches,
        keywordPartialFragments: keywordMatchData.partialMatches,
        descriptionExactFragments: descriptionMatchData.exactMatches,
        descriptionPartialFragments: descriptionMatchData.partialMatches
      };
    })
    .filter(item => item.keywordExact > 0 || item.descriptionExact > 0 || item.score >= 2)
    .sort((a, b) => b.score - a.score);
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
  const selectedRubricsEl = document.getElementById('selectedRubrics');

  const items = Array.from(selectedMap.values());

  if (!items.length) {
    selectedCodesEl.textContent = 'Пока ничего не выбрано';
    selectedItemsEl.innerHTML = '';
    selectedRubricsEl.textContent = 'Пока ничего не выбрано';
    return;
  }

  const uniqueCodes = [...new Set(items.map(item => String(item.code)))];
  const uniqueRubrics = [...new Set(items.map(item => String(item.rubric)))];

  selectedCodesEl.textContent = uniqueCodes.join('\n');
  selectedRubricsEl.textContent = uniqueRubrics.join('\n');

  selectedItemsEl.innerHTML = items
    .map(item => `<li>${escapeHtml(item.rubric)} — <strong>${escapeHtml(item.code)}</strong></li>`)
    .join('');
}

function createResultCard(item, mode) {
  const checked = isChecked(String(item.code), item.rubric);

  const scoreBlock = mode === 'simple'
    ? `
      <div class="meta"><strong>Всего совпадений:</strong> ${item.totalMatches}</div>
      <div class="meta"><strong>Совпадения в ключевых словах:</strong> ${item.keywordMatches}</div>
      <div class="meta"><strong>Совпадения в описании:</strong> ${item.descriptionMatches}</div>
    `
    : `
      <div class="meta"><strong>Общий балл:</strong> ${item.score}</div>
      <div class="meta"><strong>Точные совпадения в ключевых словах:</strong> ${item.keywordExact}</div>
      <div class="meta"><strong>Точные совпадения в описании:</strong> ${item.descriptionExact}</div>
      <div class="meta"><strong>Частичные совпадения в ключевых словах:</strong> ${item.keywordPartial}</div>
      <div class="meta"><strong>Частичные совпадения в описании:</strong> ${item.descriptionPartial}</div>
    `;

  return `
    <div class="result-card">
      <div class="result-top">
        <input
          type="checkbox"
          class="result-checkbox"
          data-code="${escapeHtml(String(item.code))}"
          data-rubric="${escapeHtml(String(item.rubric))}"
          ${checked ? 'checked' : ''}
        />
        <div class="result-main">
          <div><strong>Рубрика:</strong> ${escapeHtml(item.rubric)}</div>
          <div><strong>Код рубрики:</strong> ${escapeHtml(item.code)}</div>
          ${scoreBlock}
          <div class="meta"><strong>Ключевые слова:</strong> ${highlightTextByMatchType(
            item.keywordsText,
            item.keywordExactFragments,
            item.keywordPartialFragments
          )}</div>
          <div class="meta"><strong>Описание:</strong> ${highlightTextByMatchType(
            item.descriptionText,
            item.descriptionExactFragments,
            item.descriptionPartialFragments
          )}</div>
        </div>
      </div>
    </div>
  `;
}

function renderSimpleResults(results) {
  const container = document.getElementById('simpleResults');

  if (!results.length) {
    container.innerHTML = '<div class="empty">Нет совпадений</div>';
    return;
  }

  container.innerHTML = results
    .map(item => createResultCard(item, 'simple'))
    .join('');

  attachCheckboxHandlers();
}

function renderSmartResults(results) {
  const container = document.getElementById('smartResults');

  if (!results.length) {
    container.innerHTML = '<div class="empty">Нет совпадений</div>';
    return;
  }

  container.innerHTML = results
    .map(item => createResultCard(item, 'smart'))
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

  renderSimpleResults(simpleResults);
  renderSmartResults(smartResults);
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

document.getElementById('copyRubricsBtn').addEventListener('click', async function () {
  const text = document.getElementById('selectedRubrics').textContent;

  if (!text || text === 'Пока ничего не выбрано') {
    alert('Нет выбранных рубрик для копирования');
    return;
  }

  try {
    await navigator.clipboard.writeText(text);
    alert('Рубрики скопированы');
  } catch (error) {
    alert('Не удалось скопировать рубрики');
  }
});

document.getElementById('clearSelectionBtn').addEventListener('click', function () {
  selectedMap.clear();
  syncCheckboxes();
  renderSelectedPanel();
});

loadExcel();
renderSelectedPanel();
