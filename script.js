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
            exactMatches
