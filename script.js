let rubricsData = [];

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
    document.getElementById('simpleResults').innerHTML = '<p>Ошибка загрузки Excel-файла</p>';
    document.getElementById('smartResults').innerHTML = '<p>Ошибка загрузки Excel-файла</p>';
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

function simpleSearch(query) {
  const
