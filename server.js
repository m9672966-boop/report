require('dotenv').config();

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs-extra');
const path = require('path');
const moment = require('moment');
const FormData = require('form-data');
const fetch = require('node-fetch');

const app = express();
const PORT = process.env.PORT || 10000;

app.use(cors());
app.use(express.static('.'));
app.use(express.json());

const UPLOAD_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) {
  fs.mkdirSync(UPLOAD_DIR, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});
const upload = multer({ storage });

// === ЗАГРУЗКА ФАЙЛА В KAITEN ===
async function uploadFileToKaiten(filePath, fileName, cardId) {
  try {
    const stats = fs.statSync(filePath);
    if (stats.size === 0) {
      console.error("Файл пустой:", fileName);
      return false;
    }

    const form = new FormData();
    form.append('file', fs.createReadStream(filePath), {
      filename: fileName,
      knownLength: stats.size
    });

    const response = await fetch(`https://panna.kaiten.ru/api/latest/cards/${cardId}/files`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${process.env.KAITEN_API_TOKEN}`,
        'Accept': 'application/json'
      },
      body: form
    });

    if (response.ok) {
      console.log(`✅ Файл "${fileName}" успешно загружен в карточку ${cardId}`);
      return true;
    } else {
      const errorText = await response.text();
      console.error(`❌ Ошибка загрузки "${fileName}": ${response.status} - ${errorText}`);
      return false;
    }
  } catch (error) {
    console.error(`❌ Ошибка при загрузке "${fileName}":`, error.message);
    return false;
  }
}

// === ПАРСЕР ДАТЫ ===
function parseDate(value) {
  if (value == null || value === '') return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return null;
    const dateFromStr = new Date(trimmed);
    if (!isNaN(dateFromStr.getTime())) return dateFromStr;

    const num = parseFloat(trimmed.replace(/,/g, '.'));
    if (!isNaN(num)) {
      const epoch = new Date(1899, 11, 30);
      return new Date(epoch.getTime() + (num - 1) * 24 * 60 * 60 * 1000);
    }
    return null;
  }

  if (typeof value === 'number') {
    const epoch = new Date(1899, 11, 30);
    return new Date(epoch.getTime() + (value - 1) * 24 * 60 * 60 * 1000);
  }

  return null;
}

// === ОЧИСТКА ЗАГОЛОВКА ===
function cleanHeader(str) {
  if (typeof str !== 'string') return '';
  return str.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();
}

// === ГЕНЕРАЦИЯ ОТЧЕТА ===
function generateReport(gridData, archiveData, monthName, year) {
  console.log("=== НАЧАЛО ФОРМИРОВАНИЯ ОТЧЕТА ===");
  console.log(`Параметры: месяц=${monthName}, год=${year}`);

  const allData = [...gridData, ...archiveData];
  console.log(`Объединено строк: ${allData.length}`);

  const processed = allData.map(row => {
    const cleanRow = {};
    for (const key in row) {
      const cleanKey = cleanHeader(key);
      cleanRow[cleanKey] = row[key];
    }
    cleanRow['Дата создания'] = parseDate(cleanRow['Дата создания']);
    cleanRow['Выполнена'] = parseDate(cleanRow['Выполнена']);
    if (!cleanRow['Ответственный'] || cleanRow['Ответственный'].toString().trim() === '') {
      cleanRow['Ответственный'] = 'Неизвестно';
    }
    return cleanRow;
  });

  // 🔍 Поиск целевой задачи (для отладки)
  const target = processed.find(r =>
    typeof r['Название'] === 'string' &&
    r['Название'].includes('Новогодняя овечка')
  );
  if (target) {
    console.log("🎯 Найдена задача с оценкой 10.0:", target['Оценка работы']);
  }

  const monthObj = moment(monthName, 'MMMM', true);
  if (!monthObj.isValid()) throw new Error("Неверный месяц");
  const monthPeriod = `${year}-${(monthObj.month() + 1).toString().padStart(2, '0')}`;

  const textAuthors = ['Наталия Пятницкая', 'Валентина Кулябина', 'Пятницкая', 'Кулябина'];
  const isDesigner = (row) => {
    const resp = row['Ответственный'];
    return resp !== 'Неизвестно' && !textAuthors.includes(resp);
  };

  const completedDesign = processed.filter(row => {
    const completed = row['Выполнена'];
    return (
      isDesigner(row) &&
      completed &&
      moment(completed).isValid() &&
      moment(completed).format('YYYY-MM') === monthPeriod
    );
  });

  // Сбор статистики БЕЗ строки "ИТОГО"
  const reportMap = {};
  for (const row of completedDesign) {
    const resp = row['Ответственный'];
    if (!reportMap[resp]) {
      reportMap[resp] = { Задачи: 0, Макеты: 0, Варианты: 0, Оценка: 0, count: 0 };
    }
    reportMap[resp].Задачи += 1;
    reportMap[resp].Макеты += parseInt(row['Количество макетов']) || 0;
    reportMap[resp].Варианты += parseInt(row['Количество предложенных вариантов']) || 0;

    const scoreRaw = row['Оценка работы'];
    if (scoreRaw != null && scoreRaw !== '') {
      const score = parseFloat(scoreRaw);
      if (!isNaN(score)) {
        reportMap[resp].Оценка += score;
        reportMap[resp].count += 1;
      }
    }
  }

  // Формируем отчёт — ТОЛЬКО дизайнеры, БЕЗ "ИТОГО"
  const report = Object.keys(reportMap).map(resp => ({
    Ответственный: resp,
    Задачи: reportMap[resp].Задачи,
    Макеты: reportMap[resp].Макеты,
    Варианты: reportMap[resp].Варианты,
    Оценка: reportMap[resp].count > 0 ? (reportMap[resp].Оценка / reportMap[resp].count).toFixed(2) : '—'
  }));

  const textReport = `ОТЧЕТ ЗА ${monthName.toUpperCase()} ${year} ГОДА\n\nДизайнеры — выполнено задач: ${completedDesign.length}`;

  console.log("✅ Отчёт сформирован");
  return { report, textReport };
}

// === МАРШРУТЫ ===

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/report', (req, res) => {
  res.sendFile(path.join(__dirname, 'report.html'));
});

app.post('/api/upload', upload.fields([
  { name: 'grid', maxCount: 1 },
  { name: 'archive', maxCount: 1 }
]), async (req, res) => {
  try {
    const { month, year } = req.body;
    if (!req.files?.grid || !req.files?.archive) {
      return res.status(400).json({ error: 'Загрузите оба файла' });
    }

    const gridPath = req.files.grid[0].path;
    const archivePath = req.files.archive[0].path;

    const gridWB = xlsx.readFile(gridPath);
    const archiveWB = xlsx.readFile(archivePath);

    const gridSheet = gridWB.Sheets[gridWB.SheetNames[0]];
    const archiveSheet = archiveWB.Sheets[archiveWB.SheetNames[0]];

    // Читаем как объекты — экономим память
    const gridData = xlsx.utils.sheet_to_json(gridSheet, { defval: '' });
    const archiveData = xlsx.utils.sheet_to_json(archiveSheet, { defval: '' });

    const { report, textReport } = generateReport(gridData, archiveData, month, parseInt(year));

    const tempDir = path.join(UPLOAD_DIR, `temp_${Date.now()}`);
    await fs.mkdir(tempDir);

    const ws = xlsx.utils.json_to_sheet(report);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "Отчёт");
    const excelPath = path.join(tempDir, `Отчет_${month}_${year}.xlsx`);
    xlsx.writeFile(wb, excelPath);

    const txtPath = path.join(tempDir, `Статистика_${month}_${year}.txt`);
    await fs.writeFile(txtPath, textReport, 'utf8');

    const cardId = process.env.KAITEN_CARD_ID;
    if (cardId) {
      await uploadFileToKaiten(excelPath, `Отчет_${month}_${year}.xlsx`, cardId);
      await uploadFileToKaiten(txtPath, `Статистика_${month}_${year}.txt`, cardId);
    }

    await fs.unlink(gridPath);
    await fs.unlink(archivePath);
    await fs.remove(tempDir);

    res.json({ success: true, textReport, report });

  } catch (error) {
    console.error("❌ Ошибка:", error.message);
    res.status(500).json({ error: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Сервер запущен на порту ${PORT}`);
});
