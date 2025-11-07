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

// === НАДЁЖНЫЙ ПАРСЕР ДАТ ИЗ EXCEL ===
function excelDateToJSDate(serial) {
  if (serial == null || serial === '') return null;
  if (serial instanceof Date && !isNaN(serial.getTime())) return serial;

  if (typeof serial === 'string') {
    const trimmed = serial.trim();
    if (trimmed === '') return null;
    const dateFromStr = new Date(trimmed);
    if (!isNaN(dateFromStr.getTime())) return dateFromStr;
    const parsed = parseFloat(trimmed.replace(/,/g, '.'));
    if (!isNaN(parsed)) {
      serial = parsed;
    } else {
      return null;
    }
  }

  if (typeof serial === 'number') {
    const excelEpochUTC = Date.UTC(1899, 11, 30);
    const utcDays = Math.floor(serial);
    const fractionalDay = serial - utcDays;
    const milliseconds = utcDays * 24 * 60 * 60 * 1000 + fractionalDay * 24 * 60 * 60 * 1000;
    return new Date(excelEpochUTC + milliseconds);
  }

  return null;
}

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

// === ОЧИСТКА ЗАГОЛОВКА ===
function cleanHeader(str) {
  if (typeof str !== 'string') return '';
  return str.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();
}

// === ГЕНЕРАЦИЯ ПОЛНОГО ОТЧЁТА ===
function generateReport(gridData, archiveData, monthName, year) {
  console.log("=== НАЧАЛО ФОРМИРОВАНИЯ ОТЧЕТА ===");
  console.log(`Месяц: ${monthName}, Год: ${year}`);

  const allData = [...gridData, ...archiveData];
  console.log(`Всего строк: ${allData.length}`);

  const textAuthors = ['Наталия Пятницкая', 'Валентина Кулябина', 'Пятницкая', 'Кулябина'];
  const isTextAuthor = (resp) => textAuthors.includes(resp);
  const isDesigner = (resp) => resp !== 'Неизвестно' && !isTextAuthor(resp);

  const processed = allData.map(row => {
    const cleanRow = {};
    for (const key in row) {
      cleanRow[cleanHeader(key)] = row[key];
    }
    cleanRow['Дата создания'] = excelDateToJSDate(cleanRow['Дата создания']);
    cleanRow['Выполнена'] = excelDateToJSDate(cleanRow['Выполнена']);
    cleanRow['Ответственный'] = (!cleanRow['Ответственный'] || cleanRow['Ответственный'].toString().trim() === '')
      ? 'Неизвестно'
      : cleanRow['Ответственный'].toString().trim();
    return cleanRow;
  });

  const monthPeriod = moment(`${year}-${monthName}`, 'YYYY-MMMM').format('YYYY-MM');
  if (!moment(monthPeriod, 'YYYY-MM', true).isValid()) {
    throw new Error("Неверный месяц или год");
  }

  let createdDesign = 0, completedDesign = 0;
  let createdText = 0, completedText = 0;
  let createdUnknown = 0, completedUnknown = 0;

  const completedDesignRows = [];

  for (const row of processed) {
    const resp = row['Ответственный'];
    const created = row['Дата создания'];
    const completed = row['Выполнена'];

    // Поступило
    if (created && moment(created).isValid() && moment(created).format('YYYY-MM') === monthPeriod) {
      if (isDesigner(resp)) createdDesign++;
      else if (isTextAuthor(resp)) createdText++;
      else createdUnknown++;
    }

    // Выполнено
    if (completed && moment(completed).isValid() && moment(completed).format('YYYY-MM') === monthPeriod) {
      if (isDesigner(resp)) {
        completedDesign++;
        completedDesignRows.push(row);
      } else if (isTextAuthor(resp)) {
        completedText++;
      } else {
        completedUnknown++;
      }
    }
  }

  // === Агрегация по дизайнерам (только выполненные) ===
  const reportMap = {};
  for (const row of completedDesignRows) {
    const resp = row['Ответственный'];
    if (!reportMap[resp]) {
      reportMap[resp] = { Задачи: 0, Макеты: 0, Варианты: 0, Оценка: 0, count: 0 };
    }
    reportMap[resp].Задачи += 1;
    reportMap[resp].Макеты += parseInt(row['Количество макетов']) || 0;
    reportMap[resp].Варианты += parseInt(row['Количество предложенных вариантов']) || 0;
    const score = parseFloat(row['Оценка работы']);
    if (!isNaN(score)) {
      reportMap[resp].Оценка += score;
      reportMap[resp].count += 1;
    }
  }

  const report = [];
  for (const resp in reportMap) {
    report.push({
      Ответственный: resp,
      Задачи: reportMap[resp].Задачи,
      Макеты: reportMap[resp].Макеты,
      Варианты: reportMap[resp].Варианты,
      Оценка: reportMap[resp].count > 0 ? (reportMap[resp].Оценка / reportMap[resp].count).toFixed(2) : '—'
    });
  }

  // === Текстовый отчёт ===
  const textReport = `ОТЧЕТ ЗА ${monthName.toUpperCase()} ${year} ГОДА

Дизайнеры:
- Поступило задач: ${createdDesign}
- Выполнено задач: ${completedDesign}

Текстовые задачи:
- Поступило: ${createdText}
- Выполнено: ${completedText}

Задачи без ответственного:
- Поступило: ${createdUnknown}
- Выполнено: ${completedUnknown}`;

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

    const gridDataRaw = xlsx.utils.sheet_to_json(gridSheet, { defval: '' });
    const archiveDataRaw = xlsx.utils.sheet_to_json(archiveSheet, { defval: '' });

    const neededColumns = [
      'Название',
      'Дата создания',
      'Выполнена',
      'Ответственный',
      'Оценка работы',
      'Количество макетов',
      'Количество предложенных вариантов'
    ];

    const filterColumns = (rows) => rows.map(row => {
      const filtered = {};
      neededColumns.forEach(col => {
        filtered[col] = row[col];
      });
      return filtered;
    });

    const gridData = filterColumns(gridDataRaw);
    const archiveData = filterColumns(archiveDataRaw);

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

app.listen(PORT, '0.0.0.0', () => {
  console.log(`🚀 Сервер запущен на порту ${PORT}`);
});
