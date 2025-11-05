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
const PORT = process.env.PORT || 3000;

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

// === НАДЕЖНАЯ ФУНКЦИЯ ПРЕОБРАЗОВАНИЯ EXCEL ДАТЫ ===
function excelDateToJSDate(serial) {
  if (serial == null || serial === '') return null;
  if (serial instanceof Date) return serial;

  if (typeof serial === 'string') {
    const s = serial.trim();

    // Поддержка DD.MM.YYYY HH:MM:SS
    const datetimeMatch = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (datetimeMatch) {
      const [, day, month, year, hour, minute, second] = datetimeMatch;
      const iso = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}T${hour.padStart(2, '0')}:${minute.padStart(2, '0')}:${(second || '00').padStart(2, '0')}`;
      const date = new Date(iso);
      if (!isNaN(date.getTime())) return date;
    }

    // Поддержка DD.MM.YYYY
    const dateMatch = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (dateMatch) {
      const [, day, month, year] = dateMatch;
      const iso = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
      const date = new Date(iso);
      if (!isNaN(date.getTime())) return date;
    }

    // Поддержка MM/DD/YYYY и MM/DD/YY (для совместимости с Excel)
    const usDateMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (usDateMatch) {
      let [, month, day, year] = usDateMatch;
      if (year.length === 2) {
        year = parseInt(year) >= 70 ? '19' + year : '20' + year;
      }
      const iso = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
      const date = new Date(iso);
      if (!isNaN(date.getTime())) return date;
    }

    // Fallback: стандартный парсинг
    const fallback = new Date(s);
    if (!isNaN(fallback.getTime())) return fallback;
    return null;
  }

  if (typeof serial === 'number') {
    const excelEpochWithError = new Date(1899, 11, 30);
    const utcDays = Math.floor(serial - 1);
    const ms = utcDays * 24 * 60 * 60 * 1000;
    return new Date(excelEpochWithError.getTime() + ms);
  }

  return null;
}

// === ГЕНЕРАЦИЯ ОТЧЕТА ===
function generateReport(dfGrid, dfArchive, monthName, year) {
  console.log("=== НАЧАЛО ФОРМИРОВАНИЯ ОТЧЕТА ===");
  console.log(`Параметры: месяц=${monthName}, год=${year}`);

  // Объединяем данные из обоих файлов
  const allData = [...(dfGrid.data || []), ...(dfArchive.data || [])];

  // Преобразуем даты и нормализуем ответственных
  const processedData = allData.map(row => {
    row['Дата создания'] = excelDateToJSDate(row['Дата создания']);
    row['Выполнена'] = excelDateToJSDate(row['Выполнена']);
    if (!row['Ответственный'] || row['Ответственный'].toString().trim() === '') {
      row['Ответственный'] = 'Неизвестно';
    }
    return row;
  });

  // Определяем период отчёта
  const monthObj = moment(monthName, 'MMMM', true);
  if (!monthObj.isValid()) throw new Error("Неверный месяц");
  const monthNum = monthObj.month() + 1;
  const monthPeriod = `${year}-${monthNum.toString().padStart(2, '0')}`;
  console.log(`Фильтруем по периоду: ${monthPeriod}`);

  // Классификация ответственных
  const textAuthors = ['Наталия Пятницкая', 'Валентина Кулябина', 'Пятницкая', 'Кулябина'];
  const isTextAuthor = (name) => textAuthors.some(ta => name.includes(ta));
  const classify = (name) => {
    if (name === 'Неизвестно') return 'unknown';
    if (isTextAuthor(name)) return 'text';
    return 'designer';
  };

  // Сбор статистики
  const stats = {
    created: { designer: 0, text: 0, unknown: 0 },
    completed: { designer: 0, text: 0, unknown: 0 }
  };

  const reportMap = {};

  for (const row of processedData) {
    const resp = row['Ответственный'];
    const type = classify(resp);

    // Поступившие (из обоих файлов)
    const created = row['Дата создания'];
    if (created && moment(created).isValid() && moment(created).format('YYYY-MM') === monthPeriod) {
      stats.created[type]++;
    }

    // Выполненные (из обоих файлов)
    const completed = row['Выполнена'];
    if (completed && moment(completed).isValid() && moment(completed).format('YYYY-MM') === monthPeriod) {
      stats.completed[type]++;

      if (!reportMap[resp]) {
        reportMap[resp] = { Задачи: 0, Макеты: 0, Варианты: 0, Оценка: 0, count: 0 };
      }
      reportMap[resp].Задачи += 1;
      reportMap[resp].Макеты += parseInt(row['Количество макетов']) || 0;
      reportMap[resp].Варианты += parseInt(row['Количество предложенных вариантов']) || 0;
      if (row['Оценка работы'] != null && row['Оценка работы'] !== '') {
        const score = parseFloat(row['Оценка работы']);
        if (!isNaN(score)) {
          reportMap[resp].Оценка += score;
          reportMap[resp].count += 1;
        }
      }
    }
  }

  console.log("\n📊 СТАТИСТИКА:");
  console.log(`Дизайнеры — создано: ${stats.created.designer}, выполнено: ${stats.completed.designer}`);
  console.log(`Текстовые — создано: ${stats.created.text}, выполнено: ${stats.completed.text}`);
  console.log(`Без ответственного — создано: ${stats.created.unknown}, выполнено: ${stats.completed.unknown}`);

  // Формируем отчёт по выполненным
  let report = Object.keys(reportMap).map(resp => ({
    Ответственный: resp,
    Задачи: reportMap[resp].Задачи,
    Макеты: reportMap[resp].Макеты,
    Варианты: reportMap[resp].Варианты,
    Оценка: reportMap[resp].count > 0 ? (reportMap[resp].Оценка / reportMap[resp].count).toFixed(2) : 0
  }));

  if (report.length > 0) {
    const totalRow = {
      Ответственный: 'ИТОГО',
      Задачи: report.reduce((sum, r) => sum + r.Задачи, 0),
      Макеты: report.reduce((sum, r) => sum + r.Макеты, 0),
      Варианты: report.reduce((sum, r) => sum + r.Варианты, 0),
      Оценка: report.length > 0 ? (report.reduce((sum, r) => sum + parseFloat(r.Оценка), 0) / report.length).toFixed(2) : 0
    };
    report.push(totalRow);
  }

  // Текстовый отчёт
  const textReport = `ОТЧЕТ ЗА ${monthName.toUpperCase()} ${year} ГОДА

Дизайнеры:
- Поступило задач: ${stats.created.designer}
- Выполнено задач: ${stats.completed.designer}

Текстовые задачи:
- Поступило: ${stats.created.text}
- Выполнено: ${stats.completed.text}

Задачи без ответственного:
- Поступило: ${stats.created.unknown}
- Выполнено: ${stats.completed.unknown}

СТАТИСТИКА ПО ВЫПОЛНЕННЫМ ЗАДАЧАМ ДИЗАЙНЕРОВ И ЗАДАЧАМ БЕЗ ОТВЕТСТВЕННОГО:
(только задачи, завершенные в отчетном периоде)`;

  console.log("\n✅ ОТЧЕТ УСПЕШНО СФОРМИРОВАН");
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

    if (!req.files.grid || !req.files.archive) {
      return res.status(400).json({ error: 'Загрузите оба файла' });
    }

    const gridPath = req.files.grid[0].path;
    const archivePath = req.files.archive[0].path;

    const gridWorkbook = xlsx.readFile(gridPath);
    const archiveWorkbook = xlsx.readFile(archivePath);

    const gridSheet = gridWorkbook.Sheets[gridWorkbook.SheetNames[0]];
    const archiveSheet = archiveWorkbook.Sheets[archiveWorkbook.SheetNames[0]];

    if (!gridSheet || !archiveSheet) {
      throw new Error('Один из листов Excel пуст или не найден');
    }

    const allGridRows = xlsx.utils.sheet_to_json(gridSheet, { header: 1, defval: null });
    const allArchiveRows = xlsx.utils.sheet_to_json(archiveSheet, { header: 1, defval: null });

    // Обработка "Грид"
    let gridColumns = [];
    let gridData = [];

    if (allGridRows.length > 0) {
      let headerRowIndex = 0;
      for (let i = 0; i < allGridRows.length; i++) {
        const row = allGridRows[i];
        if (Array.isArray(row) && row.length > 0 && typeof row[0] === 'string' && row[0].trim() !== '') {
          if (row.some(cell => typeof cell === 'string' && cell.includes('Название'))) {
            headerRowIndex = i;
            break;
          }
        }
      }
      gridColumns = allGridRows[headerRowIndex].map(col => typeof col === 'string' ? col.trim() : col);
      if (allGridRows.length > headerRowIndex + 1) {
        gridData = allGridRows.slice(headerRowIndex + 1).map(row => {
          const obj = {};
          gridColumns.forEach((col, i) => {
            if (col && typeof col === 'string') {
              obj[col] = row[i];
            }
          });
          return obj;
        }).filter(row => Object.keys(row).length > 0);
      }
    }

    const dfGrid = { columns: gridColumns, data: gridData || [] };

    // Обработка "Архив"
    let archiveColumns = [];
    let archiveData = [];

    if (allArchiveRows.length > 0) {
      let headerRowIndex = 0;
      for (let i = 0; i < allArchiveRows.length; i++) {
        const row = allArchiveRows[i];
        if (Array.isArray(row) && row.length > 0 && typeof row[0] === 'string' && row[0].trim() !== '') {
          if (row.some(cell => typeof cell === 'string' && cell.includes('Название'))) {
            headerRowIndex = i;
            break;
          }
        }
      }
      archiveColumns = allArchiveRows[headerRowIndex].map(col => typeof col === 'string' ? col.trim() : col);
      if (allArchiveRows.length > headerRowIndex + 1) {
        archiveData = allArchiveRows.slice(headerRowIndex + 1).map(row => {
          const obj = {};
          archiveColumns.forEach((col, i) => {
            if (col && typeof col === 'string') {
              obj[col] = row[i];
            }
          });
          return obj;
        }).filter(row => Object.keys(row).length > 0);
      }
    }

    const dfArchive = { columns: archiveColumns, data: archiveData || [] };

    console.log("Архив: количество строк =", (dfArchive.data || []).length);
    console.log("Грид: количество строк =", (dfGrid.data || []).length);

    const { report, textReport } = generateReport(
      dfGrid,
      dfArchive,
      month,
      parseInt(year)
    );

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
    } else {
      console.warn("⚠️ KAITEN_CARD_ID не задан — файлы не будут загружены в Kaiten");
    }

    await fs.unlink(gridPath);
    await fs.unlink(archivePath);
    await fs.remove(tempDir);

    res.json({
      success: true,
      textReport: textReport,
      report: report || []
    });

  } catch (error) {
    console.error("❌ Ошибка в /api/upload:", error.message);
    res.status(500).json({ error: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Сервер запущен на порту ${PORT}`);
});
