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
    serial = serial.trim();
    const dateFromStr = new Date(serial);
    if (!isNaN(dateFromStr.getTime())) return dateFromStr;

    const parsed = parseFloat(serial.replace(/,/g, '.'));
    if (!isNaN(parsed)) {
      serial = parsed;
    } else {
      return null;
    }
  }

  if (typeof serial === 'number') {
    const excelEpochWithError = new Date(1899, 11, 30);
    const utcDays = Math.floor(serial - 1);
    const milliseconds = utcDays * 24 * 60 * 60 * 1000;
    return new Date(excelEpochWithError.getTime() + milliseconds);
  }

  return null;
}

// === ГЕНЕРАЦИЯ ОТЧЕТА ===
function generateReport(dfGrid, dfArchive, monthName, year) {
  try {
    console.log("=== НАЧАЛО ФОРМИРОВАНИЯ ОТЧЕТА ===");
    console.log(`Параметры: месяц=${monthName}, год=${year}`);

    const allData = [...(dfGrid.data || []), ...(dfArchive.data || [])];
    console.log(`Объединено строк: ${allData.length} (Грид: ${dfGrid.data?.length || 0}, Архив: ${dfArchive.data?.length || 0})`);

    // === Нормализация заголовков ===
    const cleanHeader = (str) => {
      if (typeof str !== 'string') return '';
      return str
        .replace(/\u00A0/g, ' ')     // неразрывные пробелы → обычные
        .replace(/\s+/g, ' ')        // несколько пробелов → один
        .trim();
    };

    // Применяем нормализацию ко всем строкам
    const processedData = allData.map(row => {
      const cleanedRow = {};
      for (const key in row) {
        const cleanKey = cleanHeader(key);
        cleanedRow[cleanKey] = row[key];
      }
      cleanedRow['Дата создания'] = excelDateToJSDate(cleanedRow['Дата создания']);
      cleanedRow['Выполнена'] = excelDateToJSDate(cleanedRow['Выполнена']);
      if (!cleanedRow['Ответственный'] || cleanedRow['Ответственный'].toString().trim() === '') {
        cleanedRow['Ответственный'] = 'Неизвестно';
      }
      return cleanedRow;
    });

    // 🔍 ЯВНЫЙ ПОИСК ЦЕЛЕВОЙ ЗАДАЧИ
    const targetTask = processedData.find(row =>
      typeof row['Название'] === 'string' &&
      row['Название'].includes('Новогодняя овечка')
    );

    if (targetTask) {
      console.log("🎯 ЦЕЛЕВАЯ ЗАДАЧА НАЙДЕНА:");
      console.log({
        Название: targetTask['Название'],
        Ответственный: targetTask['Ответственный'],
        Выполнена_RAW: allData.find(r => r['Название'] === targetTask['Название'])?.['Выполнена'],
        Выполнена_parsed: targetTask['Выполнена'],
        Оценка: targetTask['Оценка работы'],
        'Оценка (тип)': typeof targetTask['Оценка работы'],
        Колонки: Object.keys(targetTask).filter(k => k.includes('Оценка'))
      });
    } else {
      console.log("❌ ЦЕЛЕВАЯ ЗАДАЧА НЕ НАЙДЕНА В ОБЪЕДИНЁННЫХ ДАННЫХ");
    }

    // === 3. ОПРЕДЕЛЕНИЕ ПЕРИОДА ===
    const monthObj = moment(monthName, 'MMMM', true);
    if (!monthObj.isValid()) throw new Error("Неверный месяц");
    const monthNum = monthObj.month() + 1;
    const monthPeriod = `${year}-${monthNum.toString().padStart(2, '0')}`;
    console.log(`Фильтруем по периоду: ${monthPeriod}`);

    // === 4. КЛАССИФИКАЦИЯ ===
    const textAuthors = ['Наталия Пятницкая', 'Валентина Кулябина', 'Пятницкая', 'Кулябина'];
    const isTextAuthor = (row) => textAuthors.includes(row['Ответственный']);
    const isDesigner = (row) => !isTextAuthor(row) && row['Ответственный'] !== 'Неизвестно';
    const isUnknown = (row) => row['Ответственный'] === 'Неизвестно';

    const completedDesign = [];
    const completedUnknown = [];

    for (const row of processedData) {
      const completed = row['Выполнена'];
      if (completed && moment(completed).isValid()) {
        if (moment(completed).format('YYYY-MM') === monthPeriod) {
          if (isDesigner(row)) completedDesign.push(row);
          else if (isUnknown(row)) completedUnknown.push(row);
        }
      }
    }

    console.log(`Дизайнеры — выполнено: ${completedDesign.length}`);
    console.log(`Без ответственного — выполнено: ${completedUnknown.length}`);

    // === 6. ФОРМИРОВАНИЕ ОТЧЁТА ===
    const allCompleted = [...completedDesign, ...completedUnknown];
    let report = [];

    if (allCompleted.length > 0) {
      const reportMap = {};
      for (const row of allCompleted) {
        const resp = row['Ответственный'] || 'Неизвестно';
        if (!reportMap[resp]) {
          reportMap[resp] = { Задачи: 0, Макеты: 0, Варианты: 0, Оценка: 0, count: 0 };
        }
        reportMap[resp].Задачи += 1;
        reportMap[resp].Макеты += parseInt(row['Количество макетов']) || 0;
        reportMap[resp].Варианты += parseInt(row['Количество предложенных вариантов']) || 0;

        let scoreValue = null;
        if (row['Оценка работы'] !== undefined && row['Оценка работы'] !== null && row['Оценка работы'] !== '') {
          scoreValue = parseFloat(row['Оценка работы']);
        }
        if (scoreValue !== null && !isNaN(scoreValue)) {
          reportMap[resp].Оценка += scoreValue;
          reportMap[resp].count += 1;
          console.log(`✅ Учёт оценки: ${resp} → ${scoreValue}`);
        }
      }

      report = Object.keys(reportMap).map(resp => ({
        Ответственный: resp,
        Задачи: reportMap[resp].Задачи,
        Макеты: reportMap[resp].Макеты,
        Варианты: reportMap[resp].Варианты,
        Оценка: reportMap[resp].count > 0 ? (reportMap[resp].Оценка / reportMap[resp].count).toFixed(2) : '—'
      }));
    }

    // Итог
    if (report.length > 0) {
      const validReports = report.filter(r => r.Оценка !== '—');
      const totalRow = {
        Ответственный: 'ИТОГО',
        Задачи: report.reduce((sum, r) => sum + r.Задачи, 0),
        Макеты: report.reduce((sum, r) => sum + r.Макеты, 0),
        Варианты: report.reduce((sum, r) => sum + r.Варианты, 0),
        Оценка: validReports.length > 0
          ? (validReports.reduce((sum, r) => sum + parseFloat(r.Оценка), 0) / validReports.length).toFixed(2)
          : '—'
      };
      report.push(totalRow);
    }

    const textReport = `ОТЧЕТ ЗА ${monthName.toUpperCase()} ${year} ГОДА

Дизайнеры:
- Выполнено задач: ${completedDesign.length}

СТАТИСТИКА ПО ВЫПОЛНЕННЫМ ЗАДАЧАМ ДИЗАЙНЕРОВ И ЗАДАЧАМ БЕЗ ОТВЕТСТВЕННОГО`;

    console.log("\n✅ ОТЧЕТ УСПЕШНО СФОРМИРОВАН");
    return { report, textReport };

  } catch (error) {
    console.error("❌ ОШИБКА В generateReport:", error.message);
    throw error;
  }
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
      gridColumns = allGridRows[headerRowIndex].map(col => col ? col.toString().trim() : '');
      if (allGridRows.length > headerRowIndex + 1) {
        gridData = allGridRows.slice(headerRowIndex + 1).map(row => {
          const obj = {};
          gridColumns.forEach((col, i) => {
            if (col && col !== '') {
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
        const row = allArchiveRows[i]; // ✅ исправлено
        if (Array.isArray(row) && row.length > 0 && typeof row[0] === 'string' && row[0].trim() !== '') {
          if (row.some(cell => typeof cell === 'string' && cell.includes('Название'))) {
            headerRowIndex = i;
            break;
          }
        }
      }

      archiveColumns = allArchiveRows[headerRowIndex].map(col => col ? col.toString().trim() : '');
      if (allArchiveRows.length > headerRowIndex + 1) {
        archiveData = allArchiveRows.slice(headerRowIndex + 1).map(row => {
          const obj = {};
          archiveColumns.forEach((col, i) => {
            if (col && col !== '') {
              obj[col] = row[i];
            }
          });
          return obj;
        }).filter(row => Object.keys(row).length > 0);
      }
    }

    const dfArchive = { columns: archiveColumns, data: archiveData || [] };

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
      console.warn("⚠️ KAITEN_CARD_ID не задан");
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
