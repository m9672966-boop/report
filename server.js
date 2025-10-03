require('dotenv').config();

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs-extra');
const path = require('path');
const archiver = require('archiver');
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

// === ГЕНЕРАЦИЯ ОТЧЕТА ===
function generateReport(dfGrid, dfArchive, monthName, year) {
  try {
    console.log("=== НАЧАЛО ФОРМИРОВАНИЯ ОТЧЕТА ===");
    console.log(`Параметры: месяц=${monthName}, год=${year}`);

    // === 1. Объединение данных ===
    console.log("\n1. ОБЪЕДИНЕНИЕ ДАННЫХ ИЗ ГРИДА И АРХИВА");
    let dfMerged = {
      columns: dfArchive.columns,
      data: [...(dfArchive.data || [])]
    };
    console.log("Используем только данные из Архива для отчета");
    console.log(`Количество строк в Архиве: ${dfMerged.data.length}`);

    // === 2. Преобразование дат и обработка пустых ответственных ===
    console.log("\n2. ПРЕОБРАЗОВАНИЕ ДАТ И ОБРАБОТКА ПУСТЫХ ОТВЕТСТВЕННЫХ:");

    function excelDateToJSDate(serial) {
      if (serial === null || serial === undefined) return null;
      if (typeof serial === 'number') {
        const excelEpochWithError = new Date(1899, 11, 30);
        const utcDays = Math.floor(serial - 1);
        const milliseconds = utcDays * 24 * 60 * 60 * 1000;
        return new Date(excelEpochWithError.getTime() + milliseconds);
      }
      return null;
    }

    dfMerged.data = (dfMerged.data || []).map((row, index) => {
      row['Дата создания'] = excelDateToJSDate(row['Дата создания']);
      row['Выполнена'] = excelDateToJSDate(row['Выполнена']);
      if (!row['Ответственный'] || row['Ответственный'].toString().trim() === '') {
        row['Ответственный'] = 'Неизвестно';
      }
      return row;
    });

    // === 3. Определение месяца ===
    console.log("\n3. ОПРЕДЕЛЕНИЕ МЕСЯЦА ОТЧЕТА:");
    const monthObj = moment(monthName, 'MMMM', true);
    if (!monthObj.isValid()) {
      throw new Error("Неверный месяц");
    }
    const monthNum = monthObj.month() + 1;
    const monthPeriod = `${year}-${monthNum.toString().padStart(2, '0')}`;
    console.log(`Период для фильтрации: ${monthPeriod}`);

    // === 4. Подсчет статистики ===
    // === 4. Подсчет статистики ===
console.log("\n4. ПОДСЧЕТ СТАТИСТИКИ:");
const textAuthors = ['Наталия Пятницкая', 'Валентина Кулябина', 'Пятницкая', 'Кулябина'];

const isTextAuthor = (row) => textAuthors.includes(row['Ответственный']);
const isDesigner = (row) => !isTextAuthor(row) && row['Ответственный'] !== 'Неизвестно';
const isUnknown = (row) => row['Ответственный'] === 'Неизвестно';

// Логируем общее количество
console.log(`Всего задач в объединённом файле: ${dfMerged.data.length}`);

// Фильтруем задачи по периоду
console.log(`Фильтруем по периоду: ${monthPeriod}`);

const createdDesign = [];
const completedDesign = [];

const createdText = [];
const completedText = [];

const createdUnknown = [];
const completedUnknown = [];

for (const row of dfMerged.data) {
  // Обработка даты создания
  let created = row['Дата создания'];
  if (created) {
    created = excelDateToJSDate(created);
    if (created && moment(created).isValid()) {
      const formatted = moment(created).format('YYYY-MM');
      if (formatted === monthPeriod) {
        if (isDesigner(row)) {
          createdDesign.push(row);
        } else if (isTextAuthor(row)) {
          createdText.push(row);
        } else if (isUnknown(row)) {
          createdUnknown.push(row);
        }
      }
    }
  }

  // Обработка даты выполнения
  let completed = row['Выполнена'];
  if (completed) {
    completed = excelDateToJSDate(completed);
    if (completed && moment(completed).isValid()) {
      const formatted = moment(completed).format('YYYY-MM');
      if (formatted === monthPeriod) {
        if (isDesigner(row)) {
          completedDesign.push(row);
        } else if (isTextAuthor(row)) {
          completedText.push(row);
        } else if (isUnknown(row)) {
          completedUnknown.push(row);
        }
      }
    }
  }
}

console.log("\nДИЗАЙНЕРЫ:");
console.log(`- Всего задач в объединенном файле: ${dfMerged.data.filter(isDesigner).length}`);
console.log(`- Создано в отчетном периоде: ${createdDesign.length}`);
console.log(`- Выполнено в отчетном периоде: ${completedDesign.length}`);

console.log("\nТЕКСТОВЫЕ ЗАДАЧИ:");
console.log(`- Всего задач в объединенном файле: ${dfMerged.data.filter(isTextAuthor).length}`);
console.log(`- Создано: ${createdText.length}`);
console.log(`- Выполнено: ${completedText.length}`);

console.log("\nЗАДАЧИ БЕЗ ОТВЕТСТВЕННОГО:");
console.log(`- Всего задач в объединенном файле: ${dfMerged.data.filter(isUnknown).length}`);
console.log(`- Создано: ${createdUnknown.length}`);
console.log(`- Выполнено: ${completedUnknown.length}`);
    // === 5. Формирование отчета по дизайнерам ===
    console.log("\n5. ФОРМИРОВАНИЕ ОТЧЕТА ПО ДИЗАЙНЕРАМ:");

    let report = [];
    const allCompletedTasks = [...completedDesign, ...completedUnknown];

    if (allCompletedTasks.length > 0) {
      const reportMap = {};

      allCompletedTasks.forEach(row => {
        const resp = row['Ответственный'] || 'Неизвестно';
        if (!reportMap[resp]) {
          reportMap[resp] = { Задачи: 0, Макеты: 0, Варианты: 0, Оценка: 0, count: 0 };
        }
        reportMap[resp].Задачи += 1;
        reportMap[resp].Макеты += parseInt(row['Количество макетов']) || 0;
        reportMap[resp].Варианты += parseInt(row['Количество предложенных вариантов']) || 0;
        if (row['Оценка работы']) {
          const score = parseFloat(row['Оценка работы']);
          if (!isNaN(score)) {
            reportMap[resp].Оценка += score;
            reportMap[resp].count += 1;
          }
        }
      });

      report = Object.keys(reportMap).map(resp => {
        const item = reportMap[resp];
        return {
          Ответственный: resp,
          Задачи: item.Задачи,
          Макеты: item.Макеты,
          Варианты: item.Варианты,
          Оценка: item.count > 0 ? (item.Оценка / item.count).toFixed(2) : 0
        };
      });
    } else {
      console.warn("Нет выполненных задач для отчетного периода");
    }

    // Добавляем итоговую строку
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

    // === 6. Формирование текстового отчета ===
    const textReport = `ОТЧЕТ ЗА ${monthName.toUpperCase()} ${year} ГОДА

Дизайнеры:
- Поступило задач: ${createdDesign.length}
- Выполнено задач: ${completedDesign.length}

Текстовые задачи:
- Поступило: ${createdText.length}
- Выполнено: ${completedText.length}

Задачи без ответственного:
- Поступило: ${createdUnknown.length}
- Выполнено: ${completedUnknown.length}

СТАТИСТИКА ПО ВЫПОЛНЕННЫМ ЗАДАЧАМ ДИЗАЙНЕРОВ И ЗАДАЧАМ БЕЗ ОТВЕТСТВЕННОГО:
(только задачи, завершенные в отчетном периоде)`;

    console.log("\n=== ОТЧЕТ УСПЕШНО СФОРМИРОВАН ===");
    return { report, textReport };

  } catch (error) {
    console.error("ОШИБКА ПРИ ФОРМИРОВАНИИ ОТЧЕТА:", error.message);
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
      gridColumns = allGridRows[headerRowIndex];
      if (allGridRows.length > headerRowIndex + 1) {
        gridData = allGridRows.slice(headerRowIndex + 1).map(row => {
          const obj = {};
          gridColumns.forEach((col, i) => {
            if (col && typeof col === 'string') {
              obj[col.trim()] = row[i];
            }
          });
          return obj;
        }).filter(row => Object.keys(row).length > 0);
      }
    }

    const dfGrid = { columns: gridColumns, data: gridData || [] };

    // Обработка "Архив" — ИСПРАВЛЕНО!
    let archiveColumns = [];
    let archiveData = [];

    if (allArchiveRows.length > 0) {
      let headerRowIndex = 0;
      for (let i = 0; i < allArchiveRows.length; i++) {
        const row = allArchiveRows[i]; // ✅ ИСПРАВЛЕНО: было allGridRows[i]
        if (Array.isArray(row) && row.length > 0 && typeof row[0] === 'string' && row[0].trim() !== '') {
          if (row.some(cell => typeof cell === 'string' && cell.includes('Название'))) {
            headerRowIndex = i;
            break;
          }
        }
      }

      archiveColumns = allArchiveRows[headerRowIndex];
      if (allArchiveRows.length > headerRowIndex + 1) {
        archiveData = allArchiveRows.slice(headerRowIndex + 1).map(row => {
          const obj = {};
          archiveColumns.forEach((col, i) => {
            if (col && typeof col === 'string') {
              obj[col.trim()] = row[i];
            }
          });
          return obj;
        }).filter(row => Object.keys(row).length > 0);
      }
    }

    const dfArchive = { columns: archiveColumns, data: archiveData || [] };

    console.log("Архив: колонки =", dfArchive.columns);
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
