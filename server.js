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

    // Исправлено: убраны пробелы в URL
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

// === ГЕНЕРАЦИЯ ОТЧЁТА ===
function generateReport(dfGrid, dfArchive, monthName, year) {
  try {
    // Исправлено: правильно назначаем data
    let dfMerged = { columns: dfArchive.columns, data: dfArchive.data };

    // Проверка данных
    if (!dfMerged.data || !Array.isArray(dfMerged.data)) {
      throw new Error("Данные из файла 'Архив.xlsx' отсутствуют или повреждены");
    }

    console.log("Обработано строк из Архива:", dfMerged.data.length);

    // Преобразование дат с помощью moment
    dfMerged.data = dfMerged.data.map(row => {
      const createdDate = moment(row['Дата создания']);
      const completedDate = moment(row['Выполнена']);
      row['Дата создания'] = createdDate.isValid() ? createdDate.toDate() : null;
      row['Выполнена'] = completedDate.isValid() ? completedDate.toDate() : null;
      row['Ответственный'] = row['Ответственный'] || 'Неизвестно';
      return row;
    });

    // Определение месяца
    const monthObj = moment(monthName, 'MMMM', true);
    if (!monthObj.isValid()) {
      throw new Error("Неверный месяц");
    }
    const monthNum = monthObj.month() + 1;
    const monthPeriod = `${year}-${monthNum.toString().padStart(2, '0')}`;

    // Подсчет статистики
    const textAuthors = ['Наталия Пятницкая', 'Валентина Кулябина', 'Пятницкая', 'Кулябина'];
    const isTextAuthor = (row) => textAuthors.includes(row['Ответственный']);
    const isDesigner = (row) => !isTextAuthor(row) || row['Ответственный'] === 'Неизвестно';

    // Для статистики "поступило" — используем Грид
    const createdDesign = dfGrid.data.filter(row => 
      isDesigner({ Ответственный: row['Ответственный'] || 'Неизвестно' }) &&
      row['Дата создания'] && 
      moment(row['Дата создания']).format('YYYY-MM') === monthPeriod
    );

    const completedDesign = dfMerged.data.filter(row => 
      isDesigner(row) && 
      row['Выполнена'] && 
      moment(row['Выполнена']).format('YYYY-MM') === monthPeriod
    );

    // Формирование отчета по дизайнерам
    const reportMap = {};
    completedDesign.forEach(row => {
      const resp = row['Ответственный'];
      if (!reportMap[resp]) {
        reportMap[resp] = { 
          Задачи: 0, 
          Макеты: 0, 
          Варианты: 0, 
          Оценка: 0, 
          count: 0 
        };
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

    const report = Object.keys(reportMap).map(resp => {
      const item = reportMap[resp];
      return {
        Ответственный: resp,
        Задачи: item.Задачи,
        Макеты: item.Макеты,
        Варианты: item.Варианты,
        Оценка: item.count > 0 ? (item.Оценка / item.count).toFixed(2) : 0
      };
    });

    // Итоговая строка
    const totalRow = {
      Ответственный: 'ИТОГО',
      Задачи: report.reduce((sum, r) => sum + r.Задачи, 0),
      Макеты: report.reduce((sum, r) => sum + r.Макеты, 0),
      Варианты: report.reduce((sum, r) => sum + r.Варианты, 0),
      Оценка: report.length > 0 ? (report.reduce((sum, r) => sum + parseFloat(r.Оценка), 0) / report.length).toFixed(2) : 0
    };
    report.push(totalRow);

    // Текстовый отчёт
    const mpCardsCount = 0;

    const textReport = `ОТЧЕТ ЗА ${monthName.toUpperCase()} ${year} ГОДА

Дизайнеры:
- Поступило задач: ${createdDesign.length}
- Выполнено задач: ${completedDesign.length}
- Готовых карточек МП: ${mpCardsCount} SKU

Текстовые задачи:
- Поступило: ${dfGrid.data.filter(row => isTextAuthor({ Ответственный: row['Ответственный'] || 'Неизвестно' }) && row['Дата создания'] && moment(row['Дата создания']).format('YYYY-MM') === monthPeriod).length}
- Выполнено: ${dfMerged.data.filter(row => isTextAuthor(row) && row['Выполнена'] && moment(row['Выполнена']).format('YYYY-MM') === monthPeriod).length}

СТАТИСТИКА ПО ВЫПОЛНЕННЫМ ЗАДАЧАМ ДИЗАЙНЕРОВ:
(только задачи, завершенные в отчетном периоде)`;

    return { report, textReport };
  } catch (error) {
    console.error("Ошибка генерации отчёта:", error);
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

    // Чтение Excel файлов
    const gridWorkbook = xlsx.readFile(gridPath);
    const archiveWorkbook = xlsx.readFile(archivePath);

    const gridSheet = gridWorkbook.Sheets[gridWorkbook.SheetNames[0]];
    const archiveSheet = archiveWorkbook.Sheets[archiveWorkbook.SheetNames[0]];

    if (!gridSheet || !archiveSheet) {
      throw new Error('Один из листов Excel пуст или не найден');
    }

    const allGridRows = xlsx.utils.sheet_to_json(gridSheet, { header: 1 });
    const allArchiveRows = xlsx.utils.sheet_to_json(archiveSheet, { header: 1 });

    // Обработка "Грид"
    let gridColumns = [];
    let gridData = [];

    if (allGridRows.length > 0) {
      gridColumns = allGridRows[0];
      if (allGridRows.length > 1) {
        gridData = allGridRows.slice(1).map(row => {
          const obj = {};
          gridColumns.forEach((col, i) => {
            obj[col] = row[i];
          });
          return obj;
        });
      }
    }

    const dfGrid = {
      columns: gridColumns,
      data: gridData  // Исправлено: явно указан ключ
    };

    // Обработка "Архив"
    let archiveColumns = [];
    let archiveData = [];

    if (allArchiveRows.length > 0) {
      archiveColumns = allArchiveRows[0];
      if (allArchiveRows.length > 1) {
        archiveData = allArchiveRows.slice(1).map(row => {
          const obj = {};
          archiveColumns.forEach((col, i) => {
            obj[col] = row[i];
          });
          return obj;
        });
      }
    }

    const dfArchive = {
      columns: archiveColumns,
      data: archiveData  // Исправлено: явно указан ключ
    };

    // Логирование для отладки
    console.log("Архив: колонки =", dfArchive.columns);
    console.log("Архив: количество строк =", dfArchive.data?.length || 0);
    console.log("Грид: количество строк =", dfGrid.data?.length || 0);

    // Генерация отчёта
    const { report, textReport } = generateReport(
      dfGrid,
      dfArchive,
      month,
      parseInt(year)
    );

    // Создаём временные файлы
    const tempDir = path.join(UPLOAD_DIR, `temp_${Date.now()}`);
    await fs.mkdir(tempDir);

    // Excel файл
    const ws = xlsx.utils.json_to_sheet(report);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "Отчёт");
    const excelPath = path.join(tempDir, `Отчет_${month}_${year}.xlsx`);
    xlsx.writeFile(wb, excelPath);

    // TXT файл
    const txtPath = path.join(tempDir, `Статистика_${month}_${year}.txt`);
    await fs.writeFile(txtPath, textReport, 'utf8');

    // ID карточки из переменных окружения
    const cardId = process.env.KAITEN_CARD_ID;

    if (cardId) {
      // Загружаем файлы в Kaiten
      await uploadFileToKaiten(excelPath, `Отчет_${month}_${year}.xlsx`, cardId);
      await uploadFileToKaiten(txtPath, `Статистика_${month}_${year}.txt`, cardId);
    } else {
      console.warn("⚠️ KAITEN_CARD_ID не задан — файлы не будут загружены в Kaiten");
    }

    // Удаляем временные файлы
    await fs.unlink(gridPath);
    await fs.unlink(archivePath);
    await fs.remove(tempDir);

    res.json({
      success: true,
      textReport: textReport,
      report: report
    });

  } catch (error) {
    console.error("❌ Ошибка в /api/upload:", error);
    res.status(500).json({ error: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Сервер запущен на порту ${PORT}`);
});
