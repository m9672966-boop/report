const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs-extra');
const path = require('path');
const archiver = require('archiver');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.static('.'));
app.use(express.json());

// Папка для загрузок
const UPLOAD_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) {
  fs.mkdirSync(UPLOAD_DIR, { recursive: true });
}

// Multer для загрузки файлов
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});
const upload = multer({ storage });

// === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===

function mergeGridToArchive(dfGrid, dfArchive) {
  const commonColumns = dfGrid.columns.filter(col => dfArchive.columns.includes(col));
  if (commonColumns.length === 0) {
    console.error("Нет общих колонок для объединения!");
    return dfArchive;
  }

  const dfGridCommon = dfGrid.data.map(row => {
    const newRow = {};
    commonColumns.forEach(col => {
      newRow[col] = row[col] || null;
    });
    return newRow;
  });

  const mergedData = [...dfArchive.data, ...dfGridCommon];
  return { columns: commonColumns, data: mergedData };
}

function generateReport(dfGrid, dfArchive, monthName, year) {
  try {
    // 1. Объединение данных
    let dfMerged = mergeGridToArchive(dfGrid, dfArchive);

    // 2. Преобразование дат
    dfMerged.data = dfMerged.data.map(row => {
      row['Дата создания'] = row['Дата создания'] ? new Date(row['Дата создания']) : null;
      row['Выполнена'] = row['Выполнена'] ? new Date(row['Выполнена']) : null;
      row['Ответственный'] = row['Ответственный'] || 'Неизвестно';
      return row;
    });

    // 3. Определение месяца
    const monthObj = new Date(`${year} ${monthName}`);
    if (isNaN(monthObj.getTime())) {
      throw new Error("Неверный месяц");
    }
    const monthNum = monthObj.getMonth() + 1;
    const monthPeriod = `${year}-${monthNum.toString().padStart(2, '0')}`;

    // 4. Подсчет статистики
    const textAuthors = ['Наталия Пятницкая', 'Валентина Кулябина', 'Пятницкая', 'Кулябина'];
    const isTextAuthor = (row) => textAuthors.includes(row['Ответственный']);
    const isDesigner = (row) => !isTextAuthor(row) || row['Ответственный'] === 'Неизвестно';

    const createdDesign = dfMerged.data.filter(row => 
      isDesigner(row) && 
      row['Дата создания'] && 
      moment(row['Дата создания']).format('YYYY-MM') === monthPeriod
    );

    const completedDesign = dfMerged.data.filter(row => 
      isDesigner(row) && 
      row['Выполнена'] && 
      moment(row['Выполнена']).format('YYYY-MM') === monthPeriod
    );

    // 5. Формирование отчета по дизайнерам
    const reportMap = {};
    completedDesign.forEach(row => {
      const resp = row['Ответственный'];
      if (!reportMap[resp]) {
        reportMap[resp] = { Задачи: 0, Макеты: 0, Варианты: 0, Оценка: 0, count: 0 };
      }
      reportMap[resp].Задачи += 1;
      reportMap[resp].Макеты += row['Количество макетов'] || 0;
      reportMap[resp].Варианты += row['Количество предложенных вариантов'] || 0;
      if (row['Оценка работы']) {
        reportMap[resp].Оценка += parseFloat(row['Оценка работы']);
        reportMap[resp].count += 1;
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
    const mpCardsCount = 0; // Упрощённо — пока 0

    const textReport = `ОТЧЕТ ЗА ${monthName.toUpperCase()} ${year} ГОДА

Дизайнеры:
- Поступило задач: ${createdDesign.length}
- Выполнено задач: ${completedDesign.length}
- Готовых карточек МП: ${mpCardsCount} SKU

Текстовые задачи:
- Поступило: ${dfMerged.data.filter(row => isTextAuthor(row) && row['Дата создания'] && moment(row['Дата создания']).format('YYYY-MM') === monthPeriod).length}
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

    const dfGrid = xlsx.utils.sheet_to_json(gridSheet, { header: 1 });
    const dfArchive = xlsx.utils.sheet_to_json(archiveSheet, { header: 1 });

    // Преобразуем в объекты с именами колонок
    const gridColumns = dfGrid[0];
    const archiveColumns = dfArchive[0];

    const gridData = dfGrid.slice(1).map(row => {
      const obj = {};
      gridColumns.forEach((col, i) => {
        obj[col] = row[i];
      });
      return obj;
    });

    const archiveData = dfArchive.slice(1).map(row => {
      const obj = {};
      archiveColumns.forEach((col, i) => {
        obj[col] = row[i];
      });
      return obj;
    });

    // Генерация отчёта
    const { report, textReport } = generateReport(
      { columns: gridColumns, data: gridData },
      { columns: archiveColumns, data: archiveData },
      month,
      parseInt(year)
    );

    // Создаём временные файлы для скачивания
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

    // Архив для скачивания
    const zipPath = path.join(tempDir, `report_${month}_${year}.zip`);
    const output = fs.createWriteStream(zipPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    archive.pipe(output);
    archive.file(excelPath, { name: `Отчет_${month}_${year}.xlsx` });
    archive.file(txtPath, { name: `Статистика_${month}_${year}.txt` });
    await archive.finalize();

    // Удаляем загруженные файлы
    await fs.unlink(gridPath);
    await fs.unlink(archivePath);

    res.json({
      success: true,
      downloadUrl: `/download?file=${encodeURIComponent(path.basename(zipPath))}`,
      textReport: textReport,
      report: report
    });

  } catch (error) {
    console.error("Ошибка:", error);
    res.status(500).json({ error: error.message });
  }
});

app.get('/download', async (req, res) => {
  const fileName = req.query.file;
  const filePath = path.join(UPLOAD_DIR, fileName);

  if (!fileName || !fs.existsSync(filePath)) {
    return res.status(404).send('Файл не найден');
  }

  res.download(filePath, (err) => {
    if (err) {
      console.error("Ошибка скачивания:", err);
    } else {
      setTimeout(() => {
        fs.unlink(filePath, (err) => {
          if (err) console.error("Ошибка удаления файла:", err);
        });
      }, 5000);
    }
  });
});

app.listen(PORT, () => {
  console.log(`🚀 Сервер запущен на порту ${PORT}`);
});
