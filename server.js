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

// === ГЕНЕРАЦИЯ ОТЧЕТА (ПОЛНАЯ ЛОГИКА КАК В БОТЕ) ===
function generateReport(dfGrid, dfArchive, monthName, year) {
  try {
    console.log("=== НАЧАЛО ФОРМИРОВАНИЯ ОТЧЕТА ===");
    console.log(`Параметры: месяц=${monthName}, год=${year}`);

    // === 1. Объединение данных ===
    console.log("\n1. ОБЪЕДИНЕНИЕ ДАННЫХ ИЗ ГРИДА И АРХИВА");

    // Находим общие колонки
    const commonColumns = dfArchive.columns.filter(col => dfGrid.columns.includes(col));
    console.log("Общие колонки для объединения:", commonColumns);

    if (commonColumns.length === 0) {
      throw new Error("Нет общих колонок для объединения!");
    }

    // Создаем новый массив данных из Грида с общими колонками
    const gridCommonData = dfGrid.data.map(row => {
      const newRow = {};
      commonColumns.forEach(col => {
        newRow[col] = row[col] || null;
      });
      return newRow;
    });

    // Объединяем данные
    const mergedData = [...dfArchive.data, ...gridCommonData];
    console.log(`Объединено данных: Грид=${gridCommonData.length} строк + Архив=${dfArchive.data.length} строк = Итого=${mergedData.length} строк`);

    // Создаем объединённый DataFrame
    let dfMerged = {
      columns: dfArchive.columns,
       mergedData
    };

    // Заменяем пустые значения в Ответственном на "Неизвестно"
    dfMerged.data = dfMerged.data.map(row => {
      row['Ответственный'] = row['Ответственный'] || 'Неизвестно';
      return row;
    });

    // === 2. Преобразование дат ===
    console.log("\n2. ПРЕОБРАЗОВАНИЕ ДАТ:");
    dfMerged.data = dfMerged.data.map(row => {
      // Явно указываем форматы для парсинга
      const createdDate = moment(row['Дата создания'], ['D/M/YY', 'DD/MM/YY', 'YYYY-MM-DD', 'DD.MM.YYYY']);
      const completedDate = moment(row['Выполнена'], ['D/M/YY', 'DD/MM/YY', 'YYYY-MM-DD', 'DD.MM.YYYY']);
      row['Дата создания'] = createdDate.isValid() ? createdDate.toDate() : null;
      row['Выполнена'] = completedDate.isValid() ? completedDate.toDate() : null;
      return row;
    });
    console.log(`Преобразовано дат: ${dfMerged.data.length} строк`);

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
    console.log("\n4. ПОДСЧЕТ СТАТИСТИКИ:");
    const textAuthors = ['Наталия Пятницкая', 'Валентина Кулябина', 'Пятницкая', 'Кулябина'];

    const isTextAuthor = (row) => textAuthors.includes(row['Ответственный']);
    const isDesigner = (row) => !isTextAuthor(row);

    // Фильтруем задачи по периоду
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

    const createdText = dfMerged.data.filter(row =>
      isTextAuthor(row) &&
      row['Дата создания'] &&
      moment(row['Дата создания']).format('YYYY-MM') === monthPeriod
    );

    const completedText = dfMerged.data.filter(row =>
      isTextAuthor(row) &&
      row['Выполнена'] &&
      moment(row['Выполнена']).format('YYYY-MM') === monthPeriod
    );

    console.log("\nДИЗАЙНЕРЫ:");
    console.log(`- Всего задач в объединенном файле: ${dfMerged.data.filter(isDesigner).length}`);
    console.log(`- Создано в отчетном периоде: ${createdDesign.length}`);
    console.log(`- Выполнено в отчетном периоде: ${completedDesign.length}`);

    console.log("\nТЕКСТОВЫЕ ЗАДАЧИ:");
    console.log(`- Всего задач в объединенном файле: ${dfMerged.data.filter(isTextAuthor).length}`);
    console.log(`- Создано в отчетном периоде: ${createdText.length}`);
    console.log(`- Выполнено в отчетном периоде: ${completedText.length}`);

    // === 5. Формирование отчета по дизайнерам ===
    console.log("\n5. ФОРМИРОВАНИЕ ОТЧЕТА ПО ДИЗАЙНЕРАМ:");

    let report = [];

    if (completedDesign.length > 0) {
      // Стандартизируем названия колонок (аналогично боту)
      const colRenameMap = {
        'Кол-во макетов': 'Количество макетов',
        'Кол-во вариантов': 'Количество предложенных вариантов'
      };

      // Группируем данные вручную
      const reportMap = {};

      completedDesign.forEach(row => {
        // Переименовываем колонки в объекте row
        const renamedRow = {};
        Object.keys(row).forEach(key => {
          renamedRow[colRenameMap[key] || key] = row[key];
        });

        const resp = renamedRow['Ответственный'] || 'Неизвестно';

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

        // Обработка макетов
        const makets = parseInt(renamedRow['Количество макетов']) || 0;
        reportMap[resp].Макеты += makets;

        // Обработка вариантов
        const variants = parseInt(renamedRow['Количество предложенных вариантов']) || 0;
        reportMap[resp].Варианты += variants;

        // Обработка оценки
        if (renamedRow['Оценка работы']) {
          const score = parseFloat(renamedRow['Оценка работы']);
          if (!isNaN(score)) {
            reportMap[resp].Оценка += score;
            reportMap[resp].count += 1;
          }
        }
      });

      // Формируем массив отчёта
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

      console.log("Отчет после группировки:");
      console.log(JSON.stringify(report, null, 2));

    } else {
      console.warn("Нет выполненных задач дизайнеров для отчетного периода");
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

СТАТИСТИКА ПО ВЫПОЛНЕННЫМ ЗАДАЧАМ ДИЗАЙНЕРОВ:
(только задачи, завершенные в отчетном периоде)`;

    console.log("\n=== ОТЧЕТ УСПЕШНО СФОРМИРОВАН ===");
    return { report, textReport };

  } catch (error) {
    console.error("ОШИБКА ПРИ ФОРМИРОВАНИИ ОТЧЕТА:", error);
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
       gridData
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
       archiveData
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

    // Отправляем ответ
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
