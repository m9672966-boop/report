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

// === ГЕНЕРАЦИЯ ОТЧЕТА (ПОЛНАЯ ЛОГИКА КАК В БОТЕ + ДЕТАЛЬНОЕ ЛОГИРОВАНИЕ) ===
// === ГЕНЕРАЦИЯ ОТЧЕТА ===
function generateReport(dfGrid, dfArchive, monthName, year) {
  try {
    console.log("=== НАЧАЛО ФОРМИРОВАНИЯ ОТЧЕТА ===");
    console.log(`Параметры: месяц=${monthName}, год=${year}`);

    // === 1. Объединение данных ===
    console.log("\n1. ОБЪЕДИНЕНИЕ ДАННЫХ ИЗ ГРИДА И АРХИВА");

    // Используем только данные из Архива для отчета (как в оригинальной логике)
    let dfMerged = {
      columns: dfArchive.columns,
      data: [...(dfArchive.data || [])] // Только архивные данные
    };

    console.log("Используем только данные из Архива для отчета");
    console.log(`Количество строк в Архиве: ${dfMerged.data.length}`);

    // === 2. Преобразование дат ===
    console.log("\n2. ПРЕОБРАЗОВАНИЕ ДАТ:");
    
    // Логируем первые 5 строк для отладки
    console.log("Первые 5 строк до преобразования дат:");
    for (let i = 0; i < Math.min(5, dfMerged.data.length); i++) {
      const row = dfMerged.data[i];
      console.log(`Строка ${i+1}:`);
      console.log(`  Дата создания: ${row['Дата создания']} (тип: ${typeof row['Дата создания']})`);
      console.log(`  Выполнена: ${row['Выполнена']} (тип: ${typeof row['Выполнена']})`);
      console.log(`  Ответственный: ${row['Ответственный']}`);
    }

    dfMerged.data = (dfMerged.data || []).map((row, index) => {
      // Логируем преобразование для первых 10 строк
      if (index < 10) {
        console.log(`\nПреобразование строки ${index+1}:`);
        console.log(`  Исходная дата создания: ${row['Дата создания']}`);
        console.log(`  Исходная дата выполнения: ${row['Выполнена']}`);
      }

      let createdDate = null;
      let completedDate = null;

      // Пробуем разные форматы дат
      const dateFormats = [
        'DD.MM.YYYY', 'D.M.YYYY', 'DD/MM/YYYY', 'D/M/YYYY',
        'YYYY-MM-DD', 'YYYY-M-D', 'MM/DD/YYYY', 'M/D/YYYY'
      ];

      // Преобразование даты создания
      if (row['Дата создания']) {
        for (const format of dateFormats) {
          createdDate = moment(row['Дата создания'], format, true);
          if (createdDate.isValid()) {
            if (index < 10) console.log(`  Дата создания найдена в формате ${format}: ${createdDate.format('YYYY-MM-DD')}`);
            break;
          }
        }
      }

      // Преобразование дата выполнения
      if (row['Выполнена']) {
        for (const format of dateFormats) {
          completedDate = moment(row['Выполнена'], format, true);
          if (completedDate.isValid()) {
            if (index < 10) console.log(`  Дата выполнения найдена в формате ${format}: ${completedDate.format('YYYY-MM-DD')}`);
            break;
          }
        }
      }

      row['Дата создания'] = createdDate && createdDate.isValid() ? createdDate.toDate() : null;
      row['Выполнена'] = completedDate && completedDate.isValid() ? completedDate.toDate() : null;
      
      // Заменяем пустые значения в Ответственном
      row['Ответственный'] = row['Ответственный'] || 'Неизвестно';

      if (index < 10) {
        console.log(`  Преобразованная дата создания: ${row['Дата создания']}`);
        console.log(`  Преобразованная дата выполнения: ${row['Выполнена']}`);
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
    console.log("\n4. ПОДСЧЕТ СТАТИСТИКИ:");
    const textAuthors = ['Наталия Пятницкая', 'Валентина Кулябина', 'Пятницкая', 'Кулябина'];

    const isTextAuthor = (row) => textAuthors.includes(row['Ответственный']);
    const isDesigner = (row) => !isTextAuthor(row);

    // Фильтруем задачи по периоду с детальным логированием
    console.log("\nПоиск задач за период:", monthPeriod);

    const createdDesign = (dfMerged.data || []).filter((row, index) => {
      const isMatch = isDesigner(row) &&
        row['Дата создания'] &&
        moment(row['Дата создания']).format('YYYY-MM') === monthPeriod;

      if (isMatch && index < 10) {
        console.log(`  Найдена созданная задача дизайнера ${index+1}:`);
        console.log(`    Ответственный: ${row['Ответственный']}`);
        console.log(`    Дата создания: ${row['Дата создания']}`);
        console.log(`    Название: ${row['Название']}`);
      }
      return isMatch;
    });

    const completedDesign = (dfMerged.data || []).filter((row, index) => {
      const isMatch = isDesigner(row) &&
        row['Выполнена'] &&
        moment(row['Выполнена']).format('YYYY-MM') === monthPeriod;

      if (isMatch && index < 10) {
        console.log(`  Найдена выполненная задача дизайнера ${index+1}:`);
        console.log(`    Ответственный: ${row['Ответственный']}`);
        console.log(`    Дата выполнения: ${row['Выполнена']}`);
        console.log(`    Название: ${row['Название']}`);
      }
      return isMatch;
    });

    const createdText = (dfMerged.data || []).filter((row, index) => {
      const isMatch = isTextAuthor(row) &&
        row['Дата создания'] &&
        moment(row['Дата создания']).format('YYYY-MM') === monthPeriod;

      if (isMatch && index < 10) {
        console.log(`  Найдена созданная текстовая задача ${index+1}:`);
        console.log(`    Ответственный: ${row['Ответственный']}`);
        console.log(`    Дата создания: ${row['Дата создания']}`);
      }
      return isMatch;
    });

    const completedText = (dfMerged.data || []).filter((row, index) => {
      const isMatch = isTextAuthor(row) &&
        row['Выполнена'] &&
        moment(row['Выполнена']).format('YYYY-MM') === monthPeriod;

      if (isMatch && index < 10) {
        console.log(`  Найдена выполненная текстовая задача ${index+1}:`);
        console.log(`    Ответственный: ${row['Ответственный']}`);
        console.log(`    Дата выполнения: ${row['Выполнена']}`);
      }
      return isMatch;
    });

    console.log("\nДИЗАЙНЕРЫ:");
    console.log(`- Всего задач в объединенном файле: ${(dfMerged.data || []).filter(isDesigner).length}`);
    console.log(`- Создано в отчетном периоде: ${createdDesign.length}`);
    console.log(`- Выполнено в отчетном периоде: ${completedDesign.length}`);

    console.log("\nТЕКСТОВЫЕ ЗАДАЧИ:");
    console.log(`- Всего задач в объединенном файле: ${(dfMerged.data || []).filter(isTextAuthor).length}`);
    console.log(`- Создано: ${createdText.length}`);
    console.log(`- Выполнено: ${completedText.length}`);

    // === 5. Формирование отчета по дизайнерам ===
    console.log("\n5. ФОРМИРОВАНИЕ ОТЧЕТА ПО ДИЗАЙНЕРАМ:");

    let report = [];

    if (completedDesign.length > 0) {
      console.log(`Найдено ${completedDesign.length} выполненных задач дизайнеров`);

      const reportMap = {};

      completedDesign.forEach((row, index) => {
        const resp = row['Ответственный'] || 'Неизвестно';

        if (index < 5) {
          console.log(`Обработка задачи ${index+1}:`);
          console.log(`  Ответственный: ${resp}`);
          console.log(`  Количество макетов: ${row['Количество макетов']}`);
          console.log(`  Количество вариантов: ${row['Количество предложенных вариантов']}`);
          console.log(`  Оценка работы: ${row['Оценка работы']}`);
        }

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
        const makets = parseInt(row['Количество макетов']) || 0;
        reportMap[resp].Макеты += makets;

        // Обработка вариантов
        const variants = parseInt(row['Количество предложенных вариантов']) || 0;
        reportMap[resp].Варианты += variants;

        // Обработка оценки
        if (row['Оценка работы']) {
          const score = parseFloat(row['Оценка работы']);
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
      
      // Покажем несколько строк для отладки
      console.log("Первые 10 строк данных для анализа:");
      for (let i = 0; i < Math.min(10, dfMerged.data.length); i++) {
        const row = dfMerged.data[i];
        console.log(`Строка ${i+1}:`);
        console.log(`  Ответственный: ${row['Ответственный']}`);
        console.log(`  Дата создания: ${row['Дата создания']}`);
        console.log(`  Дата выполнения: ${row['Выполнена']}`);
        console.log(`  Название: ${row['Название']}`);
      }
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
    console.error("ОШИБКА ПРИ ФОРМИРОВАНИИ ОТЧЕТA:", error.message);
    console.error("Stack trace:", error.stack);
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

    // Читаем первый лист — он содержит данные
    const gridSheet = gridWorkbook.Sheets[gridWorkbook.SheetNames[0]];
    const archiveSheet = archiveWorkbook.Sheets[archiveWorkbook.SheetNames[0]];

    if (!gridSheet || !archiveSheet) {
      throw new Error('Один из листов Excel пуст или не найден');
    }

    // Используем { defval: null } для корректной обработки пустых ячеек
    const allGridRows = xlsx.utils.sheet_to_json(gridSheet, { header: 1, defval: null });
    const allArchiveRows = xlsx.utils.sheet_to_json(archiveSheet, { header: 1, defval: null });

    // Обработка "Грид"
    let gridColumns = [];
    let gridData = [];

    if (allGridRows.length > 0) {
      // Находим первую строку, которая выглядит как заголовок
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
        }).filter(row => Object.keys(row).length > 0); // Удаляем пустые строки
      }
    }

    const dfGrid = {
      columns: gridColumns,
      data: gridData || []
    };

    // Обработка "Архив"
    let archiveColumns = [];
    let archiveData = [];

    if (allArchiveRows.length > 0) {
      // Находим первую строку, которая выглядит как заголовок
      let headerRowIndex = 0;
      for (let i = 0; i < allArchiveRows.length; i++) {
        const row = allGridRows[i]; // ❌ ОШИБКА: должно быть allArchiveRows[i]
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
        }).filter(row => Object.keys(row).length > 0); // Удаляем пустые строки
      }
    }

    const dfArchive = {
      columns: archiveColumns,
       data: archiveData || []
    };

    // Логирование для отладки
    console.log("Архив: колонки =", dfArchive.columns);
    console.log("Архив: количество строк =", (dfArchive.data || []).length);
    console.log("Грид: количество строк =", (dfGrid.data || []).length);

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

    // Отправляем ответ — с защитой от пустого отчёта
    res.json({
      success: true,
      textReport: textReport,
      report: report || [] // на случай, если report undefined
    });

  } catch (error) {
    console.error("❌ Ошибка в /api/upload:", error.message);
    res.status(500).json({ error: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Сервер запущен на порту ${PORT}`);
});
