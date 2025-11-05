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

  // Если строка — попробовать парсить как дату
  if (typeof serial === 'string') {
    const parsed = parseFloat(serial);
    if (!isNaN(parsed)) {
      serial = parsed;
    } else {
      const date = new Date(serial);
      if (!isNaN(date.getTime())) return date;
      return null;
    }
  }

  // Если число — обработать как Excel serial date
  if (typeof serial === 'number') {
    const excelEpochWithError = new Date(1899, 11, 30); // Коррекция Excel bug
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

    // === 1. ОБЪЕДИНЕНИЕ ДАННЫХ ИЗ ГРИДА И АРХИВА ===
    const allData = [...(dfGrid.data || []), ...(dfArchive.data || [])];
    console.log(`Объединено строк: ${allData.length} (Грид: ${dfGrid.data?.length || 0}, Архив: ${dfArchive.data?.length || 0})`);

    // === 2. ПРЕОБРАЗОВАНИЕ ДАТ И ОБРАБОТКА ОТВЕТСТВЕННЫХ ===
    const processedData = allData.map(row => {
      row['Дата создания'] = excelDateToJSDate(row['Дата создания']);
      row['Выполнена'] = excelDateToJSDate(row['Выполнена']);
      
      // Нормализация имени ответственного
      if (!row['Ответственный'] || row['Ответственный'].toString().trim() === '') {
        row['Ответственный'] = 'Неизвестно';
      }
      
      // Нормализация оценки - ИСПРАВЛЕННАЯ ВЕРСИЯ ДЛЯ ТЕКСТА
      if (row['Оценка работы'] !== null && row['Оценка работы'] !== undefined && row['Оценка работы'] !== '') {
        // Преобразуем в строку и очищаем
        let scoreStr = row['Оценка работы'].toString().trim();
        
        // Удаляем все нецифровые символы кроме точки и запятой
        scoreStr = scoreStr.replace(/[^\d,.]/g, '');
        
        // Заменяем запятую на точку (для русских десятичных разделителей)
        scoreStr = scoreStr.replace(',', '.');
        
        const score = parseFloat(scoreStr);
        
        // ДЕБАГ-логирование
        if (!isNaN(score)) {
          console.log(`✅ Преобразована оценка: "${row['Оценка работы']}" -> ${score}`);
        } else {
          console.log(`❌ Не удалось преобразовать: "${row['Оценка работы']}"`);
        }
        
        row['Оценка работы'] = isNaN(score) ? null : score;
      } else {
        row['Оценка работы'] = null;
      }
      
      return row;
    });

    // === 3. ОПРЕДЕЛЕНИЕ ПЕРИОДА ===
    const monthObj = moment(monthName, 'MMMM', true);
    if (!monthObj.isValid()) throw new Error("Неверный месяц");
    const monthNum = monthObj.month() + 1;
    const monthPeriod = `${year}-${monthNum.toString().padStart(2, '0')}`;
    console.log(`Фильтруем по периоду: ${monthPeriod}`);

    // === 4. КЛАССИФИКАЦИЯ ОТВЕТСТВЕННЫХ ===
    const textAuthors = ['Наталия Пятницкая', 'Валентина Кулябина', 'Пятницкая', 'Кулябина'];
    const isTextAuthor = (row) => textAuthors.includes(row['Ответственный']);
    const isDesigner = (row) => !isTextAuthor(row) && row['Ответственный'] !== 'Неизвестно';
    const isUnknown = (row) => row['Ответственный'] === 'Неизвестно';

    // === 5. ПОДСЧЁТ СОЗДАННЫХ И ВЫПОЛНЕННЫХ ЗАДАЧ ===
    const createdDesign = [];
    const completedDesign = [];
    const createdText = [];
    const completedText = [];
    const createdUnknown = [];
    const completedUnknown = [];

    for (const row of processedData) {
      // Дата создания
      const created = row['Дата создания'];
      if (created && moment(created).isValid()) {
        if (moment(created).format('YYYY-MM') === monthPeriod) {
          if (isDesigner(row)) createdDesign.push(row);
          else if (isTextAuthor(row)) createdText.push(row);
          else if (isUnknown(row)) createdUnknown.push(row);
        }
      }

      // Дата выполнения
      const completed = row['Выполнена'];
      if (completed && moment(completed).isValid()) {
        if (moment(completed).format('YYYY-MM') === monthPeriod) {
          if (isDesigner(row)) completedDesign.push(row);
          else if (isTextAuthor(row)) completedText.push(row);
          else if (isUnknown(row)) completedUnknown.push(row);
        }
      }
    }

    console.log("\n📊 СТАТИСТИКА:");
    console.log(`Дизайнеры — создано: ${createdDesign.length}, выполнено: ${completedDesign.length}`);
    console.log(`Текстовые — создано: ${createdText.length}, выполнено: ${completedText.length}`);
    console.log(`Без ответственного — создано: ${createdUnknown.length}, выполнено: ${completedUnknown.length}`);

    // ДЕТАЛЬНАЯ ОТЛАДКА ОЦЕНОК ГНЕЗДИЛОВОЙ
    console.log("\n🔍 ДЕТАЛЬНАЯ ОТЛАДКА ОЦЕНОК ГНЕЗДИЛОВОЙ:");
    const gnezdilovaTasks = completedDesign.filter(row => row['Ответственный'] === 'Мария Гнездилова');
    console.log(`Найдено задач у Гнездиловой: ${gnezdilovaTasks.length}`);
    
    gnezdilovaTasks.forEach((task, index) => {
      console.log(`\nЗадача ${index + 1}: "${task['Название']}"`);
      console.log(`  - Оценка работы: ${task['Оценка работы']}`);
      console.log(`  - Тип оценки: ${typeof task['Оценка работы']}`);
      console.log(`  - Макеты: ${task['Количество макетов']}`);
      console.log(`  - Варианты: ${task['Количество предложенных вариантов']}`);
      console.log(`  - Дата выполнения: ${task['Выполнена']}`);
    });

    // === 6. ФОРМИРОВАНИЕ ОТЧЁТА ПО ВЫПОЛНЕННЫМ ===
    const allCompleted = [...completedDesign, ...completedUnknown];
    let report = [];

    if (allCompleted.length > 0) {
      const reportMap = {};
      
      for (const row of allCompleted) {
        const resp = row['Ответственный'] || 'Неизвестно';
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
        
        // ОБРАБОТКА ОЦЕНКИ С ДЕТАЛЬНОЙ ОТЛАДКОЙ
        const rawScore = row['Оценка работы'];
        if (rawScore !== null && rawScore !== undefined && rawScore !== '') {
          const score = parseFloat(rawScore);
          if (!isNaN(score)) {
            console.log(`✅ Учтена оценка для ${resp}: ${rawScore} -> ${score}`);
            reportMap[resp].Оценка += score;
            reportMap[resp].count += 1;
          } else {
            console.log(`❌ Нечисловая оценка для ${resp}: "${rawScore}" (тип: ${typeof rawScore})`);
          }
        }
      }

      // Вывод статистики по оценкам перед формированием отчета
      console.log("\n📈 СТАТИСТИКА ПО ОЦЕНКАМ:");
      Object.keys(reportMap).forEach(resp => {
        const data = reportMap[resp];
        console.log(`${resp}: оценки=${data.Оценка}, кол-во=${data.count}, средняя=${data.count > 0 ? (data.Оценка / data.count).toFixed(2) : 0}`);
      });

      report = Object.keys(reportMap).map(resp => ({
        Ответственный: resp,
        Задачи: reportMap[resp].Задачи,
        Макеты: reportMap[resp].Макеты,
        Варианты: reportMap[resp].Варианты,
        Оценка: reportMap[resp].count > 0 ? (reportMap[resp].Оценка / reportMap[resp].count).toFixed(2) : 0
      }));
    }

    // Итоговая строка
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

    // === 7. ТЕКСТОВЫЙ ОТЧЁТ ===
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
