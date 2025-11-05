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
        
        // ДЕБАГ-логирование для Гнездиловой
        if (row['Ответственный'] && (row['Ответственный'].toString().includes('Гнездилова') || row['Ответственный'].toString().includes('Мария')) && !isNaN(score)) {
          console.log(`✅ Гнездилова - Преобразована оценка: "${row['Оценка работы']}" -> ${score}`);
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
    const gnezdilovaTasks = completedDesign.filter(row => 
      row['Ответственный'] && 
      (row['Ответственный'].toString().includes('Гнездилова') || 
       row['Ответственный'].toString().includes('Мария'))
    );
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
            // Детальный лог для Гнездиловой
            if (resp.includes('Гнездилова') || resp.includes('Мария')) {
              console.log(`✅ ГНЕЗДИЛОВА - Учтена оценка: ${rawScore} -> ${score}`);
            }
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
        if (resp.includes('Гнездилова') || resp.includes('Мария')) {
          console.log(`🎯 ГНЕЗДИЛОВА - оценки=${data.Оценка}, кол-во=${data.count}, средняя=${data.count > 0 ? (data.Оценка / data.count).toFixed(2) : 0}`);
        } else {
          console.log(`${resp}: оценки=${data.Оценка}, кол-во=${data.count}, средняя=${data.count > 0 ? (data.Оценка / data.count).toFixed(2) : 0}`);
        }
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
