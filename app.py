# app.py - основной файл Flask приложения
import logging
import pandas as pd
from datetime import datetime
import tempfile
import os
import glob
import shutil
from flask import Flask, request, jsonify, send_file, render_template_string
from werkzeug.utils import secure_filename
from io import BytesIO
import uuid

# === Настройки ===
MARKETPLACE_PATH = r"V:\Workspace\РЕКЛАМНЫЕ МАКЕТЫ\Маркетплейсы"

# === Логирование ===
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-secret-key')

# === HTML шаблон для интерфейса ===
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Генератор отчетов для Kaiten</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h1 {
            color: #2c3e50;
            text-align: center;
            margin-bottom: 30px;
        }
        .upload-section {
            margin-bottom: 30px;
            padding: 20px;
            border: 2px dashed #ddd;
            border-radius: 8px;
            text-align: center;
        }
        .form-group {
            margin-bottom: 20px;
        }
        label {
            display: block;
            margin-bottom: 8px;
            font-weight: bold;
            color: #34495e;
        }
        select, input[type="number"] {
            width: 100%;
            padding: 12px;
            border: 2px solid #ddd;
            border-radius: 6px;
            font-size: 16px;
            box-sizing: border-box;
        }
        select:focus, input[type="number"]:focus {
            border-color: #3498db;
            outline: none;
        }
        .btn {
            background: #3498db;
            color: white;
            padding: 15px 30px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 16px;
            font-weight: bold;
            width: 100%;
            transition: background 0.3s;
        }
        .btn:hover {
            background: #2980b9;
        }
        .btn:disabled {
            background: #bdc3c7;
            cursor: not-allowed;
        }
        .file-info {
            margin-top: 10px;
            padding: 10px;
            background: #ecf0f1;
            border-radius: 4px;
            font-size: 14px;
        }
        .success {
            color: #27ae60;
            font-weight: bold;
        }
        .error {
            color: #e74c3c;
            font-weight: bold;
        }
        .download-section {
            margin-top: 30px;
            padding: 20px;
            background: #ecf0f1;
            border-radius: 8px;
            text-align: center;
        }
        .download-btn {
            background: #27ae60;
            margin: 10px;
            padding: 12px 24px;
            display: inline-block;
        }
        .download-btn:hover {
            background: #229954;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Генератор отчетов</h1>
        
        <div class="upload-section">
            <h3>📁 Загрузка файлов</h3>
            
            <div class="form-group">
                <label for="grid_file">Файл Грид (XLSX):</label>
                <input type="file" id="grid_file" accept=".xlsx" onchange="handleFileSelect('grid')">
                <div id="grid_info" class="file-info"></div>
            </div>
            
            <div class="form-group">
                <label for="archive_file">Файл Архив (XLSX):</label>
                <input type="file" id="archive_file" accept=".xlsx" onchange="handleFileSelect('archive')">
                <div id="archive_info" class="file-info"></div>
            </div>
        </div>

        <div class="form-group">
            <label for="month">Месяц:</label>
            <select id="month">
                <option value="1">Январь</option>
                <option value="2">Февраль</option>
                <option value="3">Март</option>
                <option value="4">Апрель</option>
                <option value="5">Май</option>
                <option value="6">Июнь</option>
                <option value="7">Июль</option>
                <option value="8">Август</option>
                <option value="9">Сентябрь</option>
                <option value="10">Октябрь</option>
                <option value="11">Ноябрь</option>
                <option value="12">Декабрь</option>
            </select>
        </div>

        <div class="form-group">
            <label for="year">Год:</label>
            <input type="number" id="year" min="2020" max="2030" value="{{ current_year }}">
        </div>

        <button class="btn" onclick="generateReport()" id="generate_btn">Сгенерировать отчет</button>

        <div id="result_section" class="download-section" style="display: none;">
            <h3>📄 Результаты</h3>
            <div id="text_report"></div>
            <div>
                <button class="btn download-btn" onclick="downloadFile('excel')">📊 Скачать Excel отчет</button>
                <button class="btn download-btn" onclick="downloadFile('text')">📝 Скачать текстовый отчет</button>
                <button class="btn download-btn" onclick="downloadFile('merged')">🔄 Скачать объединенный файл</button>
            </div>
        </div>
    </div>

    <script>
        let sessionId = null;
        let files = { grid: null, archive: null };

        function handleFileSelect(type) {
            const fileInput = document.getElementById(type + '_file');
            const fileInfo = document.getElementById(type + '_info');
            
            if (fileInput.files.length > 0) {
                const file = fileInput.files[0];
                files[type] = file;
                fileInfo.innerHTML = `<span class="success">✓ ${file.name} (${formatFileSize(file.size)})</span>`;
            } else {
                files[type] = null;
                fileInfo.innerHTML = '<span class="error">✗ Файл не выбран</span>';
            }
        }

        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }

        async function generateReport() {
            const btn = document.getElementById('generate_btn');
            const resultSection = document.getElementById('result_section');
            const textReport = document.getElementById('text_report');
            
            if (!files.grid || !files.archive) {
                alert('Пожалуйста, загрузите оба файла');
                return;
            }

            const month = document.getElementById('month').value;
            const year = document.getElementById('year').value;

            btn.disabled = true;
            btn.textContent = 'Генерация...';

            const formData = new FormData();
            formData.append('grid_file', files.grid);
            formData.append('archive_file', files.archive);
            formData.append('month', month);
            formData.append('year', year);

            try {
                const response = await fetch('/generate_report', {
                    method: 'POST',
                    body: formData
                });

                const data = await response.json();

                if (data.success) {
                    sessionId = data.session_id;
                    textReport.innerHTML = `<pre>${data.text_report}</pre>`;
                    resultSection.style.display = 'block';
                } else {
                    alert('Ошибка: ' + data.error);
                }
            } catch (error) {
                alert('Ошибка при генерации отчета: ' + error);
            } finally {
                btn.disabled = false;
                btn.textContent = 'Сгенерировать отчет';
            }
        }

        function downloadFile(type) {
            if (!sessionId) {
                alert('Сначала сгенерируйте отчет');
                return;
            }
            
            let filename = '';
            switch(type) {
                case 'excel':
                    filename = `Отчет_${document.getElementById('month').options[document.getElementById('month').selectedIndex].text}_${document.getElementById('year').value}.xlsx`;
                    break;
                case 'text':
                    filename = `Статистика_${document.getElementById('month').options[document.getElementById('month').selectedIndex].text}_${document.getElementById('year').value}.txt`;
                    break;
                case 'merged':
                    filename = `Объединенный_файл_${document.getElementById('month').options[document.getElementById('month').selectedIndex].text}_${document.getElementById('year').value}.xlsx`;
                    break;
            }
            
            window.open(`/download/${type}/${sessionId}?filename=${encodeURIComponent(filename)}`, '_blank');
        }
    </script>
</body>
</html>
"""

# === Вспомогательные функции ===
def count_marketplace_cards(month_num, year):
    """Считает количество карточек маркетплейсов за указанный месяц"""
    try:
        # Для демонстрации возвращаем фиктивное значение
        # В реальном использовании нужно настроить путь к файлам
        return 25
    except Exception as e:
        logger.error(f"Ошибка при подсчете карточек маркетплейсов: {e}")
        return 0

def merge_grid_to_archive(df_grid, df_archive):
    """Объединяет данные из Грида в Архив, добавляя новые строки с совпадающими колонками"""
    try:
        logger.info("=== НАЧАЛО ОБЪЕДИНЕНИЯ ДАННЫХ ===")

        # Находим общие колонки
        common_columns = list(set(df_grid.columns) & set(df_archive.columns))
        logger.info(f"Общие колонки для объединения: {common_columns}")

        if not common_columns:
            logger.error("Нет общих колонок для объединения!")
            return df_archive

        # Создаем новый DataFrame только с общими колонками из Грида
        df_grid_common = df_grid[common_columns].copy()

        # Добавляем данные из Грида в Архив
        df_merged = pd.concat([df_archive, df_grid_common], ignore_index=True)
        logger.info(
            f"Объединено данных: Грид={len(df_grid_common)} строк + Архив={len(df_archive)} строк = Итого={len(df_merged)} строк")

        return df_merged

    except Exception as e:
        logger.error(f"Ошибка при объединении данных: {e}")
        raise

def generate_report(df_grid, df_archive, month, year):
    try:
        logger.info("=== НАЧАЛО ФОРМИРОВАНИЯ ОТЧЕТА ===")
        logger.info(f"Параметры: месяц={month}, год={year}")

        # === 1. Объединение данных ===
        logger.info("\n1. ОБЪЕДИНЕНИЕ ДАННЫХ ИЗ ГРИДА И АРХИВА")
        df_merged = merge_grid_to_archive(df_grid, df_archive)

        # Заменяем пустые значения в Ответственном на "Неизвестно"
        if 'Ответственный' in df_merged.columns:
            df_merged['Ответственный'] = df_merged['Ответственный'].fillna('Неизвестно')

        # === 2. Преобразование дат ===
        logger.info("\n2. ПРЕОБРАЗОВАНИЕ ДАТ:")
        date_columns = ['Дата создания', 'Выполнена']
        for col in date_columns:
            if col in df_merged.columns:
                df_merged[col] = pd.to_datetime(df_merged[col], errors='coerce')
        
        logger.info(f"Преобразовано дат: {len(df_merged)} строк")

        # === 3. Определение месяца ===
        logger.info("\n3. ОПРЕДЕЛЕНИЕ МЕСЯЦА ОТЧЕТА:")
        month_num = int(month)
        month_names = {
            1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель',
            5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август',
            9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'
        }
        month_name = month_names.get(month_num, f"Месяц {month_num}")
        logger.info(f"Номер месяца: {month_num} -> название: {month_name}")

        month_str = f"{year}-{month_num:02d}"
        start_date = pd.Timestamp(f"{year}-{month_num:02d}-01")
        if month_num == 12:
            end_date = pd.Timestamp(f"{year+1}-01-01")
        else:
            end_date = pd.Timestamp(f"{year}-{month_num+1:02d}-01")

        logger.info(f"Период для фильтрации: с {start_date} по {end_date}")

        # === 4. Подсчет статистики ===
        logger.info("\n4. ПОДСЧЕТ СТАТИСТИКИ:")
        text_authors = ['Наталия Пятницкая', 'Валентина Кулябина', 'Пятницкая', 'Кулябина']

        # Определяем типы задач
        if 'Ответственный' in df_merged.columns:
            is_text_author = df_merged['Ответственный'].isin(text_authors)
            is_designer = ~is_text_author
        else:
            is_text_author = pd.Series([False] * len(df_merged))
            is_designer = pd.Series([True] * len(df_merged))

        # Фильтруем задачи по периоду
        if 'Дата создания' in df_merged.columns:
            created_mask = (df_merged['Дата создания'] >= start_date) & (df_merged['Дата создания'] < end_date)
            created_design = df_merged[is_designer & created_mask]
            created_text = df_merged[is_text_author & created_mask]
        else:
            created_design = pd.DataFrame()
            created_text = pd.DataFrame()

        if 'Выполнена' in df_merged.columns:
            completed_mask = (df_merged['Выполнена'] >= start_date) & (df_merged['Выполнена'] < end_date)
            completed_design = df_merged[is_designer & completed_mask]
            completed_text = df_merged[is_text_author & completed_mask]
        else:
            completed_design = pd.DataFrame()
            completed_text = pd.DataFrame()

        # === Подсчет задач без ответственного ===
        no_resp_mask = df_merged['Ответственный'].isna() if 'Ответственный' in df_merged.columns else pd.Series([False] * len(df_merged))

        # Поступившие без ответственного
        if 'Дата создания' in df_merged.columns:
            no_resp_created = df_merged[no_resp_mask & created_mask]
        else:
            no_resp_created = pd.DataFrame()

        # Завершенные без ответственного
        if 'Выполнена' in df_merged.columns:
            no_resp_completed = df_merged[no_resp_mask & completed_mask]
        else:
            no_resp_completed = pd.DataFrame()

        # Подсчет макетов и вариантов
        def sum_column(df, col):
            return df[col].sum() if col in df.columns else 0

        no_resp_created_makets = sum_column(no_resp_created, 'Количество макетов')
        no_resp_created_variants = sum_column(no_resp_created, 'Количество предложенных вариантов')
        no_resp_completed_makets = sum_column(no_resp_completed, 'Количество макетов')
        no_resp_completed_variants = sum_column(no_resp_completed, 'Количество предложенных вариантов')

        # === 5. Формирование отчета по дизайнерам ===
        logger.info("\n5. ФОРМИРОВАНИЕ ОТЧЕТА ПО ДИЗАЙНЕРАМ:")
        if not completed_design.empty and 'Ответственный' in completed_design.columns:
            # Стандартизируем названия колонок
            col_rename = {
                'Кол-во макетов': 'Количество макетов',
                'Кол-во вариантов': 'Количество предложенных вариантов'
            }
            completed_design = completed_design.rename(columns=col_rename)

            # Заполняем пропущенные значения
            if 'Количество макетов' not in completed_design.columns:
                completed_design['Количество макетов'] = 0
            if 'Количество предложенных вариантов' not in completed_design.columns:
                completed_design['Количество предложенных вариантов'] = 0

            completed_design['Количество макетов'] = completed_design['Количество макетов'].fillna(0).astype(int)
            completed_design['Количество предложенных вариантов'] = completed_design['Количество предложенных вариантов'].fillna(0).astype(int)

            # Группируем данные
            agg_dict = {
                'Задачи': ('Ответственный', 'size'),
                'Макеты': ('Количество макетов', 'sum'),
                'Варианты': ('Количество предложенных вариантов', 'sum')
            }
            
            if 'Оценка работы' in completed_design.columns:
                agg_dict['Оценка'] = ('Оценка работы', 'mean')

            report = completed_design.groupby('Ответственный').agg(**agg_dict).reset_index()

            # Добавляем строку для "Неизвестно" если есть такие задачи
            unknown_tasks = completed_design[completed_design['Ответственный'] == 'Неизвестно']
            if not unknown_tasks.empty:
                unknown_row = {
                    'Ответственный': 'Неизвестно',
                    'Задачи': len(unknown_tasks),
                    'Макеты': unknown_tasks['Количество макетов'].sum(),
                    'Варианты': unknown_tasks['Количество предложенных вариантов'].sum()
                }
                if 'Оценка работы' in unknown_tasks.columns:
                    unknown_row['Оценка'] = round(unknown_tasks['Оценка работы'].mean(), 2)
                
                report = pd.concat([report, pd.DataFrame([unknown_row])], ignore_index=True)

            logger.info("Отчет после группировки:")
            logger.info(report.to_string())
        else:
            report = pd.DataFrame(columns=['Ответственный', 'Задачи', 'Макеты', 'Варианты'])
            logger.warning("Нет выполненных задач дизайнеров для отчетного периода")

        # Добавляем итоговую строку
        if not report.empty:
            total_row = {
                'Ответственный': 'ИТОГО',
                'Задачи': report['Задачи'].sum(),
                'Макеты': report['Макеты'].sum(),
                'Варианты': report['Варианты'].sum()
            }
            if 'Оценка' in report.columns:
                total_row['Оценка'] = round(report['Оценка'].mean(), 2)
            
            report = pd.concat([report, pd.DataFrame([total_row])], ignore_index=True)

        # === 6. Формирование текстового отчета ===
        text_report = f"""ОТЧЕТ ЗА {month_name.upper()} {year} ГОДА

Дизайнеры:
- Поступило задач: {len(created_design)}
- Выполнено задач: {len(completed_design)}

Текстовые задачи:
- Поступило: {len(created_text)}
- Выполнено: {len(completed_text)}

Задачи без ответственного (поступившие):
- Задач: {len(no_resp_created)}
- Макетов: {int(no_resp_created_makets)}
- Вариантов: {int(no_resp_created_variants)}

Задачи без ответственного (завершенные):
- Задач: {len(no_resp_completed)}
- Макетов: {int(no_resp_completed_makets)}
- Вариантов: {int(no_resp_completed_variants)}

СТАТИСТИКА ПО ВЫПОЛНЕННЫМ ЗАДАЧАМ ДИЗАЙНЕРОВ:
(только задачи, завершенные в отчетном периоде)"""

        logger.info("\n=== ОТЧЕТ УСПЕШНО СФОРМИРОВАН ===")
        return report, text_report

    except Exception as e:
        logger.error(f"ОШИБКА ПРИ ФОРМИРОВАНИИ ОТЧЕТА: {str(e)}", exc_info=True)
        raise

# === Сессии для хранения временных файлов ===
sessions = {}

# === Маршруты Flask ===
@app.route('/kaiten-addon')
def serve_kaiten_addon():
    """Отдает HTML-файл для Kaiten Addon"""
    return send_file('kaiten-addon.html')
@app.route('/')
def index():
    current_year = datetime.now().year
    return render_template_string(HTML_TEMPLATE, current_year=current_year)
@app.route('/test')
def test():
    return "Hello from Flask!"
@app.route('/generate_report', methods=['POST'])
def generate_report_route():
    try:
        if 'grid_file' not in request.files or 'archive_file' not in request.files:
            return jsonify({'success': False, 'error': 'Необходимо загрузить оба файла'})
        
        grid_file = request.files['grid_file']
        archive_file = request.files['archive_file']
        month = request.form.get('month')
        year = request.form.get('year')

        if not month or not year:
            return jsonify({'success': False, 'error': 'Необходимо указать месяц и год'})

        if grid_file.filename == '' or archive_file.filename == '':
            return jsonify({'success': False, 'error': 'Файлы не выбраны'})

        # Чтение файлов
        df_grid = pd.read_excel(grid_file)
        df_archive = pd.read_excel(archive_file)

        # Генерация отчета
        report_df, text_report = generate_report(df_grid, df_archive, month, int(year))

        # Создание временных файлов
        session_id = str(uuid.uuid4())
        temp_dir = tempfile.mkdtemp()
        
        excel_path = os.path.join(temp_dir, 'report.xlsx')
        text_path = os.path.join(temp_dir, 'report.txt')
        merged_path = os.path.join(temp_dir, 'merged.xlsx')

        # Сохранение файлов
        report_df.to_excel(excel_path, index=False)
        
        with open(text_path, 'w', encoding='utf-8') as f:
            f.write(text_report)

        df_merged = merge_grid_to_archive(df_grid, df_archive)
        df_merged.to_excel(merged_path, index=False)

        # Сохранение в сессии
        sessions[session_id] = {
            'excel_path': excel_path,
            'text_path': text_path,
            'merged_path': merged_path,
            'temp_dir': temp_dir,
            'text_report': text_report
        }

        return jsonify({
            'success': True,
            'session_id': session_id,
            'text_report': text_report
        })

    except Exception as e:
        logger.error(f"Ошибка при генерации отчета: {e}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)})

@app.route('/download/<file_type>/<session_id>')
def download_file(file_type, session_id):
    try:
        if session_id not in sessions:
            return "Сессия не найдена", 404

        session_data = sessions[session_id]
        filename = request.args.get('filename', 'file')

        if file_type == 'excel':
            return send_file(session_data['excel_path'], as_attachment=True, download_name=filename)
        elif file_type == 'text':
            return send_file(session_data['text_path'], as_attachment=True, download_name=filename)
        elif file_type == 'merged':
            return send_file(session_data['merged_path'], as_attachment=True, download_name=filename)
        else:
            return "Неверный тип файла", 400

    except Exception as e:
        logger.error(f"Ошибка при скачивании файла: {e}")
        return "Ошибка при скачивании файла", 500

@app.route('/cleanup/<session_id>')
def cleanup(session_id):
    """Очистка временных файлов"""
    try:
        if session_id in sessions:
            session_data = sessions.pop(session_id)
            shutil.rmtree(session_data['temp_dir'])
            return jsonify({'success': True})
        return jsonify({'success': False, 'error': 'Сессия не найдена'})
    except Exception as e:
        logger.error(f"Ошибка при очистке: {e}")
        return jsonify({'success': False, 'error': str(e)})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
