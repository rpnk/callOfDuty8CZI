from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import json
import os
import logging
import re
import io

app = Flask(__name__)

# Настройка логирования
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def sanitize_filename(filename):
    """
    Очищает имя файла от недопустимых символов.
    """
    return re.sub(r'[\\/*?:"<>|]', '_', filename)

# Загрузка JSON-списков (например, для выбора должностей, званий и т.п.)
def load_json(filename):
    data_dir = 'data'
    os.makedirs(data_dir, exist_ok=True)
    try:
        with open(os.path.join(data_dir, f'{filename}.json'), 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        logger.error(f"Ошибка загрузки {filename}.json: {e}")
        return []

# Сохранение JSON-списков
def save_json(filename, data):
    data_dir = 'data'
    os.makedirs(data_dir, exist_ok=True)
    try:
        with open(os.path.join(data_dir, f'{filename}.json'), 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        logger.error(f"Ошибка сохранения {filename}.json: {e}")

def format_punkty(punkty):
    """
    Форматирование структуры пунктов для подстановки в шаблон.
    """
    formatted = []
    for idx, punkt in enumerate(punkty, start=1):
        typ = punkt['type']
        data_dict = punkt['data']
        if typ == 'person':
            formatted.append(f"1.{idx}. {data_dict['посада']}, {data_dict['звання']}, {data_dict['прізвище']}")
        elif typ == 'chergoviy':
            formatted.append(f"1.{idx}. Черговий підрозділ – {data_dict['підраздел']}")
        elif typ == 'narjad':
            formatted.append(f"1.{idx}. Добовий наряд від {data_dict['підраздел']}")
        elif typ == 'roboti':
            formatted.append(f"1.{idx}. Наряд на роботи від {data_dict['підраздел']} ({data_dict['деталі']})")
        elif typ == 'zbroya':
            formatted.append(f"1.{idx}. Зброю і боєприпаси для несення служби видавати: {data_dict['зброя']}, {data_dict['боєприпаси']} боєприпасів, {data_dict['особи']}")
    return "\n".join(formatted)

@app.route('/', methods=['GET', 'POST'])
def index():
    logger.debug("Доступ к маршруту index, метод: %s", request.method)
    if request.method == 'POST':
        logger.debug("Получен POST-запрос: %s", dict(request.form))
        try:
            # Собираем общие данные для замены в документе
            data = {
                'номер_частини': request.form.get('номер_частини', ''),
                'дата_наказу': request.form.get('дата_наказу', ''),
                'місто': request.form.get('місто', ''),
                'номер_наказу': request.form.get('номер_наказу', ''),
                'дата_наряду': request.form.get('дата_наряду', ''),
                'місяць_наряду': request.form.get('місяць_наряду', ''),
                'рік_наряду': request.form.get('рік_наряду', '')
            }
            
            # Собираем пункты (дополнительные данные)
            selected_punkty = []
            # Поиск уникальных индексов пунктов из имен полей (например, "punkt_type_1")
            punkt_indices = [int(key.split('_')[-1]) for key in request.form if key.startswith('punkt_type_')]
            for idx in sorted(punkt_indices):
                typ = request.form.get(f'punkt_type_{idx}')
                logger.debug(f"Обработка пункта {idx}, тип: {typ}")
                if typ == 'person':
                    posada = request.form.get(f'posada_{idx}', '')
                    zvannya = request.form.get(f'zvannya_{idx}', '')
                    prizvyshche = request.form.get(f'prizvyshche_{idx}', '')
                    if posada and zvannya and prizvyshche:
                        selected_punkty.append({
                            'type': typ,
                            'data': {'посада': posada, 'звання': zvannya, 'прізвище': prizvyshche}
                        })
                elif typ in ['chergoviy', 'narjad']:
                    pidrozdil = request.form.get(f'pidrozdil_{idx}', '')
                    if pidrozdil:
                        selected_punkty.append({
                            'type': typ,
                            'data': {'підраздел': pidrozdil}
                        })
                elif typ == 'roboti':
                    pidrozdil = request.form.get(f'pidrozdil_{idx}', '')
                    details = request.form.get(f'details_{idx}', '')
                    if pidrozdil and details:
                        selected_punkty.append({
                            'type': typ,
                            'data': {'підраздел': pidrozdil, 'деталі': details}
                        })
                elif typ == 'zbroya':
                    zbroya = request.form.get(f'zbroya_{idx}', '')
                    boepripasy = request.form.get(f'boepripasy_{idx}', '')
                    osoby = request.form.get(f'osoby_{idx}', '')
                    if zbroya and boepripasy and osoby:
                        selected_punkty.append({
                            'type': typ,
                            'data': {'зброя': zbroya, 'боєприпаси': boepripasy, 'особи': osoby}
                        })
            
            data['пункти'] = format_punkty(selected_punkty) if selected_punkty else 'Немає вибраних пунктів'
            logger.debug("Данные для замены: %s", data)
            
            # Генерация документа с использованием DocxTemplate для рендеринга шаблона
            template_path = 'dodatok-2.docx'
            if not os.path.exists(template_path):
                logger.error("Файл шаблона 'dodatok-2.docx' не найден")
                return "Ошибка: шаблон 'dodatok-2.docx' не найден", 500
            
            doc = DocxTemplate(template_path)
            doc.render(data)
            
            # Формирование имени файла с очисткой недопустимых символов
            raw_filename = f"Наказ_{data['номер_наказу'] or 'без_номера'}_{data['дата_наказу'] or 'без_даты'}.docx"
            output_filename = sanitize_filename(raw_filename)
            
            # Сохранение документа в памяти
            output_stream = io.BytesIO()
            doc.save(output_stream)
            output_stream.seek(0)
            logger.info(f"Документ сгенерирован: {output_filename}")
            return send_file(output_stream, download_name=output_filename, as_attachment=True)
        except Exception as e:
            logger.error(f"Ошибка генерации документа: {e}")
            return f"Ошибка при генерации документа: {str(e)}", 500

    # GET-запрос: загрузка списков для формы
    posady = load_json('posady')
    zvannya = load_json('zvannya')
    pidrozdily = load_json('pidrozdily')
    zbroya = load_json('zbroya')
    return render_template('form.html', posady=posady, zvannya=zvannya, pidrozdily=pidrozdily, zbroya=zbroya)

@app.route('/manage_lists', methods=['GET', 'POST'])
def manage_lists():
    lists = {
        'posady': load_json('posady'),
        'zvannya': load_json('zvannya'),
        'pidrozdily': load_json('pidrozdily'),
        'zbroya': load_json('zbroya')
    }
    
    if request.method == 'POST':
        list_name = request.form.get('list_name')
        action = request.form.get('action')
        value = request.form.get('value')
        index = request.form.get('index')
        
        if list_name in lists:
            try:
                if action == 'add' and value:
                    lists[list_name].append(value)
                elif action == 'edit' and value and index:
                    lists[list_name][int(index)] = value
                elif action == 'delete' and index:
                    lists[list_name].pop(int(index))
                save_json(list_name, lists[list_name])
                lists[list_name] = load_json(list_name)
            except Exception as ex:
                logger.error(f"Ошибка при обработке списка {list_name}: {ex}")
    
    return render_template('manage_lists.html', lists=lists)

if __name__ == '__main__':
    app.run(debug=True)
