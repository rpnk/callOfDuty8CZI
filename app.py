import sqlite3
from docx import Document
from flask import Flask, render_template, request, redirect, send_file, jsonify
from io import BytesIO
import re
import logging

# Налаштування логування для дебагу
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Підключення до бази даних
def connect_db():
    conn = sqlite3.connect('military_data.db')
    return conn

# Створення таблиць
def create_tables():
    conn = connect_db()
    cursor = conn.cursor()
    tables = [
        'unit_name',
        'weapon_ammo_duty',
    ]
    for table in tables:
        cursor.execute(f"PRAGMA table_info({table})")
        columns = [col[1] for col in cursor.fetchall()]
        if not columns or 'value' not in columns:
            cursor.execute(f'DROP TABLE IF EXISTS {table}')
            cursor.execute(f'''
                CREATE TABLE {table} (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    value TEXT NOT NULL
                )
            ''')

    cursor.execute("PRAGMA table_info(commander)")
    columns = [col[1] for col in cursor.fetchall()]
    if not columns or 'rank' not in columns or 'name' not in columns or 'unit_id' not in columns:
        cursor.execute('DROP TABLE IF EXISTS commander')
        cursor.execute('''
            CREATE TABLE commander (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rank TEXT NOT NULL,
                name TEXT NOT NULL,
                unit_id INTEGER NOT NULL,
                FOREIGN KEY (unit_id) REFERENCES unit_name(id)
            )
        ''')

    # Таблиця для персоналу (посада, звання, ПІБ) – прибираємо unit_id
    cursor.execute("PRAGMA table_info(personnel)")
    columns = [col[1] for col in cursor.fetchall()]
    if not columns or 'position' not in columns:
        cursor.execute('DROP TABLE IF EXISTS personnel')
        cursor.execute('''
            CREATE TABLE personnel (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                position TEXT NOT NULL,
                rank TEXT NOT NULL,
                name TEXT NOT NULL
            )
        ''')

    # Таблиця для пунктів
    cursor.execute("PRAGMA table_info(items)")
    columns = [col[1] for col in cursor.fetchall()]
    if not columns or 'text' not in columns:
        cursor.execute('DROP TABLE IF EXISTS items')
        cursor.execute('''
            CREATE TABLE items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                text TEXT NOT NULL
            )
        ''')

    # Таблиця для підпунктів (з посиланням на пункт)
    cursor.execute("PRAGMA table_info(sub_items)")
    columns = [col[1] for col in cursor.fetchall()]
    if not columns or 'text' not in columns or 'item_id' not in columns:
        cursor.execute('DROP TABLE IF EXISTS sub_items')
        cursor.execute('''
            CREATE TABLE sub_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                text TEXT NOT NULL,
                item_id INTEGER NOT NULL,
                FOREIGN KEY (item_id) REFERENCES items(id)
            )
        ''')

    # Таблиця для підпунктів підпунктів (з посиланням на підпункт)
    cursor.execute("PRAGMA table_info(sub_sub_items)")
    columns = [col[1] for col in cursor.fetchall()]
    if not columns or 'text' not in columns or 'sub_item_id' not in columns:
        cursor.execute('DROP TABLE IF EXISTS sub_sub_items')
        cursor.execute('''
            CREATE TABLE sub_sub_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                text TEXT NOT NULL,
                sub_item_id INTEGER NOT NULL,
                FOREIGN KEY (sub_item_id) REFERENCES sub_items(id)
            )
        ''')

    # Додаємо тестові дані
    cursor.execute('SELECT COUNT(*) FROM unit_name')
    if cursor.fetchone()[0] == 0:
        cursor.execute('INSERT INTO unit_name (value) VALUES (?)', ('8 ІЗІ та КБ',))
        logger.debug("Додано тестову військову частину: 8 ІЗІ та КБ")

    cursor.execute('SELECT COUNT(*) FROM commander')
    if cursor.fetchone()[0] == 0:
        cursor.execute('INSERT INTO commander (rank, name, unit_id) VALUES (?, ?, ?)', ('ПОЛКОВНИК', 'Плеска А.В.', 1))
        logger.debug("Додано тестового командира: ПОЛКОВНИК Плеска А.В.")

    cursor.execute('SELECT COUNT(*) FROM personnel')
    if cursor.fetchone()[0] == 0:
        cursor.execute('INSERT INTO personnel (position, rank, name) VALUES (?, ?, ?)', ('солдатка', 'рядова', 'Анюка'))
        logger.debug("Додано тестовий персонал: солдатка рядова Анюка")

    cursor.execute('SELECT COUNT(*) FROM items')
    if cursor.fetchone()[0] == 0:
        cursor.execute('INSERT INTO items (text) VALUES (?)', ('Наряди на наряд',))
        logger.debug("Додано тестовий пункт: Наряди на наряд")

    cursor.execute('SELECT COUNT(*) FROM sub_items')
    if cursor.fetchone()[0] == 0:
        cursor.execute('INSERT INTO sub_items (text, item_id) VALUES (?, ?)', ('Начальство', 1))
        logger.debug("Додано тестовий підпункт: Начальство (item_id=1)")

    cursor.execute('SELECT COUNT(*) FROM sub_sub_items')
    if cursor.fetchone()[0] == 0:
        cursor.execute('INSERT INTO sub_sub_items (text, sub_item_id) VALUES (?, ?)', ('Черговий пункту', 1))
        logger.debug("Додано тестовий підпункт підпункту: Черговий пункту (sub_item_id=1)")

    conn.commit()
    conn.close()

# Отримання даних з таблиці
def get_table_data(table_name):
    conn = connect_db()
    cursor = conn.cursor()
    if table_name == 'commander':
        cursor.execute('''
            SELECT commander.id, commander.rank, commander.name, unit_name.value 
            FROM commander 
            JOIN unit_name ON commander.unit_id = unit_name.id
        ''')
    elif table_name == 'personnel':
        cursor.execute('SELECT id, position, rank, name FROM personnel')
    else:
        cursor.execute(f'SELECT id, "value" FROM {table_name}')
    data = cursor.fetchall()
    conn.close()
    logger.debug(f"Дані з таблиці {table_name}: {data}")
    return data

# Отримання пунктів
def get_items():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('SELECT id, text FROM items')
    data = cursor.fetchall()
    conn.close()
    logger.debug(f"Пункти (items): {data}")
    return data

# Отримання підпунктів для певного пункту
def get_sub_items(item_id):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('SELECT id, text FROM sub_items WHERE item_id = ?', (item_id,))
    data = cursor.fetchall()
    conn.close()
    logger.debug(f"Підпункти для item_id={item_id}: {data}")
    return data

# Отримання підпунктів підпунктів для певного підпункту
def get_sub_sub_items(sub_item_id):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('SELECT id, text FROM sub_sub_items WHERE sub_item_id = ?', (sub_item_id,))
    data = cursor.fetchall()
    conn.close()
    logger.debug(f"Підпункти підпунктів для sub_item_id={sub_item_id}: {data}")
    return data

# Отримання значення за ID
def get_value_by_id(table_name, id):
    if not id:
        logger.warning(f"ID is None for table {table_name}")
        return '' if table_name != 'commander' else ('', '', '')
    
    conn = connect_db()
    cursor = conn.cursor()
    if table_name == 'commander':
        cursor.execute('''
            SELECT commander.rank, commander.name, unit_name.value 
            FROM commander 
            JOIN unit_name ON commander.unit_id = unit_name.id 
            WHERE commander.id=?
        ''', (id,))
        result = cursor.fetchone()
        conn.close()
        logger.debug(f"Commander data for id {id}: {result}")
        return result if result else ('', '', '')
    elif table_name == 'personnel':
        cursor.execute('SELECT position, rank, name FROM personnel WHERE id=?', (id,))
        result = cursor.fetchone()
        conn.close()
        logger.debug(f"Personnel data for id {id}: {result}")
        return result if result else ('', '', '')
    elif table_name in ['items', 'sub_items', 'sub_sub_items']:
        cursor.execute(f'SELECT text FROM {table_name} WHERE id=?', (id,))
        result = cursor.fetchone()
        conn.close()
        logger.debug(f"Text for table {table_name}, id {id}: {result}")
        return result[0] if result else ''
    else:
        cursor.execute(f'SELECT "value" FROM {table_name} WHERE id=?', (id,))
        result = cursor.fetchone()
        conn.close()
        logger.debug(f"Value for table {table_name}, id {id}: {result}")
        return result[0] if result else ''

# Додавання або оновлення даних
def upsert_data(table_name, id, value):
    conn = connect_db()
    cursor = conn.cursor()
    if table_name == 'commander':
        rank = value['rank']
        name = value['name']
        unit_id = value['unit_id']
        if id:
            cursor.execute('UPDATE commander SET rank=?, name=?, unit_id=? WHERE id=?', (rank, name, unit_id, id))
        else:
            cursor.execute('INSERT INTO commander (rank, name, unit_id) VALUES (?, ?, ?)', (rank, name, unit_id))
    elif table_name == 'personnel':
        position = value['position']
        rank = value['rank']
        name = value['name']
        if id:
            cursor.execute('UPDATE personnel SET position=?, rank=?, name=? WHERE id=?', (position, rank, name, id))
        else:
            cursor.execute('INSERT INTO personnel (position, rank, name) VALUES (?, ?, ?)', (position, rank, name))
    else:
        if id:
            cursor.execute(f'UPDATE {table_name} SET "value"=? WHERE id=?', (value, id))
        else:
            cursor.execute(f'INSERT INTO {table_name} ("value") VALUES (?)', (value,))
    conn.commit()
    conn.close()

# Додавання нового пункту
def add_item(text):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('INSERT INTO items (text) VALUES (?)', (text,))
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()
    logger.debug(f"Додано новий пункт: {text}, id={new_id}")
    return new_id

# Додавання нового підпункту
def add_sub_item(text, item_id):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('INSERT INTO sub_items (text, item_id) VALUES (?, ?)', (text, item_id))
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()
    logger.debug(f"Додано новий підпункт: {text}, item_id={item_id}, id={new_id}")
    return new_id

# Додавання нового підпункту підпункту
def add_sub_sub_item(text, sub_item_id):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('INSERT INTO sub_sub_items (text, sub_item_id) VALUES (?, ?)', (text, sub_item_id))
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()
    logger.debug(f"Додано новий підпункт підпункту: {text}, sub_item_id={sub_item_id}, id={new_id}")
    return new_id

# Видалення запису
def delete_data(table_name, id):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute(f'DELETE FROM {table_name} WHERE id=?', (id,))
    conn.commit()
    conn.close()

# Перетворення імені таблиці на зрозумілу назву
def table_name_to_display_name(table_name):
    return {
        'unit_name': 'Назва військової частини',
        'weapon_ammo_duty': 'Вид зброї, боєприпаси, особи наряду',
        'commander': 'Звання, ім’я командира та назва частини',
        'personnel': 'Персонал (посада, звання, ПІБ)'
    }[table_name]

# Функція для заміни плейсхолдерів у параграфі
def replace_placeholder_in_paragraph(paragraph, placeholder, replacement):
    full_text = ''.join(run.text for run in paragraph.runs).strip()
    logger.debug(f"Checking placeholder {placeholder} in text: {full_text}")
    if placeholder in full_text:
        new_text = full_text.replace(placeholder, replacement)
        runs = paragraph.runs
        if runs:
            first_run = runs[0]
            font = first_run.font
            bold = first_run.bold
            italic = first_run.italic
            underline = first_run.underline

            paragraph.clear()
            new_run = paragraph.add_run(new_text)
            new_run.bold = bold
            new_run.italic = italic
            new_run.underline = underline
            if font:
                new_run.font.name = font.name
                new_run.font.size = font.size
        logger.debug(f"Replaced {placeholder} with {replacement}")
        return True
    if placeholder == '{date}' and '00.00.0000' in full_text:
        new_text = full_text.replace('00.00.0000', replacement)
        runs = paragraph.runs
        if runs:
            first_run = runs[0]
            font = first_run.font
            bold = first_run.bold
            italic = first_run.italic
            underline = first_run.underline

            paragraph.clear()
            new_run = paragraph.add_run(new_text)
            new_run.bold = bold
            new_run.italic = italic
            new_run.underline = underline
            if font:
                new_run.font.name = font.name
                new_run.font.size = font.size
        logger.debug(f"Replaced 00.00.0000 with {replacement}")
        return True
    return False

# Функція для створення документа
def create_order_document(date, order_number, unit_id, commander_id, items):
    doc = Document('template.docx')

    # Спочатку замінюємо плейсхолдери, які йдуть до {items}
    for para in doc.paragraphs:
        replace_placeholder_in_paragraph(para, '{date}', date)
        replace_placeholder_in_paragraph(para, '{order_number}', order_number)
        unit_name = get_value_by_id('unit_name', unit_id)
        replace_placeholder_in_paragraph(para, '{unit_id}', unit_name)

    # Знаходимо місце для вставки пунктів
    insert_index = None
    for i, para in enumerate(doc.paragraphs):
        if '{items}' in ''.join(run.text for run in para.runs):
            insert_index = i
            break

    if insert_index is None:
        logger.error("Placeholder {items} not found in template")
        return None

    # Видаляємо плейсхолдер {items}
    doc.paragraphs[insert_index].clear()

    # Додаємо пункти та підпункти
    item_number = 1
    for item in items:
        # Додаємо пункт (наприклад, 1.)
        item_text = item['text']
        new_para = doc.add_paragraph(f"{item_number}. {item_text}")
        doc.paragraphs[insert_index]._element.getparent().insert(insert_index, new_para._element)
        insert_index += 1

        # Додаємо підпункти (наприклад, 1.1.)
        sub_items = item.get('sub_items', [])
        sub_item_number = 1
        for sub_item in sub_items:
            sub_item_text = sub_item['text']
            person = sub_item.get('person', '')
            full_text = f"{item_number}.{sub_item_number}. {sub_item_text}"
            if person:
                full_text += f" – {person}"
            new_sub_para = doc.add_paragraph(full_text)
            doc.paragraphs[insert_index]._element.getparent().insert(insert_index, new_sub_para._element)
            insert_index += 1

            # Додаємо підпункти підпунктів (наприклад, 1.1.1.)
            sub_sub_items = sub_item.get('sub_sub_items', [])
            sub_sub_item_number = 1
            for sub_sub_item in sub_sub_items:
                sub_sub_item_text = sub_sub_item['text']
                sub_person = sub_sub_item.get('person', '')
                sub_full_text = f"{item_number}.{sub_item_number}.{sub_sub_item_number}. {sub_sub_item_text}"
                if sub_person:
                    sub_full_text += f" – {sub_person}"
                new_sub_sub_para = doc.add_paragraph(sub_full_text)
                doc.paragraphs[insert_index]._element.getparent().insert(insert_index, new_sub_sub_para._element)
                insert_index += 1
                sub_sub_item_number += 1

            sub_item_number += 1

        item_number += 1

    # Замінюємо плейсхолдери, які йдуть після {items}
    commander_rank, commander_name, commander_unit = get_value_by_id('commander', commander_id)
    commander_full = f"{commander_rank} {commander_name}" if commander_rank and commander_name else ''
    for para in doc.paragraphs:
        replace_placeholder_in_paragraph(para, '{commander}', commander_full)
        replace_placeholder_in_paragraph(para, '{rank}', commander_rank if commander_rank else '')
        replace_placeholder_in_paragraph(para, '{name}', commander_name if commander_name else '')
        replace_placeholder_in_paragraph(para, '{commander_unit}', commander_unit if commander_unit else '')

    # Збереження документа
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Головна сторінка
@app.route('/')
def index():
    tables = [
        ('unit_name', 'Назва військової частини'),
        ('weapon_ammo_duty', 'Вид зброї, боєприпаси, особи наряду'),
        ('commander', 'Звання, ім’я командира та назва частини'),
        ('personnel', 'Персонал (посада, звання, ПІБ)')
    ]
    return render_template('index.html', tables=tables)

# Сторінка таблиці
@app.route('/table/<table_name>', methods=['GET', 'POST'])
def table_view(table_name):
    if request.method == 'POST':
        if table_name == 'commander':
            rank = request.form['rank']
            name = request.form['name']
            unit_id = request.form['unit_id']
            id = request.form.get('id')
            upsert_data(table_name, id, {'rank': rank, 'name': name, 'unit_id': unit_id})
        elif table_name == 'personnel':
            position = request.form['position']
            rank = request.form['rank']
            name = request.form['name']
            id = request.form.get('id')
            upsert_data(table_name, id, {'position': position, 'rank': rank, 'name': name})
        else:
            value = request.form['value']
            id = request.form.get('id')
            upsert_data(table_name, id, value)
        return redirect(f'/table/{table_name}')
    
    data = get_table_data(table_name)
    column_name = table_name_to_display_name(table_name)
    unit_names = get_table_data('unit_name') if table_name == 'commander' else []
    return render_template('table.html', table_name=table_name, data=data, column_name=column_name, unit_names=unit_names)

# Видалення запису
@app.route('/delete/<table_name>/<int:id>')
def delete(table_name, id):
    delete_data(table_name, id)
    return redirect(f'/table/{table_name}')

# Ендпоінт для отримання підпунктів
@app.route('/get_sub_items/<int:item_id>')
def get_sub_items_endpoint(item_id):
    sub_items = get_sub_items(item_id)
    return jsonify(sub_items)

# Ендпоінт для отримання підпунктів підпунктів
@app.route('/get_sub_sub_items/<int:sub_item_id>')
def get_sub_sub_items_endpoint(sub_item_id):
    sub_sub_items = get_sub_sub_items(sub_item_id)
    return jsonify(sub_sub_items)

# Сторінка генерації документа
@app.route('/generate', methods=['GET', 'POST'])
def generate():
    if request.method == 'POST':
        date = request.form['date']
        order_number = request.form['order_number']
        unit_id = request.form['unit_id']
        commander_id = request.form.get('commander_id', '1')

        # Обробка пунктів
        items = []
        item_index = 1
        while f'item-{item_index}-select' in request.form:
            item_select = request.form[f'item-{item_index}-select']
            item_text = request.form.get(f'item-{item_index}-text', '')

            # Якщо вибрано "new", додаємо новий пункт
            if item_select == 'new' and item_text:
                new_item_id = add_item(item_text)
                item_text_to_use = item_text
            else:
                item_text_to_use = get_value_by_id('items', int(item_select) if item_select else None)

            sub_items = []
            sub_item_index = 1
            while f'sub-item-{item_index}-{sub_item_index}-select' in request.form:
                sub_item_select = request.form[f'sub-item-{item_index}-{sub_item_index}-select']
                sub_item_text = request.form.get(f'sub-item-{item_index}-{sub_item_index}-text', '')
                person_id = request.form.get(f'sub-item-{item_index}-{sub_item_index}-person', '')

                # Якщо вибрано "new", додаємо новий підпункт
                if sub_item_select == 'new' and sub_item_text:
                    new_sub_item_id = add_sub_item(sub_item_text, int(item_select) if item_select != 'new' else new_item_id)
                    sub_item_text_to_use = sub_item_text
                else:
                    sub_item_text_to_use = get_value_by_id('sub_items', int(sub_item_select) if sub_item_select else None)

                # Отримуємо повну інформацію про особу
                person_full = ''
                if person_id:
                    position, rank, name = get_value_by_id('personnel', int(person_id))
                    person_full = f"{position} {rank} {name}".strip()

                sub_sub_items = []
                sub_sub_item_index = 1
                while f'sub-sub-item-{item_index}-{sub_item_index}-{sub_sub_item_index}-select' in request.form:
                    sub_sub_item_select = request.form[f'sub-sub-item-{item_index}-{sub_item_index}-{sub_sub_item_index}-select']
                    sub_sub_item_text = request.form.get(f'sub-sub-item-{item_index}-{sub_item_index}-{sub_sub_item_index}-text', '')
                    sub_person_id = request.form.get(f'sub-sub-item-{item_index}-{sub_item_index}-{sub_sub_item_index}-person', '')

                    # Якщо вибрано "new", додаємо новий підпункт підпункту
                    if sub_sub_item_select == 'new' and sub_sub_item_text:
                        new_sub_sub_item_id = add_sub_sub_item(sub_sub_item_text, int(sub_item_select) if sub_item_select != 'new' else new_sub_item_id)
                        sub_sub_item_text_to_use = sub_sub_item_text
                    else:
                        sub_sub_item_text_to_use = get_value_by_id('sub_sub_items', int(sub_sub_item_select) if sub_sub_item_select else None)

                    # Отримуємо повну інформацію про особу для підпункту підпункту
                    sub_person_full = ''
                    if sub_person_id:
                        sub_position, sub_rank, sub_name = get_value_by_id('personnel', int(sub_person_id))
                        sub_person_full = f"{sub_position} {sub_rank} {sub_name}".strip()

                    sub_sub_items.append({
                        'text': sub_sub_item_text_to_use,
                        'person': sub_person_full
                    })
                    sub_sub_item_index += 1

                sub_items.append({
                    'text': sub_item_text_to_use,
                    'person': person_full,
                    'sub_sub_items': sub_sub_items
                })
                sub_item_index += 1

            items.append({
                'text': item_text_to_use,
                'sub_items': sub_items
            })
            item_index += 1

        logger.debug(f"Items: {items}")

        buffer = create_order_document(date, order_number, unit_id, commander_id, items)
        if buffer:
            return send_file(
                buffer,
                as_attachment=True,
                download_name=f"order_{date}.docx",
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            return "Помилка при генерації документа", 500

    unit_data = get_table_data('unit_name')
    personnel_data = get_table_data('personnel')
    commander_data = get_table_data('commander')
    items_data = get_items()
    
    return render_template('generate.html', 
                         unit_data=unit_data,
                         personnel_data=personnel_data,
                         commander_data=commander_data,
                         items_data=items_data)

# Ініціалізація
if __name__ == '__main__':
    create_tables()
    app.run(debug=True, host='0.0.0.0', port=5000)