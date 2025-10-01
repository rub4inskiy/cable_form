from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file
from datetime import datetime
import os
import pandas as pd
import io

app = Flask(__name__)
app.secret_key = "secret_key"

# Функція для парсингу файлів з папки
def parse_files_from_directory(directory_path):
    """
    Рекурсивно отримує всі файли з усіх підпапок без розширення
    """
    files_list = []
    
    try:
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                # Видаляємо розширення файлу
                filename_without_extension = os.path.splitext(file)[0]
                files_list.append(filename_without_extension)
    except Exception as e:
        print(f"Помилка при парсингу папки: {e}")
    
    # Видаляємо дублікати і сортуємо
    return sorted(list(set(files_list)))

# Функція для завантаження операторів з файлу
def load_operators_from_file(filename="operators.txt"):
    operators = []
    try:
        with open(filename, "r", encoding="utf-8") as f:
            for line in f:
                operator = line.strip()
                if operator:  # Додаємо тільки непорожні рядки
                    operators.append(operator)
    except FileNotFoundError:
        print(f"Файл {filename} не знайдено. Використовуються стандартні оператори.")
        operators = ["Оператор 1", "Оператор 2", "Оператор 3"]
    return operators

# Функція для завантаження ліній з файлу
def load_lines_from_file(filename="lines.txt"):
    lines = []
    try:
        with open(filename, "r", encoding="utf-8") as f:
            for line in f:
                line_name = line.strip()
                if line_name:  # Додаємо тільки непорожні рядки
                    lines.append(line_name)
    except FileNotFoundError:
        print(f"Файл {filename} не знайдено. Використовуються стандартні лінії.")
        lines = ["К-90", "К-80", "FTTH", "SCL", "SZL"]
    return lines

# Завантажуємо дані з файлів
OPERATORS = load_operators_from_file()
LINES = load_lines_from_file()

# Шлях до папки з конструкціями
CONSTRUCTIONS_DIR = r"D:\Workdirectory\БАЗА\Конструкції"

def save_to_excel(data):
    """Функція для збереження даних у Excel"""
    excel_data = {
        'Лінія': [session['line']],
        'Оператор': [session['operator']],
        'Дата запису': [datetime.now().strftime('%d-%m-%Y %H:%M:%S')],
        'Час авторизації': [session['login_time']],
        'Продукція (марка)': [', '.join(data.get('product', ['']))],
        
        # Діаметр оболонки кабелю
        'Діаметр оболонки кабелю': [', '.join(data.get('defect', [])).count('Діаметр оболонки кабелю') > 0],
        'Діаметр - вище допустимого': [', '.join(data.get('diameter_issue', [])) if 'Вище допустимого' in data.get('diameter_issue', []) else ''],
        'Діаметр - вище значення': [', '.join(data.get('diameter_high_value', ['']))],
        'Діаметр - менше допустимого': [', '.join(data.get('diameter_issue', [])) if 'Менше допустимого' in data.get('diameter_issue', []) else ''],
        'Діаметр - менше значення': [', '.join(data.get('diameter_low_value', ['']))],
        
        # Цілісність оболонки кабелю
        'Цілісність оболонки кабелю': [', '.join(data.get('defect', [])).count('Цілісність оболонки кабелю') > 0],
        'Розрив оболонки': ['Так' if 'Розрив оболонки' in data.get('integrity', []) else ''],
        'Розрив оболонки опис': [', '.join(data.get('break_value', ['']))],
        'Наплив маси': ['Так' if 'Наплив маси' in data.get('integrity', []) else ''],
        'Наплив маси опис': [', '.join(data.get('mass_value', ['']))],
        'Нерівномірність діаметру': ['Так' if 'Нерівномірність діаметру' in data.get('integrity', []) else ''],
        'Нерівномірність діаметру опис': [', '.join(data.get('uneven_value', ['']))],
        'Потоншення': ['Так' if 'Потоншення' in data.get('integrity', []) else ''],
        'Потоншення опис': [', '.join(data.get('thin_value', ['']))],
        
        # Інші дефекти
        'Асиметричність': [', '.join(data.get('defect', [])).count('Асиметричність розташування елементів кабелю') > 0],
        'НДВ': [', '.join(data.get('defect', [])).count('НДВ') > 0],
        'НДВ значення': [', '.join(data.get('ndv_value', ['']))],
        'Нерівномірність розкладки': [', '.join(data.get('defect', [])).count('Нерівномірність розкладки кабелю') > 0],
        'Нечіткість маркування': [', '.join(data.get('defect', [])).count('Нечіткість тексту маркування') > 0],
        'Невідповідність маркування': [', '.join(data.get('defect', [])).count('Невідповідність тексту маркування') > 0],
        'Відстань між позначками': [', '.join(data.get('defect', [])).count('Відстань між позначками метражу') > 0],
        'Номер ОВ': [', '.join(data.get('OVnumber_value', ['']))],
        'Відстань значення': [', '.join(data.get('distance_value', ['']))],
        
        # Обрив прутка
        'Обрив прутка': [', '.join(data.get('defect', [])).count('Обрив прутка') > 0],
        'Віддачик №1': ['Так' if 'Віддачик №1' in data.get('rod', []) else ''],
        'Віддачик №1 коментар': [', '.join(data.get('rod1_comment', ['']))],
        'Віддачик №2': ['Так' if 'Віддачик №2' in data.get('rod', []) else ''],
        'Віддачик №2 коментар': [', '.join(data.get('rod2_comment', ['']))],
        'Натяг': ['Так' if 'Натяг, кг' in data.get('rod', []) else ''],
        'Натяг значення': [', '.join(data.get('tension_value', ['']))],
        'Швидкість': ['Так' if 'Швидкість, м/хв' in data.get('rod', []) else ''],
        'Швидкість значення': [', '.join(data.get('speed_value', ['']))],
        
        # Обрив модуля
        'Обрив модуля': [', '.join(data.get('defect', [])).count('Обрив модуля') > 0],
        'Обрив модуля натяг': [', '.join(data.get('module_tension', ['']))]
    }
    
    # Створюємо DataFrame
    df = pd.DataFrame(excel_data)
    
    # Перевіряємо чи існує файл
    filename = 'quality_control_data.xlsx'
    if os.path.exists(filename):
        # Якщо файл існує, додаємо нові дані
        existing_df = pd.read_excel(filename)
        df = pd.concat([existing_df, df], ignore_index=True)
    
    # Зберігаємо у Excel
    df.to_excel(filename, index=False, engine='openpyxl')
    
    return filename

@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        # Перевіряємо, чи це клік по кнопці експорту
        if 'export' in request.form:
            return export_excel()
        
        line = request.form["line"]
        operator = request.form["operator"]
        login_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        session["line"] = line
        session["operator"] = operator
        session["login_time"] = login_time
        return redirect(url_for("form"))
    
    return render_template("login.html", operators=OPERATORS, lines=LINES)

@app.route("/export")
def export_excel():
    """Експорт Excel файлу"""
    filename = 'quality_control_data.xlsx'
    
    if os.path.exists(filename):
        return send_file(
            filename,
            as_attachment=True,
            download_name=f"quality_control_data_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        return "Файл не знайдено", 404

@app.route("/form", methods=["GET", "POST"])
def form():
    if "operator" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        data = request.form.to_dict(flat=False)
        
        # Зберігаємо дані
        filename = save_to_excel(data)
        # Перенаправляємо з параметром для сповіщення
        return redirect(url_for('form') + '?saved=true')
    
    # Отримуємо список продуктів для випадаючого списку
    products = parse_files_from_directory(CONSTRUCTIONS_DIR)
    return render_template("form.html", 
                         operator=session["operator"], 
                         line=session["line"], 
                         login_time=session["login_time"],
                         products=products)

@app.route("/api/search_products")
def search_products():
    """API для пошуку продуктів"""
    products = parse_files_from_directory(CONSTRUCTIONS_DIR)
    return jsonify(products)

@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5005, debug=True)