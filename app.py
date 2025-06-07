from flask import Flask, request, render_template_string
import openpyxl
import io

app = Flask(__name__)

# HTML форма для загрузки файла
HTML_FORM = '''
<!doctype html>
<title>Обработка Excel файла</title>
<h2>Загрузите Excel (xlsx) файл</h2>
<form method=post enctype=multipart/form-data>
  <input type=file name=file accept=".xlsx">
  <input type=submit value="Загрузить и обработать">
</form>
'''

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file:
            return "Файл не выбран"
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        # Найдем индекс столбца review_moderation_result (по первой строке)
        header = [cell.value for cell in next(ws.iter_rows(max_row=1))]
        try:
            col_idx = header.index('review_moderation_result') + 1  # Нумерация с 1
        except ValueError:
            return "Столбец 'review_moderation_result' не найден"

        # Проходим по строкам и заменяем значения
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            cell = row[0]
            if cell.value == 'Позитивный':
                cell.value = 'ok'
            elif cell.value == 'Негативный':
                cell.value = 'не ok'

        # Преобразуем данные в формат HTML
        from openpyxl.utils import get_column_letter
        table_html = '<table border="1">'
        # Заголовок
        table_html += '<tr>' + ''.join(f'<th>{cell.value}</th>' for cell in ws[1]) + '</tr>'
        # Данные
        for row in ws.iter_rows(min_row=2):
            row_html = ''
            for cell in row:
                # Если это столбец с 'input_productid', сделаем ссылку
                if header[cell.column - 1] == 'input_productid' and cell.value:
                    link = f'https://www.aliexpress.com/item/{cell.value}.html'
                    cell_value = f'<a href="{link}" target="_blank">{cell.value}</a>'
                else:
                    cell_value = '' if cell.value is None else cell.value
                row_html += f'<td>{cell_value}</td>'
            table_html += f'<tr>{row_html}</tr>'
        table_html += '</table>'

        # Вернуть HTML страницу с таблицей
        return render_template_string(f'''
            <!doctype html>
            <html>
            <head>
                <title>Результат обработки</title>
            </head>
            <body>
                <h2>Результат обработки файла</h2>
                {table_html}
                <br><a href="/">Назад</a>
            </body>
            </html>
        ''')

    return render_template_string(HTML_FORM)

if __name__ == '__main__':
    app.run(debug=True)