import collections
import datetime
from collections import OrderedDict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape


def render_page(excel_result, year_now, year_opening):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    products = excel_result
    grouped_wines = collections.defaultdict(list)

    for wine in products:
        grouped_wines[wine["Категория"]].append(wine)

    wine_sorting = OrderedDict(grouped_wines)
    wine_sorting.move_to_end('Напитки')

    page = template.render(wine_information=wine_sorting, years=f"Уже {year_now - year_opening} года с вами")

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(page)


def read_excel(file_path):
    excel = pd.read_excel(
        io=file_path,
        sheet_name="Лист1",
        na_values=['N/A', 'NA'], keep_default_na=False
    ).to_dict(orient='records')
    return excel


def main():
    user_file = input("Введите путь до файла эксель: ")
    file_path = user_file if user_file else "wine.xlsx"
    year_now = datetime.datetime.now().year
    year_opening = 1920
    excel_result = read_excel(file_path)
    render_page(excel_result, year_now, year_opening)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
