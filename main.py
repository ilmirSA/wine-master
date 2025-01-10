import argparse
import collections
import datetime
from collections import OrderedDict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape


def sort_table(excel_result):
    products = excel_result
    grouped_wines = collections.defaultdict(list)

    for wine in products:
        grouped_wines[wine["Категория"]].append(wine)

    sorted_wines = OrderedDict(grouped_wines)
    sorted_wines.move_to_end('Напитки')
    return sorted_wines


def get_year(year_opening):
    year_now = datetime.datetime.now().year
    formatted_date = f"Уже {year_now - year_opening} года с вами"
    return formatted_date


def render_page(sorted_wines, year):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    page = template.render(wine_information=sorted_wines, years=year)

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
    parser = argparse.ArgumentParser(
        description="Введите путь до файла эксель:"
    )
    parser.add_argument(
        "--path",
        help="Полный путь до таблицы эксель",
        default="wine.xlsx"
    )
    args = parser.parse_args()
    year_opening = 1920
    formatted_year = get_year(year_opening)
    excel_result = read_excel(args.path)
    sorted_wines = sort_table(excel_result)
    render_page(sorted_wines, formatted_year)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
