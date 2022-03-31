import collections
import datetime
from collections import OrderedDict
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

excel_data_df = pd.read_excel(
    io='wine3.xlsx',
    sheet_name="Лист1",
    na_values=['N/A', 'NA'], keep_default_na=False
)
open_table = excel_data_df.to_dict(orient='records')
create_category_wine = collections.defaultdict(list)

for wine in open_table:
    create_category_wine[wine["Категория"]].append(wine)

category_wine = OrderedDict(create_category_wine)
category_wine.move_to_end('Напитки')

now = datetime.datetime.now().year

rendered_page = template.render(wine_information=category_wine, years=f"Уже {now - 1920} года с вами")

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
