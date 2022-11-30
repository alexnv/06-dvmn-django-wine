import collections
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_year_ending(year, first="год", second="года", third="лет"):
    ending = year % 100

    if 4 <= ending <= 21:
        return third

    ending = year % 10

    if ending == 1:
        return first
    elif ending in [2, 3, 4]:
        return second
    else:
        return third


def read_wines_from_excel(filename):
    wines_data = pandas.read_excel(filename, sheet_name='Лист1', na_values=['N/A', 'NA'],
                                   keep_default_na=False).to_dict('records')
    categories = collections.defaultdict(list)
    for wine in wines_data:
        categories[wine['Категория']].append({
            'Название': wine['Название'],
            'Сорт': wine['Сорт'],
            'Цена': wine['Цена'],
            'Картинка': wine['Картинка'],
            'Акция': wine['Акция']
        })

    return categories


env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.jinja2')

today = datetime.today()
year_of_foundation = datetime(year=1920, month=1, day=1, hour=0)

delta_years = (today - year_of_foundation).days // 365

wines_data = read_wines_from_excel("wine3.xlsx")

rendered_page = template.render(
    company_age=f'{delta_years} {get_year_ending(delta_years)}',
    categories=wines_data
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('127.0.0.1', 8080), SimpleHTTPRequestHandler)
server.serve_forever()
