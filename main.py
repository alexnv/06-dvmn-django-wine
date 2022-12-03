import collections
import os
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


def read_wines_from_excel(filepath):
    categories_and_wines = pandas.read_excel(filepath, sheet_name='Лист1', na_values=['N/A', 'NA'],
                                   keep_default_na=False).to_dict('records')
    categories = collections.defaultdict(list)
    for wine in categories_and_wines:
        categories[wine['Категория']].append({
            'Название': wine['Название'],
            'Сорт': wine['Сорт'],
            'Цена': wine['Цена'],
            'Картинка': wine['Картинка'],
            'Акция': wine['Акция']
        })

    return categories


def calculate_company_age():
    today = datetime.today()
    foundation_year = 1920
    delta_years = today.year - foundation_year
    return delta_years


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.jinja2')

    company_age = calculate_company_age()
    excel_filepath = os.environ['EXCEL_FILE']
    if os.path.isfile(excel_filepath):
        categories = read_wines_from_excel(excel_filepath)
    else:
        raise FileNotFoundError

    rendered_page = template.render(
        company_age=f'{company_age} {get_year_ending(company_age)}',
        categories=categories
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('127.0.0.1', 8080), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
