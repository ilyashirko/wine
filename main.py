from datetime import datetime as dt
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

YEAR_DECLINATION = {
    "1": "год",
    "2": "года",
    "3": "года",
    "4": "года",
    "5": "лет",
    "6": "лет",
    "7": "лет",
    "8": "лет",
    "9": "лет",
    "0": "лет"
}


def company_age():
    year = str(dt.now().year - 1920)
    return f'{year} {YEAR_DECLINATION[year[len(year)-1]]}'


def excel_refactor(file_name):
    file_data = pandas.read_excel(
        file_name, 
        na_values=["NaN", "nan"], 
        keep_default_na=False
    ).to_dict("records")

    file_data_ref = {}
    [file_data_ref.setdefault(drink_info["Категория"], []).append(drink_info) for drink_info in file_data]
    return dict(sorted(file_data_ref.items()))


if __name__ == '__main__':
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        company_age=company_age(),
        wines=excel_refactor("wine3.xlsx")
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
