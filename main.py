from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas

from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_current_age(reference_point):
    age = datetime.now().year - reference_point
    if (age % 10 == 1 and age % 100 != 11) or age % 100 == 1:
        return f'{age} год'
    elif (age % 100 in (2, 3, 4) or age % 10 in (2, 3, 4)) and age % 100 not in (12, 13, 14):
        return f'{age} года'
    else:
        return f'{age} лет'


def get_wines_description(file_name):
    excel_data_df = pandas.read_excel(file_name, na_values=None, keep_default_na=False)
    data = excel_data_df.to_dict(orient='split')
    wine_description = defaultdict(list)
    for item in data['data']:
        wine_description[item[0]].append(dict(zip(data['columns'], item)))
    return dict(wine_description)


reference_point = 1920
current_age = get_current_age(reference_point)
wine_data = 'wine3.xlsx'
wine_descr = get_wines_description(file_name=wine_data)

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

rendered_page = template.render(current_age=current_age, wine_descr=wine_descr, wine_groups=sorted(wine_descr))

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 9000), SimpleHTTPRequestHandler)
server.serve_forever()
