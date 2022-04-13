import collections
from collections import defaultdict, Counter

import pandas
from openpyxl import load_workbook


def make_report(log_file_name, report_template_file_name, report_output_file_name):
    # Чтение и анализ данных из Excel
    excel_data = pandas.read_excel(log_file_name, sheet_name='log', engine='openpyxl')
    excel_data_dict = excel_data.to_dict(orient='records')

    # pprint(excel_data_dict)

    visits_dict = defaultdict(int)
    goods_dict = []
    m_goods = []
    f_goods = []
    for element in excel_data_dict:
        # Добавляем элемент в словарь sales_dict
        # element['item'] - название товара
        # Если ключа с таким названием в sales_dict нет, то будет значение 0,
        # таким образом мы просто увеличим его на 1
        visits_dict[element['Браузер']] += 1

        if element['Купленные товары']:
            temp_goods = element['Купленные товары'].split(',')
            for el in temp_goods:
                goods_dict.append(el)

        if element['Пол'] == 'м':
            temp_goods = element['Купленные товары'].split(',')
            for el in temp_goods:
                m_goods.append(el)
        if element['Пол'] == 'ж':
            temp_goods = element['Купленные товары'].split(',')
            for el in temp_goods:
                f_goods.append(el)

    # Ищем самые популярные браузеры и товары и самые непопулярные
    most_popular_browsers = Counter(visits_dict).most_common(7)

    most_popular_goods = collections.Counter(goods_dict).most_common(7)

    most_popular_m_goods = collections.Counter(m_goods).most_common(2)
    temp_less_popular_m_goods = collections.Counter(m_goods).most_common()
    less_popular_m_goods = temp_less_popular_m_goods[:-(len(temp_less_popular_m_goods)+1):-1][0]

    most_popular_f_goods = collections.Counter(f_goods).most_common(2)
    temp_less_popular_f_goods = collections.Counter(f_goods).most_common()
    less_popular_f_goods = temp_less_popular_f_goods[:-(len(temp_less_popular_f_goods)+1):-1][0]

    # Открываем файл шаблона отчета report_template.xlsx
    wb = load_workbook(filename=report_template_file_name)
    ws = wb.active
    ws['A5'] = str(most_popular_browsers[0][0])
    ws['A6'] = str(most_popular_browsers[1][0])
    ws['A7'] = str(most_popular_browsers[2][0])
    ws['A8'] = str(most_popular_browsers[3][0])
    ws['A9'] = str(most_popular_browsers[4][0])
    ws['A10'] = str(most_popular_browsers[5][0])
    ws['A11'] = str(most_popular_browsers[6][0])

    ws['A19'] = str(most_popular_goods[0][0])
    ws['A20'] = str(most_popular_goods[1][0])
    ws['A21'] = str(most_popular_goods[2][0])
    ws['A22'] = str(most_popular_goods[3][0])
    ws['A23'] = str(most_popular_goods[4][0])
    ws['A24'] = str(most_popular_goods[5][0])
    ws['A25'] = str(most_popular_goods[6][0])

    ws['B31'] = str(most_popular_m_goods[0][0])
    ws['B32'] = str(most_popular_f_goods[0][0])
    ws['B33'] = str(less_popular_m_goods[0])
    ws['B34'] = str(less_popular_f_goods[0])

    # Сохраняем файл-отчет
    wb.save(report_output_file_name)
