from win32com.client import Dispatch, gencache
import os
import time
import pandas as pd

def get_kompas_api():
    API7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0)
    API5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0)

    KompasObject = Dispatch('Kompas.Application.5', None, API5.KompasObject.CLSID)
    app7 = Dispatch('Kompas.Application.7')

    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants

    return API7, API5, app7, KompasObject, const

def get_details_from_spec(doc):
    """Функция итерируется по спецификации и собирает для стандартных и прочих деталей словарь,
        в котором ключ - название детали, а значение - количество
        doc - Объект IKompasDocument, Спецификация
        result - словарь с деталями"""
    # Описание спецификации
    desc = doc.SpecificationDescriptions.Active
    # Все объекты спецификации (детали)
    objects = desc.BaseObjects
    # Инициализируем словарь с деталями
    details = {}
    for obj in objects:
    # Для каждого объекта проверяем его наличие в словаре
    # Если он уже есть, то увеличиваем колличество
        if obj.Section in (25, 30):
            columns = obj.Columns
            name = columns.Column(5, 1, 0).Text.Str
            quantity = int(columns.Column(6, 1, 0).Text.Str)
            if name not in details.keys():
                details[name] = quantity
            else:
                details[name] += quantity
    return details

def create_df_from_specifications(specifications):
    """Функция, которая из спецификаций достает позиции для закопки и складывает в DF"""
    df_result = pd.DataFrame(columns=['Позиция'])
    for file_name in specifications:

        spec_path = docs_path + '\\'  + file_name
        doc = app7.Documents.Open(spec_path, True, False)

        details = get_details_from_spec(doc)
        df = pd.DataFrame.from_dict(details, orient='index').reset_index()
        df.columns = ['Позиция', doc.Name.strip('.spw')]
        print('INFO: Спецификация с названием {} обработана'.format(file_name))
        doc.Close(0)
        df_result = df_result.merge(df, on='Позиция', how='outer')

    return df_result

def save_df_result(df_result):
    df_result.fillna(0, inplace=True)
    df_result['Сумма'] = df_result.sum(axis=1, numeric_only=True)
    columns_order = list(df_result.columns)[:-1]
    columns_order.insert(1,'Сумма')
    df_result = df_result[columns_order]
    df_result.to_excel(result_path + '\\' + 'Заявка.xlsx', index=False)

try:
    module7, _, api7, _, _ = get_kompas_api()
    app7 = api7.Application
    app7.HideMessage = 1

    #doc = app7.ActiveDocument
    print('Укажите расположение спецификаций:')
    docs_path = input()
    print('Укажите, куда сохранить заявку:')
    result_path = input()

    files = os.listdir(docs_path)
    specifications = [file for file in files if file.endswith(".spw")]
    
    if len(specifications) < 1:
        print('ERROR: Спецификации не найдены')

    df_result = create_df_from_specifications(specifications)
    
    app7.HideMessage = 0

    save_df_result(df_result)
    
    print('INFO: Заявка сохранена')
    
except Exception as e:
    print(f'ERROR: {e}')
    app7.HideMessage = 0
print('Конец программы...')
time.sleep(60)