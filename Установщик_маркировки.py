from win32com.client import Dispatch, gencache
import time
import re
from decimal import Decimal
from pythoncom import VT_EMPTY
from win32com.client import VARIANT

def get_kompas_api():
    API7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0)
    API5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0)

    KompasObject = Dispatch('Kompas.Application.5', None, API5.KompasObject.CLSID)
    app7 = Dispatch('Kompas.Application.7')

    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    
    return API7, API5, app7, KompasObject, const

def get_base_objects(part):
    """Функция возвращает список объектов деталей и список объектов подсборок"""
    global API7, iPropertyMng, _doc
    parts = {}
    assemblies = {}

    for item in part.PartsEx(0):
        # Получаем свойство Раздел спецификации
        properties = API7.IPropertyKeeper(item)
        # Вызываем указатель на свойство объекта в сборке (см Редактор свойст)
        property_marking = iPropertyMng.GetProperty(_doc, 15) # 15 - Раздел спецификации
        spec_section = properties.GetPropertyValue(property_marking, 0, True)[1]
        if spec_section == 'Детали':
            if item.Name not in parts.keys():
                parts[item.Name] = item
        elif spec_section == 'Сборочные единицы':
            if item.Name not in assemblies.keys():
                assemblies[item.Name] = item
    # Сортируем детали по имени 
    parts = sorted(parts.values(), key=lambda x: x.Name)
    assemblies = sorted(assemblies.values(), key=lambda x: x.Name)
    return parts, assemblies

def change_marking(item, marking):
    """Функция, которая открывает документ источник для детали,
        меняет маркировку и сохраняет файл."""
    global app7, API7, iPropertyMng, _doc
    # Интерфейс коллекции документов
    iDocuments = app7.Documents
    # Получаем путь к файлу исходнику
    part_path = item.FileName.replace('>', '')
    # Открываем файл, получаем верхнюю деталь (сам объект)
    iDoc = iDocuments.Open(part_path, False, False)
    iKompasDocument3Ddoc = API7.IKompasDocument3D(iDoc)
    iPart7doc = iKompasDocument3Ddoc.TopPart
    # Меняем обозначение, перестраиваем и закрываем
    iPart7doc.Marking = marking
    iPart7doc.Update()
    iDoc.Save()
    iDoc.Close(1)

    # Свойства объекта в сборке
    properties = API7.IPropertyKeeper(item)
    # Вызываем указатель на свойство объекта в сборке (см Редактор свойст)
    property_marking = iPropertyMng.GetProperty(_doc, 0) # 0 - Обозначение
    # В активной сборке устанавливаем свойство для маркировки
    # Маркировка строится по Источнику (передаем специальное значение в свойство)
    properties.SetPropertyValue(property_marking, VARIANT(VT_EMPTY, None), True)
    item.Update()

try:
    API7, API5, app7, _, _ = get_kompas_api()
    # Автоматически перестраиваем все документы
    app7.HideMessage = 1
    iPropertyMng = API7.IPropertyMng(app7)

    _doc = app7.ActiveDocument
    doc = API7.IKompasDocument3D(_doc)
    # Получаем объект сборки
    model = doc.TopPart

    print('Введите обозначение:')
    top_marking = str(input())

    patern = r'[v/]\d+\.[\d-]*'

    match = re.findall(patern, top_marking)

    # Если в маркировке найдено больше одного паттерна, то просим уточнить
    if len(match) == 1:
        assembly_marking = match[0]
        assembly_marking = assembly_marking.strip('v/-')
    else:
        print('INFO: Цифровая часть в маркировке не расспознана. Введите кодировку для подсборок:')
        assembly_marking = str(input())
        assembly_marking = assembly_marking.strip('v/-')
    # Если ввели все равно енверно, то выдаем ошибку
    if assembly_marking in top_marking:
        rest_marking = top_marking.split(assembly_marking)

        assembly_cnt = Decimal(assembly_marking)
        detail_cnt = Decimal("0")
    else:
        print('ERROR: Кодировка не найдена в маркировке.')

    parts, assemblies = get_base_objects(model)
    
    # for assembly in assemblies:
    #     assembly_cnt += Decimal("0.01")
    #     marking = str(float(assembly_cnt)).join(rest_marking)
    #     # Устанавливаем обозначение
    #     try:
    #         change_marking(assembly, marking)
    #         print(f'INFO: Установлено обозначение {marking} для сборки {assembly.Name}')
    #     except:
    #         print(f'ERROR: Не удалось установить обозначение {marking} для сборки {assembly.Name}')
    
    counter_info = 0
    counter_error = 0
    for part in parts:
        detail_cnt += Decimal("0.01")
        marking = top_marking + str(detail_cnt)[1:]
        # Устанавливаем обозначение
        try:
            change_marking(part, marking)
            counter_info += 1
            print(f'INFO: Установлено обозначение {marking} для детали {part.Name}')
        except:
            counter_error += 1
            print(f'ERROR: Не удалось установить обозначение {marking} для детали {part.Name}')
    
    print(f'INFO: Обозначение было установленно для {counter_info} деталей {counter_error}')

    # Меняем обозначение у самой сборки
    model.Marking = top_marking
    model.Update()
    # Перестраиваем активную сборку 
    doc.RebuildDocument()
    _doc.Save()
    app7.HideMessage = 0
    #app7.Quit()
except Exception as e:
    print(f'ERROR: {e}')
    app7.HideMessage = 0
    #app7.Quit()

print('Конец программы...')
time.sleep(60)
