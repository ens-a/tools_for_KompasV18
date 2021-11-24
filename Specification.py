from typing import Container
from win32com.client import Dispatch, gencache
import os
import pythoncom

def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const

module7, api7, const7 = get_kompas_api7()
app7 = api7.Application
#app7.Visible = True
app7.HideMessage = const7.ksHideMessageNo

def get_details_from_spec(doc):
    """Функция итерируется по спецификации и собирает для листовых деталей словарь,
        в котором ключ - название детали, а значение - список вида [количество, толщина листа]
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
        print(obj.Section)
        columns = obj.Columns
        name = columns.Column(5, 1, 0).Text.Str
        if name not in details.keys():
            quantity = int(columns.Column(6, 1, 0).Text.Str)
            material = columns.Column(9, 1, 0).Text.Str
            if 'Лист' in material:
                # Листовой материал всегда вида Лист$d1,5 ГОСТ
                thickness = float(material[6:9].replace(',', '.'))
                details[name] = [quantity, thickness]
        else:
            details[name][0] += quantity
    return details

#doc = app7.ActiveDocument

# spec_path = input()
spec_path = r'C:\Users\UserPC\Documents\Macro Kompas\Спецификация.spw'
doc = app7.Documents.Open(spec_path, True, False)

details = get_details_from_spec(doc)
print(details)
#iKompasDocument = app7.ActiveDocument
#path = iKompasDocument.PathName
#print(iKompasDocument)

app7.Quit()