from ast import NamedExpr
from collections import namedtuple
from typing import Container
import pythoncom
from win32com.client import Dispatch, gencache
import os
import time

def get_kompas_api():
    API7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0)
    API5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0)

    KompasObject = Dispatch('Kompas.Application.5', None, API5.KompasObject.CLSID)
    app7 = Dispatch('Kompas.Application.7')

    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants

    return API7, API5, app7, KompasObject, const

def get_objects_to_copy(doc):
    global API7
    # Всегда обращаемся к активному виду 
    view = doc.ViewsAndLayersManager.Views.ActiveView
    # Инициализируем контейнер со всеми объектами вида
    container = API7.IDrawingContainer(view)
    # Проходимся по контейнеру и отбираем только объкты-линии
    obj_to_copy = []
    for obj in container.Objects(0):
        # Тип объекта на чертеже не всегда геометрический (напр ILeader)
        try:
            obj_type = obj.DrawingObjectType
        except:
            obj_type = 0

        if obj_type in [1, 2, 3, 8, 31, 32, 33, 34, 35]:
            # Если линия обычная
            if obj.Style == 1:
                # Берем не сам объект, а его reference (long)
                obj_to_copy.append(obj.Reference)
    return obj_to_copy

def copy_to_new_view(doc, doc_app5, obj_to_copy):
    global API7

    # Добавляем новый вид за пределами чертежа
    new_view = doc.ViewsAndLayersManager.Views.Add(1)
    new_view.X = 300
    new_view.Y = 0
    new_view.Update()

    # Создаем новую группу объектов и складываем туда все объекты для копирования
    g = doc_app5.ksNewGroup(1)
    for x in obj_to_copy:
        doc_app5.ksAddObjGroup(g, x)
    # Копируем группу в буфер обмена
    doc_app5.ksWriteGroupToClip(g, 1)
    # Считываем группу из буфера и записываем в вид
    g_ = doc_app5.ksReadGroupFromClip()
    doc_app5.ksStoreTmpGroup(g_)

def destroy_views(doc, doc_app5):
    # Разрушаем все виды, кроме системного
    i = 1
    views =  doc.ViewsAndLayersManager.Views
    # view - это массив и мы итерируемся по нему через метод Items и пердаем индекс из массива
    while views.Item(i):
        doc_app5.ksDestroyObjects(views.Item(i).Reference)
        i += 1

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
        columns = obj.Columns
        name = columns.Column(5, 1, 0).Text.Str
        if name not in details.keys():
            quantity = int(columns.Column(6, 1, 0).Text.Str)
            material = columns.Column(9, 1, 0).Text.Str
            if 'Лист' in material:
                # Листовой материал всегда вида Лист$d1,5 ГОСТ
                thickness = material.split(' ')[0].strip('Лист$d')
                details[name] = [quantity, thickness]
        else:
            details[name][0] += quantity
    return details

try:
    print('Запуск программы...')

    API7, API5, app7, app5, const7 = get_kompas_api()
    # Автоматически перестраиваем все документы
    app7.HideMessage = 1

    spec_doc = app7.ActiveDocument
    # Если активный документ не Спецификация
    if spec_doc.DocumentType != 3:
        raise NameError("ERROR: Спецификация не открыта в Компасе")
    
    print('Укажите путь к чертежам:')
    docs_path = input()

    print('Запуск создания чертежей...')

    new_folder_name = 'DXF {}'.format(spec_doc.Name.strip('.spw'))
    # Если папка под DXF не создана, то создаем 
    if not os.path.exists(docs_path + '\\{}'.format(new_folder_name)):
        os.mkdir(docs_path + '\\{}'.format(new_folder_name))

    try:
        details = get_details_from_spec(spec_doc)
    except:
        raise NameError("EROOR: Не распознаны листовые тела в спецификации")

    for detail_name, (quantity, thickness) in details.items():
        
        try:
            path = docs_path + '\\' + detail_name + '.cdw'
            iKompasDocument = app7.Documents.Open(path, True, False)
            doc = API7.IKompasDocument2D(iKompasDocument)
            doc2d = API7.IKompasDocument2D1(iKompasDocument)
            # Захватываем активный чертеж приложением api5
            doc_app5 = app5.ActiveDocument2D()
        except:
            raise NameError(f'ERROR: Чертеж {detail_name} не был найден.')
        
        try:
            obj_to_copy = get_objects_to_copy(doc)
            copy_to_new_view(doc, doc_app5, obj_to_copy)
            destroy_views(doc, doc_app5)

            if not obj_to_copy:
                raise NameError('ERROR: Не были скопированы объекты для чертежа {}'.format(detail_name))
        except Exception as e:
            raise NameError('ERROR: Ошибка при создании вида для детали {}'.format(detail_name))

        doc2d.RebuildDocument()

        new_path = docs_path + '\\{}\\{} {}шт {}мм.cdw'.format(new_folder_name, detail_name, quantity, thickness)
        flag = doc_app5.ksSaveDocumentEx(new_path, -1)
        if flag:
            print('INFO: Чертеж для детали {} успешно сохранен'.format(detail_name))
        else:
            raise NameError('ERROR: Ошибка при сохранении чертежа для детали {}'.format(detail_name))
        doc_app5.ksCloseDocument()
        doc.Close(0)
    app7.HideMessage = 0
except Exception as e:
    print(f'ERROR: {e}')
    app7.HideMessage = 0

print('Конец программы...')
time.sleep(60)