from typing import Container
import pythoncom
from win32com.client import Dispatch, gencache
import os

def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const

def get_kompas_api5():
    module = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0)
    api = Dispatch('Kompas.Application.5')
    return module, api

def copy_to_new_view():
    global doc, doc_app5
    # Всегда обращаемся к активному виду 
    view = doc.ViewsAndLayersManager.Views.ActiveView
    # Инициализируем контейнер со всеми объектами вида
    container = module7.IDrawingContainer(view)
    # Для удаления объектов
    # iKompasDocument1.Delete(line_dims)
    # Проходимся по контейнеру и отбираем только объкты-линии
    obj_to_copy = []
    for obj in container.Objects(0):
        if obj.DrawingObjectType in [1, 2, 3, 8, 31, 32, 33, 34]:
            if obj.Style == 1:
                # Берем не сам объект, а его reference (long)
                obj_to_copy.append(obj.Reference)

    # Добавляем новый вид за пределами чертежа
    new_view = doc.ViewsAndLayersManager.Views.Add(1)
    new_view.X = 300
    new_view.Y = 0
    new_view.Update()

    #kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    #  Создаем новый документ
    #kompas_document = Documents.AddWithDefaultSettings(kompas6_constants.ksDocumentFragment, True)

    #kompas_document_2d = API7.IKompasDocument2D(kompas_document)
    #iDocument2D = kompas_object.ActiveDocument2D

    # Создаем новую группу объектов и складываем туда все объекты для копирования
    g = doc_app5.ksNewGroup(1)
    for x in obj_to_copy:
        doc_app5.ksAddObjGroup(g, x)
    # Копируем группу в буфер обмена
    doc_app5.ksWriteGroupToClip(g,1)
    # Считываем группу из буфера и записываем в вид
    g_ = doc_app5.ksReadGroupFromClip()
    doc_app5.ksStoreTmpGroup(g_)

    # Разрушаем все виды, кроме системного
    i = 1
    views =  doc.ViewsAndLayersManager.Views
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
    module7, api7, const7 = get_kompas_api7()
    app7 = api7.Application
    #app7.Visible = True
    app7.HideMessage = const7.ksHideMessageNo

    module5, app5 = get_kompas_api5()

    # print('Укажите путь к файлу спецификации')
    # spec_path = input()

    print('Укажите путь к чертежам:')
    docs_path = input()

    # spec_path = r'C:\Users\UserPC\Documents\Macro Kompas\Спецификация.spw'
    # print(spec_path)
    # spec_doc = app7.Documents.Open(spec_path, True, False)

    spec_doc = app7.ActiveDocument
    # Если активный документ не Спецификация
    if spec_doc.DocumentType != 3:
        print("ERROR: Спецификация не открыта в Компасе")
    print('Запуск создания чертежей...')
    details = get_details_from_spec(spec_doc)

    # docs_path = r'C:\Users\UserPC\Documents\Macro Kompas\Конструкторские'

    new_folder_name = 'DXF {}'.format(spec_doc.Name.strip('.spw'))
   
    if not os.path.exists(docs_path + '\\{}'.format(new_folder_name)):
        os.mkdir(docs_path + '\\{}'.format(new_folder_name))

    for detail_name, (quantity, thickness) in details.items():

        path = docs_path + '\\' + detail_name + '.cdw'
        try:
            iKompasDocument = app7.Documents.Open(path, True, False)
            doc = module7.IKompasDocument2D(iKompasDocument)
            # Захватываем активный чертеж приложением api5
            doc_app5 = app5.ActiveDocument2D
        except:
            print(f'ERROR: Чертеж {detail_name} не был найден.')
        copy_to_new_view()

        new_path = docs_path + '\\{}\\{} {}шт {}мм.cdw'.format(new_folder_name, detail_name, quantity, thickness)

        flag = doc_app5.ksSaveDocumentEx(new_path, -1)
        if flag:
            print('INFO: Чертеж для детали {} успешно сохранен'.format(detail_name))
        else:
            print('ERROR: Ошибка при сохранении чертежа для детали {}'.format(detail_name))
        doc_app5.ksCloseDocument()
        doc.Close(0)
    # app7.Quit()
except Exception as e:
    print(e)
    # app7.Quit()
