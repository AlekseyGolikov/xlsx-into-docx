import pandas as pd

file=pd.read_excel('sourses/sourse.xlsx')
# print(file.head())
data_frame = pd.DataFrame(file)
# print(data_frame)
# dict[data_frame[1][0]]=data_frame[1][1]
# print(data_frame['Основное средство'][0])

from docxtpl import DocxTemplate

observBuilding1 = 'осмотр несущих конструкций, свайного основания, стеновых покрытий, дверных и'
observBuilding2 = 'оконных проемов.'
observMacht = 'осмотр несущих конструкций, свайного основания.'
damageList = ('ДП2001490','ДП2001510','ДП2001406','ДП2001403','ДП2001401','АА8639','ДП2001547','ДП2001548','ДП2001549','ДП2001550','АА0011240','ДП2001389')

def doc_creating():

    files = []


    for i in range(len(data_frame)):
        doc = DocxTemplate("sourses/template.docx")
        unit = str(data_frame['Основное средство'][i])
        # if 'ОЯЯНГКМ. 1. 1. ' in unit:
        #     unit = unit.replace('ОЯЯНГКМ. 1. 1. ','')
        if '"' in unit:
            unit = unit.replace('"','')

        if 'Емкость дренажно-канализационная' in unit:
            unit = unit.replace('Емкость дренажно-канализационная','Блок-бокс над емкостью дренажно-канализационной')
        elif 'Сбор и транспорт газа' in unit:
            unit = unit.replace('Сбор и транспорт газа','БКУЭ кранового узла ТГП')

        if 'ОЯЯНГКМ. 1. 1. ' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 1. ','')
        elif 'ОЯЯНГКМ. 1. 1.' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 1.','')
        elif 'ОЯЯНГКМ. 1.1. ' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1.1. ','')
        elif 'ОЯЯНГКМ. 1. 2. ' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 2. ','')
        elif 'ОЯЯНГКМ. 1. 2.' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 2.','')
        elif 'ОЯЯНГКМ. 1. 8. ' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 8. ','')
        elif 'ОЯЯНГКМ. 1. 8.' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 8.','')
        elif 'ОЯЯНГКМ. 1. 11. ' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 11. ','')
        elif 'ОЯЯНГКМ. 1. 11.' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 11.','')
        elif 'ОЯЯНГКМ. 1. 12. ' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 12. ','')
        elif 'ОЯЯНГКМ. 1. 12.' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 12.','')
        elif 'ОЯЯНГКМ. 1. 17. ' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 17. ','')
        elif 'ОЯЯНГКМ. 1. 17.' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 17.','')
        elif 'ОЯЯНГКМ. 1. 18. ' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 18. ','')
        elif 'ОЯЯНГКМ. 1. 18.' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 18.','')
        elif 'ОЯЯНГКМ. 1. 19. ' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 19. ','')
        elif 'ОЯЯНГКМ. 1. 19.' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 19.','')
        elif 'ОЯЯНГКМ. 1. 20. ' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 20. ','')
        elif 'ОЯЯНГКМ. 1. 20.' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 20.','')
        elif 'ОЯЯНГКМ. 1. 21. ' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 21. ','')
        elif 'ОЯЯНГКМ. 1. 21.' in unit:
            unit = unit.replace('ОЯЯНГКМ. 1. 21.','')
        elif 'ОНЯЯНГКМ.' in unit:
            unit = unit.replace('ОНЯЯНГКМ.','')
        elif 'ОЯЯНГКМ. ' in unit:
            unit = unit.replace('ОЯЯНГКМ. ','')
        elif 'ОЯЯНГКМ.' in unit:
            unit = unit.replace('ОЯЯНГКМ.','')
        elif ' Технологические сооружения.' in unit:
            unit = unit.replace(' Технологические сооружения.','')


        if ('БПО' in unit) or ('ДП' in unit):
            date = '17 ноября 2025г.'
        elif ('КЭ' in unit) or ('ПР' in unit):
            date = '17 ноября 2025г.'
        elif ('КОС' in unit) or ('ППС' in unit):
            date = '17 ноября 2025г.'
        elif ('КС' in unit) or ('ДКС' in unit):
            date = '18 ноября 2025г.'
        elif ('УДК' in unit) or ('УКПГ' in unit):
            date = '18 ноября 2025г.'
        elif ('УПН' in unit) or ('СОВ' in unit):
            date = '18 ноября 2025г.'
        elif ('ООВЭиОЭН' in unit) or ('№ 7' in unit):
            date = '18 ноября 2025г.'
        elif ('№11' in unit) or ('№ 11' in unit) or ('№3' in unit)  or ('№12' in unit)  or ('№ 12' in unit):
            date = '19 ноября 2025г.'
        elif ('№4' in unit) or ('№ 4' in unit)  or ('№15' in unit) or ('№ 15' in unit):
            date = '19 ноября 2025г.'
        elif ('№2' in unit) or ('№ 2' in unit)  or ('№5' in unit) or ('№ 5' in unit)  or ('№71' in unit) or ('№ 71' in unit):
            date = '19 ноября 2025г.'
        elif ('№7' in unit) or ('№ 7' in unit) :
            date = '19 ноября 2025г.'



        if 'ачта' in unit:
            # context = { 'unit' : data_frame['Основное средство'][i] , 'code': data_frame['Инвентарный номер'][i] , 'date' : date, 'observ1' : observMacht, 'observ2' : '', 'obj' : 'сооружения'}
            observ1 = observMacht
            observ2 = ''
            obj = 'сооружения'
        elif 'олниеотвод' in unit:
            # context = { 'unit' : data_frame['Основное средство'][i] , 'code': data_frame['Инвентарный номер'][i] , 'date' : date, 'observ1' : observMacht, 'observ2' : '', 'obj' : 'сооружения'}
            observ1 = observMacht
            observ2 = ''
            obj = 'сооружения'
        else:
            observ1 = observBuilding1
            observ2 = observBuilding2
            obj = 'здания'
        context = { 'unit' : unit , 'code': str(data_frame['Инвентарный номер'][i]) , 'i' : i+1, 'date' : date, 'observ1' : observ1, 'observ2' : observ2, 'obj' : obj}
        doc.render(context)
        if str(data_frame['Инвентарный номер'][i]) not in damageList:
            path_to_file = 'reports/'+ str(i+1) + " " + unit + '.docx'
        else:
            path_to_file = 'reports/damages/'+ str(i+1) + " " + unit + '.docx'
        doc.save(path_to_file)
        files.append(path_to_file)
    return files

from docxcompose.composer import Composer
from docx import Document

def merge_docs_with_page_breaks(output_path, *input_paths):

    base_doc = Document(input_paths[0])
    composer = Composer(base_doc)


    for file_path in input_paths[1:]:
        doc = Document(file_path)

        # adding page break before merging each document
        base_doc.add_page_break()
        composer.append(doc)

    composer.save(output_path)
    print(f"Documents merged successfully into {output_path}")


if __name__ == '__main__':
    files=doc_creating()
    output_file = "merged_document.docx"
    merge_docs_with_page_breaks(output_file, *files)
