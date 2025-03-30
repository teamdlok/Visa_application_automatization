
from docxtpl import DocxTemplate, InlineImage
import datetime
from FUNCTIONS_AND_CLASSES.FUNCTIONS import *

anketa = DocxTemplate('./templates\АВАНТА/АНКЕТА/anketa_template.docx')
hodataistvo = DocxTemplate('./templates/АВАНТА/ХОДАТАЙСТВО/hodataistvo_template.docx')

lastname_rus = input_prettier("Фамилию (рус)")
print(prettier_check_tool(lastname_rus))

name_rus = input_prettier("Имя (рус)")
print(prettier_check_tool(name_rus))

address_of_living = "ГОРОД ХАБАРОВСК, УЛИЦА ТУРБИННАЯ ДОМ 2, 4 КВАРТИРА, 45 ПОДЪЕЗД".split()
word_list = []
# word_str = " ".join(address_of_living)
address_str = ""
address_next_str = ""
# print(address_of_living)

for index, word in enumerate(address_of_living):
        if index == 0:
                address_str = word
        else:
                if not len(address_str) > 40:
                        address_str += ' ' + word
                else: 
                        address_next_str += ' ' + word
                print(len(address_str))

table_contents = {
        'lastname_rus': "Малов",
        'lastname_eng': 'Malov',
        'name_rus': "Игорь",
        'name_eng': "Igor",
        'birth_date': '20.07.2003',
        'passport_series': 'EL',
        'passport_number': '32423423',
        'passport_date_from': '10.08.2016',
        'passport_date_to': '09.08.2026',
        'who_meet': 'БАЙ ЛИЛИЯ СЕРГЕЕВНА',
        'who_meet_company': 'ИП БАЙ ЛИЛИЯ СЕРГЕЕВНА ИНН 439854359843',
        'address_of_living': address_str,
        'address_of_living_next': address_next_str,
        'profession': 'ПОВАР',
        'visa_blank_series': "12",
        'visa_number': '43434323',
        'visa_identificator': 'ШЕН32342432',
        'visa_date_start': '10.10.2024',
        'visa_date_end': '09.10.2025',
        'invitation_number': 'ЗДФ324234'
        }

context = table_contents
current_time = datetime.now().strftime('%d.%m.%Y')

hodataistvo.render(context)
hodataistvo.save(f'{lastname_rus} {name_rus} {current_time} АНКЕТА.docx')

anketa.render(context)
anketa.save(f'{lastname_rus} {name_rus} {current_time} ХОДАТАЙСТВО.docx')

