from datetime import datetime
from FUNCTIONS_AND_CLASSES.FUNCTIONS import *
from docxtpl import DocxTemplate

# Параметры стилей для документа заявлений
font_params = {
    "fontname": "Arial Narrow",
    "fontsize": 11,
    "bold": True,
}


def address_splitter(address_of_living):
    string_is_too_big = False
    address_str = ""
    address_next_str = ""
    address_of_living = address_of_living.split()
    for index, word in enumerate(address_of_living):
        if index == 0:
                address_str = word
        else:
                if not len(address_str + word) > 42 and string_is_too_big == False:
                        address_str += f" {word}"
                else: 
                        address_next_str += f" {word}"
                        string_is_too_big = True



     
    return address_str, address_next_str


class AnketaConstructor():
    """"Конструктор для создания заявления, получает информацию о человеке, и вписывает
    в документ для подачи заявления, работает как для новеньких так и для старых типов заявлений"""

    def __init__(self, company_name,  lastname_rus, lastname_eng, name_rus, name_eng,
                birth_date, passport_series, passport_number,
                passport_date_from, passport_date_to, who_meet,
                who_meet_company, address_of_living,
                profession, visa_blank_series, visa_number,
                visa_identificator, visa_date_start, visa_date_end,
                invitation_number, male, female,
                ):
        super().__init__()
        self.company_name = company_name
        self.lastname_rus = lastname_rus
        self.lastname_eng = lastname_eng
        self.name_rus = name_rus
        self.name_eng = name_eng
        self.birth_date = birth_date #31.12.2024
        self.passport_series = passport_series
        self.passport_number = passport_number
        self.passport_date_from = passport_date_from #31.12.2024
        self.passport_date_to = passport_date_to #31.12.2024
        self.who_meet = who_meet
        self.who_meet_company = who_meet_company
        self.address_of_living = address_of_living
        self.profession = profession
        self.visa_blank_series = visa_blank_series
        self.visa_number = visa_number
        self.visa_identificator = visa_identificator
        self.visa_date_start = visa_date_start
        self.visa_date_end = visa_date_end
        self.invitation_number = invitation_number
        self.male = male
        self.female = female


    def save_document(self, table_contents, company_name, lastname, name ):
        current_time = datetime.now().strftime('%d.%m.%Y')
        anketa_path = f"./templates/{company_name}/АНКЕТА/anketa_template.docx"
        anketa = DocxTemplate(anketa_path)
        anketa.render(table_contents)
        anketa.save(f'OUTPUT/{company_name}/АНКЕТА/{lastname} {name} {current_time} АНКЕТА КЛАССОМ.docx')


    def anketa_factory(self):
        

        address_str, address_next_str = address_splitter(self.address_of_living)
        print(f"ВОТ АДРЕСС КОТОРЫЙ ПОЛУЧАЕМ - {self.address_of_living} \n")

        print(address_str)
        print(address_next_str)

        table_contents = {
        'lastname_rus': self.lastname_rus,
        'lastname_eng': self.lastname_eng,
        'name_rus': self.name_rus,
        'name_eng': self.name_eng,
        'birth_date': self.birth_date,
        'passport_series': self.passport_series,
        'passport_number': self.passport_number,
        'passport_date_from': self.passport_date_from,
        'passport_date_to': self.passport_date_to,
        'who_meet': self.who_meet,
        'who_meet_company': self.who_meet_company,
        'address_of_living': ''.join(address_str),
        'address_of_living_next': ''.join(address_next_str),
        'profession': self.profession,
        'visa_blank_series': self.visa_blank_series,
        'visa_number': self.visa_number,
        'visa_identificator': self.visa_identificator,
        'visa_date_start': self.visa_date_start,
        'visa_date_end': self.visa_date_end,
        'invitation_number': self.invitation_number,
        'male': self.male,
        'female': self.female,
        }


        self.save_document(table_contents, self.company_name, self.lastname_rus, self.name_rus)
