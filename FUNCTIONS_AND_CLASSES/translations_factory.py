from datetime import datetime
from FUNCTIONS_AND_CLASSES.FUNCTIONS import *
from docxtpl import DocxTemplate

# Параметры стилей для документа заявлений
font_params = {
    "fontname": "Arial Narrow",
    "fontsize": 11,
    "bold": True,
}


# Используется для конвертации числового месяца в месяц словом
month_by_word = {1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля', 5: 'мая', 6: 'июня',
      7: 'июля', 8: 'августа', 9: 'сентября', 10: 'октября', 11: 'ноября', 12: 'декабря'} 



class TranslationConstructor():
    """"Конструктор для создания заявления, получает информацию о человеке, и вписывает
    в документ для подачи заявления, работает как для новеньких так и для старых типов заявлений"""

    def __init__(self, company_name,  lastname_rus, lastname_eng, name_rus, name_eng,
                birth_day, birth_month, birth_year, passport_series, passport_number,
                passport_date_day_from, passport_date_month_from, passport_date_year_from, passport_day_to,
                sex, nationality, place_of_birth, place_of_issue, old_passport,
                old_passport_series, old_passport_number, old_passport_day,
                old_passport_month, old_passport_year, old_passport_city,
                passport_number_machine, who_give, passport_month_to,
                passport_year_to, visa_date_start, visa_date_end
                ):
        super().__init__()
        self.company_name = company_name
        self.lastname_rus = lastname_rus
        self.lastname_eng = lastname_eng
        self.name_rus = name_rus
        self.name_eng = name_eng
        self.birth_day = birth_day
        self.birth_month = birth_month
        self.birth_year = birth_year
        self.passport_series = passport_series
        self.passport_number = passport_number
        self.passport_date_day_from = passport_date_day_from
        self.passport_date_month_from = passport_date_month_from
        self.passport_date_year_from = passport_date_year_from
        self.passport_day_to = passport_day_to
        self.sex = sex
        self.nationality = nationality
        self.place_of_birth = place_of_birth
        self.place_of_issue = place_of_issue
        self.old_passport = old_passport
        self.old_passport_series = old_passport_series
        self.old_passport_number = old_passport_number
        self.old_passport_day = old_passport_day
        self.old_passport_month = old_passport_month
        self.old_passport_year = old_passport_year
        self.old_passport_city = old_passport_city
        self.passport_number_machine = passport_number_machine
        self.who_give = who_give
        self.passport_month_to = passport_month_to
        self.passport_year_to = passport_year_to
        self.visa_date_start = visa_date_start
        self.visa_date_end = visa_date_end


    def save_document(self, table_contents, company_name, lastname, name ):
        current_time = datetime.now().strftime('%d.%m.%Y')
        anketa_path = f"./templates/{company_name}/ХОДАТАЙСТВО/anketa_template.docx"
        anketa = DocxTemplate(anketa_path)
        anketa.render(table_contents)
        anketa.save(f'OUTPUT/{company_name}/АНКЕТА/{lastname} {name} {current_time} ХОДАТАЙСТВО КЛАССОМ.docx')


    def translation_factory(self):
        

        table_contents = {

        }


        self.save_document(table_contents, self.company_name, self.lastname_rus, self.name_rus)
