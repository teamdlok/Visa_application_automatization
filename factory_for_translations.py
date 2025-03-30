from docxtpl import DocxTemplate, InlineImage
from datetime import datetime
from pathlib import Path
from FUNCTIONS_AND_CLASSES.FUNCTIONS import *



table_contents = []

company_name = input("Введите название компании ").upper().replace(" ", "")
print(prettier_check_tool(company_name))


blank = DocxTemplate('templates/translation_template.docx')

blank_spravki = DocxTemplate('templates/translation_spravki_template.docx')


lastname_rus = input("\nВведите фамилию (рус) ").upper()
print(prettier_check_tool(lastname_rus))

name_rus = input("Введите имя (рус) ").upper()
print(prettier_check_tool(name_rus))

lastname_eng = input("Введите фамилия (анг) ").upper()
print(prettier_check_tool(lastname_eng))

name_eng = input("Введите имя (анг) ").upper()
print(prettier_check_tool(name_eng))

passport_series = input("\nВведите серию паспорта ").upper()
print(prettier_check_tool(passport_series))

passport_number = input("Введите номер паспорта ")
print(prettier_check_tool(passport_number))

day_birth = ''.join(input_prettier("день рождения ", True))
print(prettier_check_tool(day_birth))

month_birth = input("Введите месяц рождения ").lower()
print(prettier_check_tool(month_birth))

year_birth = ''.join(input_prettier("год рождения ", normalize_year=True))
print(prettier_check_tool(year_birth))

passport_day_from = ''.join(input_prettier("день начала паспорта ", True))
print(prettier_check_tool(passport_day_from))

passport_month_from = input("Введите месяц начала паспорта ").lower()
print(prettier_check_tool(passport_month_from))

passport_year_from = ''.join(input_prettier("год начала паспорта ", normalize_year=True))
print(prettier_check_tool(passport_year_from))

passport_day_to = str(int(passport_day_from)-1)
if int(passport_day_to) < 1:
    print(F"____ОШИБКА ДЕНЬ  МЕНЬШЕ 1 ({passport_day_to}) НУЖНО БУДЕТ ИСПРАВИТЬ В ДОКУМЕНТЕ")
print(prettier_check_tool(passport_day_to))

if int(passport_day_to) < 10:
    passport_day_to = f"0{passport_day_to}"


passport_month_to = passport_month_from
print(prettier_check_tool(passport_month_to))

passport_year_to = str(int(passport_year_from)+10)
print(prettier_check_tool(passport_year_to))

sex = input("\nВведите пол (М) - (Ж) ").upper()
if sex == "М":
    nationality = "Китаец"
elif sex == "Ж":
    nationality = "Китаянка"
else: 
    print(f"________Ошибка в поле? Пол - {sex} ???_____")
print(prettier_check_tool(sex))

place_of_birth = input("\nВведите место рождения  \n 1 - Хэйлунцзян \n 2 - Цзилинь \n 3 - Шаньдун  \n 4 - Гуандун" +
                       "\n 5 - ЧЖЭЦЗЯН \n 6 - ХЭБЭЙ  \t").capitalize()
if place_of_birth == "1":
    place_of_birth = "Хэйлунцзян"
elif place_of_birth == "2":
    place_of_birth = "Цзилинь"
elif place_of_birth == "3":
    place_of_birth = "Шаньдун"
elif place_of_birth == "4":
    place_Of_birth = "Гуандун"
elif place_of_birth == "5":
    place_Of_birth = "ЧЖЭЦЗЯН"
elif place_of_birth == "6":
    place_Of_birth = "ХЭБЭЙ"

print(prettier_check_tool(place_of_birth))

place_of_issue = input("Введите место выдачи паспорта \n 1 - Хэйлунцзян \n 2 - Цзилинь \n 3 - Хабаровск \n 4 - Гуандун" +
                       " \n 5 - ЧЖЭЦЗЯН  \n 6 - ХЭБЭЙ \t").capitalize()
if place_of_issue == "1":
    place_of_issue = "Хэйлунцзян"
elif place_of_issue == "2":
    place_of_issue = "Цзилинь"    
elif place_of_issue == "3":
    place_of_issue = "Хабаровск"
elif place_of_issue == "4":
    place_of_issue == "Гуандун"
elif place_of_issue == "5":
    place_of_issue == "ЧЖЭЦЗЯН"
elif place_of_issue == "6":
    place_of_issue == "ХЭБЭЙ"
print(prettier_check_tool(place_of_issue))

who_give = input("\nВведите кем выдан паспорт \n 1 - Государственное управление по делам иммиграции КНР" +
                "\n 2 - Министерство общественной безопасности въезда и выезда из страны" +
                "\n 3 - Генеральное консульство Китайской Народной Республики в г.Хабаровске \t")
if who_give == "1":
    who_give = "Государственное управление по делам иммиграции КНР"
elif who_give == "2":
    who_give = "Министерство общественной безопасности въезда и выезда из страны"
elif who_give == "3":
    who_give = "Генеральное консульство Китайской Народной Республики в г.Хабаровске"

print(prettier_check_tool(who_give))

passport_number_machine = input("Введите значения снизу паспорта (НЕ ВПИСЫВАТЬ ПАСПОРТ И '<<<<' ) \t").replace("<", "")
passport_number_machine = f"{passport_series}{passport_number}{passport_number_machine}"
print(prettier_check_tool(passport_number_machine))


has_old_passport = input("\nЕсть сведения о старом паспорте? (Да) - (Нет) ").upper()
print(prettier_check_tool(has_old_passport))

if not has_old_passport == "НЕТ":
    old_passport = True
    old_passport_series = input("Введите серию старого паспорта ").upper()
    print(prettier_check_tool(old_passport_series))
    old_passport_number = input("Введите номер старого паспорта ")
    print(prettier_check_tool(old_passport_number))
    # old_passport_day = input("Введите день замены старого паспорта ")
    # if int(old_passport_day) < 10:
    #     old_passport_day = f"0{old_passport_day}"
    # print(prettier_check_tool(old_passport_day))
    # old_passport_month = input("Введите месяц замены старого паспорта ").lower()
    # print(prettier_check_tool(old_passport_month))
    # old_passport_year = ''.join(input_prettier("Введите год замены старого паспорта ", normalize_year=True))
    # print(prettier_check_tool(old_passport_year))
    # old_passport_city = input("Введите город замены старого паспорта \n" +
    #                            "1 - Хэйлунцзян \n" +
    #                            "2 - Цзилинь \n" +
    #                            "3 - Хабаровск \t").capitalize()
    # if old_passport_city == "1":
    #     old_passport_city = "Хэйлунцзян"
    # elif old_passport_city == "2":
    #     old_passport_city = "Цзилинь"
    # elif old_passport_city == "3":
    #     old_passport_city = "Хабаровск"

elif has_old_passport == "НЕТ":
    old_passport = False
    old_passport_series = False
    old_passport_number = False
    old_passport_day = False
    old_passport_month = False
    old_passport_year = False
    old_passport_city = False
else: 
    print(f"______Ошибка? введено {has_old_passport} ???_______")
    old_passport = False
    old_passport_series = False
    old_passport_number = False
    old_passport_day = False
    old_passport_month = False
    old_passport_year = False
    old_passport_city = False

table_contents = dict(
    lastname_rus = lastname_rus,
    lastname_eng = lastname_eng,
    name_rus = name_rus,
    name_eng = name_eng,
    day_birth = day_birth,
    month_birth = month_birth,
    year_birth = year_birth,
    passport_series = passport_series,
    passport_number = passport_number,
    passport_day_from = passport_day_from,
    passport_month_from = passport_month_from,
    passport_year_from = passport_year_from,
    passport_day_to = passport_day_to,
    passport_month_to = passport_month_to,
    passport_year_to = passport_year_to,
    sex = sex,
    nationality = nationality,
    place_of_birth = place_of_birth,
    place_of_issue = place_of_issue,
    old_passport = old_passport,
    old_passport_series = old_passport_series,
    old_passport_number = old_passport_number,
    old_passport_day = passport_day_from,
    old_passport_month = passport_month_from,
    old_passport_year = passport_year_from,
    old_passport_city = place_of_issue,
    passport_number_machine = passport_number_machine,
    who_give = who_give,
)

context = table_contents

Path(f"OUTPUT/translations_for_sergey/{company_name}").mkdir(parents=False, exist_ok=True)
Path(f"OUTPUT/translations_for_sergey/{company_name}/ДЛЯ_СПРАВОК").mkdir(parents=False, exist_ok=True)
current_time = datetime.now().strftime('%d.%m.%Y')

blank.render(context)

blank.save(f'output_reserv/n{lastname_rus} {name_rus} - {current_time} БОЛЬШОЙ.docx')
blank.save(f'OUTPUT/translations_for_sergey/{company_name}/{lastname_rus} {name_rus} - {current_time} БОЛЬШОЙ.docx')


blank_spravki.render(context)

blank_spravki.save(f'output_reserv/n{lastname_rus} {name_rus} - {current_time} СПРАВКИ.docx')
blank_spravki.save(f'OUTPUT/translations_for_sergey/{company_name}/ДЛЯ_СПРАВОК/{lastname_rus} {name_rus} - {current_time} СПРАВКИ.docx')

