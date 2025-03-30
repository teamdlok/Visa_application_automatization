from pathlib import Path
from docx import Document
import numpy as np
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import calendar
from FUNCTIONS_AND_CLASSES.FUNCTIONS import *
from FUNCTIONS_AND_CLASSES.request_factory import ZayavlenieConstructor
from FUNCTIONS_AND_CLASSES.anketa_factory import AnketaConstructor
from FUNCTIONS_AND_CLASSES.hodataistvo_factory import HodataistvoConstructor
from docxtpl import DocxTemplate



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

font_params = {
    "fontname": "Arial Narrow",
    "fontsize": 11,
    "bold": True,
}

company_choice = int(input("ВЫБЕРИТЕ КОМПАНИЮ \n 1 - АВАНТА \n 2 - ПРОФЕССИОНАЛ \n 3 - ВИЗАР ВОСТОК \n 4 - ТЭНФЭЙ \n"))
if company_choice == 1:
    company_name = "АВАНТА"
elif company_choice == 2:
    company_name = "ПРОФЕССИОНАЛ"
elif company_choice == 3:
    company_name = "ВИЗАР ВОСТОК"
elif company_choice == 4:
    company_name = "ТЭНФЭЙ"
    
zayavlenie_path = f"./templates/{company_name}/ЗАЯВЛЕНИЕ/zayavlenie_template.docx"
anketa_path = f"./templates/{company_name}/АНКЕТА/anketa_template.docx"
hodataistvo_path = f"./templates/{company_name}/ХОДАТАЙСТВО/hodataistvo_template.docx"


request_required = input("Нужно ли заявление? \n (ДА)  (НЕТ) \t").upper()
if request_required == "ДА" or request_required == "LF":
    request_required = True
else:
    request_required = False
month_dict = {}

# translation_required = input("Нужен ли перевод? \n (ДА) (НЕТ \t)").upper()
# if translation_required == "ДА" or request_required == "LF":
#     translation_required = True
# else:
#     translation_required = False

if request_required == True:
    newby_or_old = input("Делаем: 1 - Старый сотрудник \n 2 - Новый сотрудник \t")
    if newby_or_old == "1": 
        newby_request = False
    elif newby_or_old == "2":
        newby_request = True
    else: 
        raise Exception("Введено неправильное значение. Нужно цифры 1 или 2")

request = Document(zayavlenie_path) 
tables = request.tables


if request_required == True:

    rnp_input = input_prettier("номер рнп")
    if rnp_input == []:
        rnp_input = "2400"
    elif len(rnp_input) > 6:
        rnp_input = f" {''.join(rnp_input)}"
    else:
        rnp_input = f" {''.join(rnp_input)}"
        print(f"_________Ошибка? {''.join(rnp_input)}?")
    print(prettier_check_tool(rnp_input))

    day_rnp_start = input_prettier("день начала разрешения", True)
    print(prettier_check_tool(day_rnp_start))

    month_rnp_start = input_prettier("месяц начала разрешения", True)
    print(prettier_check_tool(month_rnp_start))

    year_rnp_start = input_prettier("год начала разрешения", False, True)
    print(prettier_check_tool(year_rnp_start))

    day_rnp_end = input_prettier("день конца разрешения", True)
    print(prettier_check_tool(day_rnp_end))

    month_rnp_end = input_prettier("месяц конца разрешения", True)
    print(prettier_check_tool(month_rnp_end))

    year_rnp_end = input_prettier("год конца разрешения", False, True)
    print(prettier_check_tool(year_rnp_end))

lastname_input = input_prettier("ФАМИЛИЮ")
print(prettier_check_tool(lastname_input))

name_input = input_prettier("ИМЯ")
print(prettier_check_tool(name_input))

lastname_eng_input = input_prettier("ФАМИЛИЮ НА АНГЛИЙСКОМ")

name_eng_input = input_prettier("ИМЯ НА АНГЛИЙСКОМ")

if request_required == True:

    namechange_input = input_prettier("СВЕДЕНИЯ ОБ ИЗМЕНЕНИИ ФИО")
    if str(''.join(namechange_input)) == "":
        namechange_input = list("НЕ МЕНЯЛ")
        namechange_check = False
    else:
        namechange_check = True
    print(prettier_check_tool(namechange_input))

if request_required == True:

    birthplace_input = input_prettier("МЕСТО РОЖДЕНИЯ")
    print(prettier_check_tool(birthplace_input))

birth_day = input_prettier("ДЕНЬ РОЖДЕНИЯ", True)
print(prettier_check_tool(birth_day))

birth_month = input_prettier("МЕСЯЦ РОЖДЕНИЯ (слово, например AUG)")
month_number = numbered_month(''.join(birth_month))
if month_number < 10:
    month_number = f"0{month_number}"
month_number = list(str(month_number))
print(prettier_check_tool(month_number))

birth_year = input_prettier("ГОД РОЖДЕНИЯ", False, True)
birth_year_int = int(''.join(birth_year))
if birth_year_int > 2005:
    print(f"______Ошибка? Родился в {birth_year_int}?______")
print(prettier_check_tool(birth_year))

birthdate_string = f"{''.join(birth_day)}.{''.join(month_number)}.{''.join(birth_year)}"

sex = input_prettier("ПОЛ (М) - (Ж)")
if str(''.join(sex)) == "М":
    male = "V"
    female = ""
elif str(''.join(sex)) == "Ж":
    female = "V"
    male = ""

else:
    print("Введён некоректный пол, ПОЛЕ ПОЛА БУДЕТ ПУСТЫМ")
    female = ""
    male = ""
print(prettier_check_tool(sex))

# if translation_required:
#     if sex == "М":
#         nationality = "Китаец"
#     elif sex == "Ж":
#         nationality = "Китаянка"
#     else: 
#         print(f"________Ошибка в поле? Пол - {sex} ???_____")
#         nationality = "Китаец"

passport_series = input_prettier("СЕРИЮ ПАСПОРТА")
if len(passport_series) > 3:
    print("___ОПЕЧАТКА?___ ПОЛЕ СЕРИИ ПАСПОРТА БУДЕТ ПУСТЫМ")
    passport_series = ""
print(prettier_check_tool(passport_series))

passport_number = input_prettier("НОМЕР ПАСПОРТА")
if len(passport_number) < 6:
    print(f"_____Ошибка? В номере паспорта всего {len(passport_number)} цифры")
print(prettier_check_tool(passport_number))

passport_date_day = input_prettier("ДЕНЬ ВЫДАЧИ ПАСПОРТА", True)
print(prettier_check_tool(passport_date_day))

passport_date_month = input_prettier("МЕСЯЦ ВЫДАЧИ ПАСПОРТА", True)
print(prettier_check_tool(passport_date_month))

passport_date_year = input_prettier("ГОД ВЫДАЧИ ПАСПОРТА", False, True)
passport_year_int = int(''.join(passport_date_year))
if passport_year_int < 2014:
    print(f"___Ошибка? Паспорт сделан в {passport_year_int}?___")
print(prettier_check_tool(passport_date_year))

passport_date_from_string = f"{''.join(passport_date_day)}.{''.join(passport_date_month)}.{''.join(passport_date_year)}"
if not int(''.join(passport_date_day)) < 2:
    if int(''.join(passport_date_day)) < 11:
        passport_date_to_string = f"0{int(''.join(passport_date_day))-1}.{''.join(passport_date_month)}.{int(''.join(passport_date_year))+10}"
    else:
        passport_date_to_string = f"{int(''.join(passport_date_day))-1}.{''.join(passport_date_month)}.{int(''.join(passport_date_year))+10}"
else:
    print("Паспорт начинается в 1 день месяца, функционал " +
          "для этого не готов, подправить месяц и день в окончании паспорта")
    if int(''.join(passport_date_day)) < 11:
        passport_date_to_string = f"0{int(''.join(passport_date_day))-1}.{''.join(passport_date_month)}.{int(''.join(passport_date_year))+10}"
    else:
        passport_date_to_string = f"{int(''.join(passport_date_day))-1}.{''.join(passport_date_month)}.{int(''.join(passport_date_year))+10}"

if request_required == True:

    passport_creator = input_prettier("КЕМ ВЫДАН ПАСПОРТ")
    print(prettier_check_tool(passport_creator))


address_of_living = input("Введите адрес регистрации  ").upper()
print(prettier_check_tool(address_of_living))

address_str, address_next_str = address_splitter(address_of_living)

who_meet = input_prettier("КТО ВСТРЕЧАЕТ РАБОТНИКА")
who_meet_company = input_prettier("КАКАЯ КОМПАНИЯ ВСТРАЧАЕТ (С ИНН)")

profession_name = input_prettier("ПРОФЕССИЮ")
print(prettier_check_tool(profession_name))

#### Дальше только код для анкеты и ходотайства

visa_blank_series = input_prettier("ВВЕДИТЕ СЕРИЮ ВИЗЫ")
visa_number = input_prettier("ВВЕДИТЕ НОМЕР ВИЗЫ")
visa_identificator = input_prettier("ВВЕДИТЕ ИДЕНТИФИКАТОР ВИЗЫ")
visa_invitation = input_prettier("ВВЕДИТЕ НОМЕР ПРИГЛАШЕНИЯ")

visa_day_start = input_prettier("ВВЕДИТЕ ДЕНЬ НАЧАЛА ВИЗЫ", True)
visa_month_start = input_prettier("ВВЕДИТЕ МЕСЯЦ НАЧАЛА ВИЗЫ", True)
visa_year_start = input_prettier("ВВЕДИТЕ ГОД НАЧАЛА ВИЗЫ", False, True)
visa_start_str = f"{''.join(visa_day_start)}.{''.join(visa_month_start)}.{''.join(visa_year_start)}"

visa_day_end = input_prettier("ВВЕДИТЕ ДЕНЬ КОНЦА ВИЗЫ", True)
visa_month_end = input_prettier("ВВЕДИТЕ КОНЦА НАЧАЛА ВИЗЫ", True)
visa_year_end = input_prettier("ВВЕДИТЕ ГОД КОНЦА ВИЗЫ \t", False, True)
visa_end_str = f"{''.join(visa_day_end)}.{''.join(visa_month_end)}.{''.join(visa_year_end)}"

table_contents = {
        'company_name': company_name,
        'lastname_rus': ''.join(lastname_input),
        'lastname_eng': ''.join(lastname_eng_input),
        'name_rus': ''.join(name_input),
        'name_eng': ''.join(name_eng_input),
        'birth_date': ''.join(birthdate_string),
        'passport_series': ''.join(passport_series),
        'passport_number': ''.join(passport_number),
        'passport_date_from': ''.join(passport_date_from_string),
        'passport_date_to': ''.join(passport_date_to_string),
        'who_meet': ''.join(who_meet),
        'who_meet_company': ''.join(who_meet_company),
        'address_of_living': ''.join(address_str), # Надо будет в класс передать без join, сырую версию
        'address_of_living_next': ''.join(address_next_str),
        'profession': ''.join(profession_name),
        'visa_blank_series': ''.join(visa_blank_series),
        'visa_number': ''.join(visa_number),
        'visa_identificator': ''.join(visa_identificator),
        'visa_date_start': ''.join(visa_start_str),
        'visa_date_end': ''.join(visa_end_str),
        'invitation_number': ''.join(visa_invitation),
        'male': male,
        'female': female,
        }


anketa = DocxTemplate(anketa_path)
hodataistvo = DocxTemplate(hodataistvo_path)

context = table_contents
current_time = datetime.now().strftime('%d.%m.%Y')

Path(f"OUTPUT/{company_name}").mkdir(parents=False, exist_ok=True)
Path(f"OUTPUT/{company_name}/ЗАЯВЛЕНИЕ").mkdir(parents=False, exist_ok=True)
Path(f"OUTPUT/{company_name}/ХОДАТАЙСТВО").mkdir(parents=False, exist_ok=True)
Path(f"OUTPUT/{company_name}/АНКЕТА").mkdir(parents=False, exist_ok=True)

if request_required == True:
    pass
    # request.save(f'OUTPUT/{company_name}/ЗАЯВЛЕНИЕ/{''.join(lastname_input)} {''.join(name_input)} - {current_time} .docx')

# anketa.render(context)
# anketa.save(f'OUTPUT/{company_name}/АНКЕТА/{''.join(lastname_input)} {''.join(name_input)} {current_time} АНКЕТА.docx')

# hodataistvo.render(context)
# hodataistvo.save(f'OUTPUT/{company_name}/ХОДАТАЙСТВО/{''.join(lastname_input)} {''.join(name_input)} {current_time} ХОДАТАЙСТВО.docx')

if request_required == True:

    zayavlenie_dict_of_params = dict(
        newby_request = newby_request,
        company_name = company_name,
        rnp_input = rnp_input,
        day_rnp_start = day_rnp_start,
        month_rnp_start = month_rnp_start,
        year_rnp_start = year_rnp_start,
        day_rnp_end = day_rnp_end,
        month_rnp_end = month_rnp_end,
        year_rnp_end = year_rnp_end,
        lastname_input = lastname_input,
        name_input = name_input,
        namechange_input = namechange_input,
        birthplace_input = birthplace_input,
        birth_day = birth_day,
        birth_month = birth_month,
        birth_year = birth_year,
        month_number = month_number,
        sex = sex,
        passport_series = passport_series,
        passport_number = passport_number,
        passport_date_day = passport_date_day,
        passport_date_month = passport_date_month,
        passport_date_year = passport_date_year,
        passport_creator = passport_creator,
        address_of_living = address_of_living,
        profession_name = profession_name
    )

anketa_dict_of_params = dict(
    company_name = company_name,
    lastname_rus = ''.join(lastname_input),
    lastname_eng = ''.join(lastname_eng_input),
    name_rus = ''.join(name_input),
    name_eng = ''.join(name_eng_input),
    birth_date = ''.join(birthdate_string),
    passport_series = ''.join(passport_series),
    passport_number = ''.join(passport_number),
    passport_date_from = ''.join(passport_date_from_string),
    passport_date_to = ''.join(passport_date_to_string),
    who_meet = ''.join(who_meet),
    who_meet_company = ''.join(who_meet_company),
    address_of_living = address_of_living, # Надо будет в класс передать без join, сырую версию
    profession = ''.join(profession_name),
    visa_blank_series = ''.join(visa_blank_series),
    visa_number = ''.join(visa_number),
    visa_identificator = ''.join(visa_identificator),
    visa_date_start = ''.join(visa_start_str),
    visa_date_end = ''.join(visa_end_str),
    invitation_number = ''.join(visa_invitation),
    male =  male,
    female = female,
)


hodataistvo_dict_of_params = dict(
    company_name = company_name,
    lastname_rus = ''.join(lastname_input),
    lastname_eng = ''.join(lastname_eng_input),
    name_rus = ''.join(name_input),
    name_eng = ''.join(name_eng_input),
    birth_date = ''.join(birthdate_string),
    passport_series = ''.join(passport_series),
    passport_number = ''.join(passport_number),
    passport_date_from = ''.join(passport_date_from_string),
    passport_date_to = ''.join(passport_date_to_string),
    visa_blank_series = ''.join(visa_blank_series),
    visa_number = ''.join(visa_number),
    visa_date_start = ''.join(visa_start_str),
    visa_date_end = ''.join(visa_end_str),
    male =  male,
    female = female,
)

if request_required == True:
    ZayavlenieConstructor(**zayavlenie_dict_of_params).request_factory()

AnketaConstructor(**anketa_dict_of_params).anketa_factory()
HodataistvoConstructor(**hodataistvo_dict_of_params).hodataistvo_factory()