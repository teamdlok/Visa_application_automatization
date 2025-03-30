def palladium_transliterate(text):
    # Словарь транскрипции (латиница -> кириллица)
    rules = {
        "zhao": "чжао",
        "men": "мэнь",
        "meng": "мэн",
        # Триграфы
        "ang": "ан",
        "eng": "эн",
        "ing": "ин",
        "ong": "ун",
        # Диграфы
        "zh": "чж",
        "ch": "ч",
        "sh": "ш",
        "ia": "я",
        "ie": "е",
        "iu": "ю",
        "ao": "ао",
        "ou": "оу",
        # Одиночные символы
        "a": "а",
        # "b": "б",
        # "c": "ц",
        # "d": "д",
        # "e": "э",
        # "f": "ф",
        # "g": "г",
        # "h": "х",
        # "i": "и",
        # "j": "цз",
        # "k": "к",
        # "l": "л",
        # "m": "м",
        # "n": "н",
        # "o": "о",
        # "p": "п",
        # "q": "ц",
        # "r": "ж",
        # "s": "с",
        # "t": "т",
        # "u": "у",
        # "v": "в",
        # "w": "в",
        # "x": "с",
        # "y": "й",
        # "z": "цз",
    }
    
    # Определяем максимальную длину ключа
    max_length = max(len(key) for key in rules.keys())
    
    lower_text = text.lower()
    result = []
    i = 0
    n = len(lower_text)
    
    while i < n:
        # Ищем совпадение для самой длинной возможной подстроки
        found = False
        for l in range(max_length, 0, -1):
            if i + l > n:
                continue  # Пропускаем, если выходим за пределы строки
            
            substr = lower_text[i:i+l]
            print(substr)
            if substr in rules:
                result.append(rules[substr])
                print(result)
                i += l
                found = True
                break
        
        if not found:
            # Если совпадений нет, добавляем оригинальный символ
            result.append(lower_text[i])
            i += 1
    
    # Восстанавливаем заглавную букву
    transliterated = ''.join(result)
    if text and text[0].isupper():
        transliterated = transliterated[0].upper() + transliterated[1:]
    
    return transliterated

# Примеры
# print(palladium_transliterate("Angming"))  # Анмин
# print(palladium_transliterate("zhong"))    # чжун
# print(palladium_transliterate("shang"))    # шан
print(palladium_transliterate("Zhaomenong"))      # янь (если добавить "ian": "янь" в правила)