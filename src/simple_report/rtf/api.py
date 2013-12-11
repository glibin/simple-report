#!coding:utf-8
__author__ = 'khalikov'

CHARS_MAP = dict(
    zip(
        u'АБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЫЪЬЭЮЯабвгдежзийклмнопрстуфхцчшщыъьэюя',
        range(1040, 1103)
    )
)
CHARS_MAP[u'Ё'] = 1025
CHARS_MAP[u'ё'] = 1105

CHARS = dict([(x, '\\u%s\\\'3f' % y) for x, y in CHARS_MAP.items()])
CHARS[' '] = ' '
for i in range(10):
    CHARS[str(i)] = str(i)

def convert_dict(dictionary):
    """
    Конвертирует значения из словаря в пригодные для записи в rtf-файл
    """
    new_dictionary = {}
    for key, value in dictionary.items():
        #if not isinstance(value, basestring):
        try:
            value = unicode(value)
        except Exception:
            continue
        res = []
        for v in value:
            val = v
            if v in CHARS:
                val = CHARS[v]
            res.append(val)
            #res.append(to_hex(v.encode('cp1251')))
            #res.append(v.encode('utf-8'))
            #res.append(v.encode('unicode_escape'))
        new_dictionary[key] = ''.join(res)
        #new_dictionary[key] = to_hex(value.encode('cp1251'))
        #print new_dictionary[key]
        #new_dictionary[key] = value.encode('utf-8')
        #new_dictionary[key] = to_hex(value.encode('cp1251'))
    return new_dictionary


def to_hex(string):
    """
    Перекодирует строку в hex с нужным для записи в rtf префиксом
    """
    lst = []
    for ch in string:
        hv = hex(ord(ch)).replace('0x', '')
        if len(hv) == 1:
            hv = '0'+hv
        hv = '\\\''+hv
        lst.append(hv)
    return ''.join(lst)


def do_replace(text, params):
    """
    Ищет знаки '#' в тексте rtf-шаблона и подставляет значения из словаря
    """
    for key_param, value in params.items():
        if key_param in text:
            text = text.replace('#%s#' % key_param, value)
    return text