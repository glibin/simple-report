#!coding:utf-8
__author__ = 'khalikov'

"""
Здесь должен находится код, который распарсит rtf файл, удалит лишние теги из
него и вернет новую строку с rtf
"""

class TokenList(list):
    def __init__(self, iterable, parent=None, sub_list=None, name=None):
        super(TokenList, self).__init__(iterable)
        self.parent = parent
        self.sub_list = sub_list
        self.name = name


def _get_cur_st(st, previous_st):
    cur_st = st
    #return st
    if previous_st is not None:
        if len(previous_st) > 1 and len(st) > 1:
            if st_equal(previous_st, st):
                cur_st = previous_st

            #if previous_st[:2] == st[:2]:
            #    cur_st = previous_st
    return cur_st


def st_equal(first, second):
    if len(first) >= len(second):
        st1 = first
        st2 = second
    else:
        st1 = second
        st2 = first
    length = len(st2)
    for i, token in enumerate(st1):
        if length < i + 1:
            return False
        if not token:
            return False
        if token != st2[i]:
            st2i = st2[i]
            ordinary_words_in_token = (
                token[0] != '\\'
            ) or (
                token[:1] == '\\u' and token[:4] not in ('\\uc1', '\\uc2')
            )
            if ordinary_words_in_token:
                return True

    return False


def parse(input_string):
    #top_top_structure = TokenList([], name='very_top_structure')
    top_structure = TokenList([], name='top_structure')
    structure = TokenList([], parent=top_structure, name='init_structure')
    st = structure
    chars_list = []
    previous_st = None
    lines_count = 1
    col_count = 0
    for i, char in enumerate(input_string):
        col_count += 1
        if char == '{':
            if chars_list:
                cur_st = _get_cur_st(st, previous_st)
                cur_st.append(''.join(chars_list).strip('\n'))
                #print cur_st[-1], 'cur_st #1', cur_st.name, repr(char), lines_count, col_count
            chars_list = []
            st = TokenList([], parent=st, sub_list=[], name=str('%s__%s' % (lines_count, col_count)))
            st.parent.append(st)
        elif char == '}':
            cur_st = st  # _get_cur_st(st, previous_st)
            cur_st.append(''.join(chars_list).strip('\n'))
            #print cur_st[-1], 'cur_st #2', cur_st.name, repr(char), lines_count, col_count
            chars_list = []
            previous_st = cur_st
            st = cur_st.parent
        else:
            if char == '\n':
                lines_count += 1
                col_count = 0
            if char in ' \n':
                if chars_list:
                    cur_st = st  # _get_cur_st(st, previous_st)
                    cur_st.append(''.join(chars_list).strip('\n'))
                    #print cur_st[-1], 'cur_st #3', cur_st.name, repr(char), lines_count, col_count
                    #if '\\insrsid' not in val or True:
                    #    cur_st.append(val)
                chars_list = []
            chars_list.append(char)
            #st.append(char)

    return structure


def normalize_rtf(filepath):
    with open(filepath) as fl1:
        input_string = fl1.read()
    return parse_string(input_string)


def parse_string(input_string):
    return input_string
    structure = parse(input_string)
    #print structure
    list_tokens = []
    make_rtf_tokens(structure[:1], list_tokens)
    return ''.join(list_tokens)


def make_rtf_tokens(structure, list_tokens):
    for token in structure:
        if isinstance(token, basestring):
            list_tokens.append(token)
        else:
            list_tokens.append('{')
            make_rtf_tokens(token, list_tokens)
            list_tokens.append('}')

if __name__ == '__main__':
    pass
    #with open('../tests/test_data/res-sluzh.rtf', 'w') as fl2:
    #    fl2.write(normalize_rtf('../tests/test_data/test-sluzh.rtf'))
    #print parse_file('../tests/test_data/test-sluzh.rtf')
    #with open('../tests/test_data/res-roll.rtf', 'w') as fl2:
    #    fl2.write(normalize_rtf('../tests/test_data/test_roll.rtf'))