# 清洗excel中的非法字符，都是不常见的不可显示字符，例如退格，响铃等
import os
import re
import sys

workdir = os.path.dirname(os.path.realpath(sys.argv[0]))


def data_clean(text):
    ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
    text = ILLEGAL_CHARACTERS_RE.sub(r'', text)
    return text


# 清洗windows文件名中的非法字符，只保留中英文和数字
def data_win_file(text):
    ILLEGAL_CHARACTERS_RE = re.compile(r'[^\u4e00-\u9fa5^a-z^A-Z^\d]')
    txt = ILLEGAL_CHARACTERS_RE.sub(r'()', text)
    return txt


# 设置默认输入值函数
def get_input(msg, default='Y'):
    r = input(msg)
    if r == '':
        return default
    return r
