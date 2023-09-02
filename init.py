import os
import re
import sys

import regex

WORKDIR = os.path.dirname(os.path.realpath(sys.argv[0]))


# 清洗后的文本中的非打印 ASCII 字符
def data_clean(text):
    ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
    text = ILLEGAL_CHARACTERS_RE.sub(r'', text)
    return text


# 清洗 windows 文件名中的非法字符，只保留中英文、数字与下划线
def data_win_file(text):
    ILLEGAL_CHARACTERS_RE = re.compile(r'[^\u4e00-\u9fa5^a-z^A-Z^\d_]')
    txt = ILLEGAL_CHARACTERS_RE.sub(r'()', text)
    return txt


# 设置默认输入值函数
def get_input(msg, default='Y'):
    r = input(msg)
    if r == '':
        return default
    return r


def get_custom_filename_with_params(rule_string, message_group, message_object, index):
    matches = []
    text = rule_string
    while True:
        match = regex.search(r'(?<!["\'])\[((?:[^\[\]\x02]+|(?R)|\x02(?!\]))*)\](?<!["\'])', rule_string)
        if not match:
            break
        content = match.group(1).replace("\x02", "")
        matches.append(content)
        rule_string = regex.sub(r'(?<!["\'])\[((?:[^\[\]\x02]+|(?R)|\x02(?!\]))*)\](?<!["\'])', '', rule_string,
                                count=1)
    for content in matches:
        if content == '消息分组':
            text = text.replace('[消息分组]', message_group)
        if content == '消息对象':
            text = text.replace('[消息对象]', message_object)
        if content == '序号':
            text = text.replace('[序号]', str(index))
    return data_win_file(text)
