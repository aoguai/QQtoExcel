import os
import re
import sys
import time

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


def get_custom_filename_with_params(rule_string, message_group, message_object):
    matches = []
    text = rule_string

    # 用于查找[]中包含的所有模式的正则表达式
    pattern = r'(?<!["\'])\[((?:[^\[\]\x02]+|(?R)|\x02(?!\]))*)\](?<!["\'])'

    while True:
        match = regex.search(pattern, rule_string)
        if not match:
            break
        content = match.group(1).replace("\x02", "")
        matches.append(content)
        rule_string = regex.sub(pattern, '', rule_string, count=1)

    for content in matches:
        if content == '消息分组':
            text = text.replace('[消息分组]', message_group)
        if content == '消息对象':
            text = text.replace('[消息对象]', message_object)

    # 确保文件名长度不超过26个字符，并清除多余符号
    if len(text) > 26:
        # 清除多余符号
        text = text.replace("()", "")[:26]

        # 检查清除后的长度，如果为空字符串，设置默认文件名为 default_filename + 时间戳
        if not text:
            timestamp = str(int(time.time()))  # 获取当前时间戳，转换为字符串
            text = f"default_filename_{timestamp}"  # 设置默认文件名格式

    return data_win_file(text)
