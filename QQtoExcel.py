import sys
import openpyxl
import os
import re
import time
from tqdm import tqdm

workdir = os.path.dirname(os.path.realpath(sys.argv[0]))


def init_info():
    print("—————————————————欢迎使用———————————————————")
    print("——————————————QQtoExcelV1.1————————————————")
    print("项目地址：https://github.com/aoguai/QQtoExcel\n")


# 清洗excel中的非法字符，都是不常见的不可显示字符，例如退格，响铃等
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


def get_QQChat_record(route, time_list_out=True, name_list_out=True, uid_list_out=True, cont_list_out=True):
    object_file_name_list = []  # 文件名称列表（分组_备注）
    object_list = []  # 消息对象列表

    time_list = []  # 消息时间列表
    name_list = []  # 昵称列表
    uid_list = []  # QQ or 邮箱列表
    cont_list = []  # 内容列表

    f = open(route, encoding="utf-8")
    strs = f.read()
    f.close()

    q_pattern = r'(={64}([\s\S\消息分组:\s\S]{9,32})={64}([\s\S\消息分组:\s\S]{9,32})={64})'  # 定义分隔符
    result = re.split(q_pattern, strs)  # 以pattern的值 分割字符串

    # # 验证是否是4位一循环
    # print(result[0])  # 默认无关内容
    # print(result[1+(4*n)])  # 分组-昵称
    # print(result[2+(4*n)])  # 分组
    # print(result[3+(4*n)])  # 昵称
    # print(result[4+(4*n)])  # 消息内容
    # print((len(result)-1)/4)  # 总人数

    # 多对象获取消息内容
    for i in range(int((len(result) - 1) / 4)):
        object_file_name_list.append(data_win_file(data_clean(
            result[2 + (4 * i)].replace('\n', '').replace('\r', '').replace(' ', '').replace('消息分组:',
                                                                                             ''))) + "_" + data_win_file(
            data_clean(
                result[3 + (4 * i)].replace('\n', '').replace('\r', '').replace(' ', '').replace('消息对象:', ''))))
        # 群消息匹配规则
        pattern = re.compile(
            '(\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{1,2}:\d{1,2})(.+[(|<](.*)[)|>])([\s\S]*?)(\n\s*\n)')
        # 好友消息匹配规则
        pattern2 = re.compile(
            '(\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{1,2}:\d{1,2})(.+)([\s\S]*?)(\n\s*\n)')

        # 添加该消息对象各项消息
        m = pattern.findall(result[4 + (4 * i)])
        if len(m) > 0:
            for j in m:
                if time_list_out:
                    time_list.append(j[0])
                if name_list_out:
                    name_list.append(data_clean(j[1].replace(j[2], '').replace('()', '').replace('<>', '')))
                if uid_list_out:
                    uid_list.append(j[2])
                if cont_list_out:
                    cont_list.append(data_clean(j[3][1:]))
        else:
            m = pattern2.findall(result[4 + (4 * i)])
            if len(m) >= 0:
                for j in m:
                    if time_list_out:
                        time_list.append(j[0])
                    if name_list_out:
                        name_list.append(data_clean(j[1].replace(j[2], '').replace('()', '').replace('<>', '')))
                    if uid_list_out:
                        uid_list.append('')
                    if cont_list_out:
                        cont_list.append(data_clean(j[2][1:]))
            else:
                break

        object_list.append([time_list, name_list, uid_list, cont_list])

        # # 输出
        # print(len(time_list),len(name_list),len(uid_list),len(cont_list))
        # for i in range(len(time_list)):
        #     print("time:"+time_list[i]+"\nname:"+name_list[i]+"\nuid:"+uid_list[i]+"\ncont:"+cont_list[i]+"\n===========")
        # print("共：", str(len(time_list)), "条消息")

        # 清空列表
        time_list = []
        name_list = []
        uid_list = []
        cont_list = []

    # print(len(object_file_name_list),len(object_list))  # 判断是否一对象一文件
    return object_file_name_list, object_list


# 创建Excel文件
def QQtoExcel(QQChat_route, workdirs=workdir, sheet_name="test", time_list_out=True,
              name_list_out=True,
              uid_list_out=True, cont_list_out=True):
    object_file_name_list, object_list = get_QQChat_record(QQChat_route, time_list_out, name_list_out,
                                                           uid_list_out, cont_list_out)

    file_path = []  # 输出目录列表
    for i in range(len(object_file_name_list)):
        file_path.append(os.path.join(workdirs, object_file_name_list[i] + '.xls'))

    row0_z = get_input("是否自定义可选项标题（留空则默认N）（Y/N）：", "N")

    # 写入Excel标题
    row0 = []  # 标题列表
    row1 = []  # 标题序号列表
    if row0_z == 'N' or row0_z == 'n':
        if time_list_out:
            row0.append("时间")
            row1.append("time_list")
        if name_list_out:
            row0.append("昵称")
            row1.append("name_list")
        if uid_list_out:
            row0.append("QQ（邮箱）")
            row1.append("uid_list")
        if cont_list_out:
            row0.append("内容")
            row1.append("cont_list")
    else:
        if time_list_out:
            row0.append(get_input("请输入原'时间'可选项标题（留空则默认'时间'）：", "时间"))
            row1.append("time_list")
        if name_list_out:
            row0.append(get_input("请输入原'昵称'可选项标题（留空则默认'昵称'）：", "昵称"))
            row1.append("name_list")
        if uid_list_out:
            row0.append(get_input("请输入原'QQ（邮箱）'可选项标题（留空则默认'QQ（邮箱）'）：", "QQ（邮箱）"))
            row1.append("uid_list")
        if cont_list_out:
            row0.append(get_input("请输入原'内容'可选项标题（留空则默认'内容'）：", "内容"))
            row1.append("cont_list")

    if len(row0) <= 0:
        print("请选择至少一项输出内容")
        return

    for i in tqdm(range(len(file_path)), '导出中'):
        time_list = object_list[i][0]
        name_list = object_list[i][1]
        uid_list = object_list[i][2]
        cont_list = object_list[i][3]

        # 创建workbook和sheet对象
        workboot = openpyxl.Workbook()
        worksheet = workboot.active
        worksheet.title = sheet_name  # 设置工作表的名字

        for j in range(len(row0)):
            worksheet.cell(1, j + 1, row0[j])

        # 写入内容
        if time_list_out:
            for k in range(len(time_list)):
                worksheet.cell(k + 2, row1.index("time_list") + 1, time_list[k])
        if name_list_out:
            for k in range(len(name_list)):
                worksheet.cell(k + 2, row1.index("name_list") + 1, name_list[k])
        if uid_list_out:
            for k in range(len(uid_list)):
                worksheet.cell(k + 2, row1.index("uid_list") + 1, uid_list[k])
        if cont_list_out:
            for k in range(len(cont_list)):
                worksheet.cell(k + 2, row1.index("cont_list") + 1, cont_list[k])

        workboot.save(file_path[i])
        workboot.close()


if __name__ == "__main__":
    init_info()
    while True:
        QQChat_route = input("请输入您导出的聊天记录txt路径（支持单好友、群聊与[全部消息记录.txt]）：")
        workdirs = get_input("请输入您转换后欲保存的目录（留空则默认保存在out目录中）：", workdir + "\\out\\")
        sheet_name = data_clean(get_input("请输入工作表名（留空则默认test）：", "test"))
        time_o = get_input("是否导出时间（留空则默认Y）（Y/N）：")
        name_o = get_input("是否导出昵称（留空则默认Y）（Y/N）：")
        uid_o = get_input("是否导出uid（留空则默认Y）（Y/N）：")
        cont_o = get_input("是否导出内容（留空则默认Y）（Y/N）：")
        if time_o == "Y" or time_o == "y":
            time_list_out = True
        else:
            time_list_out = False
        if name_o == "Y" or name_o == "y":
            name_list_out = True
        else:
            name_list_out = False
        if uid_o == "Y" or uid_o == "y":
            uid_list_out = True
        else:
            uid_list_out = False
        if cont_o == "Y" or cont_o == "y":
            cont_list_out = True
        else:
            cont_list_out = False

        start = time.perf_counter()
        QQtoExcel(QQChat_route=QQChat_route, workdirs=workdirs, sheet_name=sheet_name, time_list_out=time_list_out,
                  name_list_out=name_list_out,
                  uid_list_out=uid_list_out, cont_list_out=cont_list_out)
        end = time.perf_counter()
        print("导出完成\n耗时：", end - start, "秒")

        out_txt = input("是否退出（Y/N）：")
        if out_txt == "Y" or out_txt == "y":
            break

    sys.exit(0)
