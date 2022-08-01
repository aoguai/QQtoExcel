import time

from QQtoExcel import QQtoExcel
from init import *


def init_info():
    print("—————————————————欢迎使用———————————————————")
    print("——————————————QQtoExcelV1.5————————————————")
    print("项目地址：https://github.com/aoguai/QQtoExcel\n")


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
        out_o = get_input("是否按好友导出（留空则默认Y）（Y/N）：")
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
        if out_o == "Y" or out_o == "y":
            out_type = 0
        else:
            out_type = 1

        row0_z = get_input("是否自定义可选项标题（留空则默认N）（Y/N）：", "N")

        time_row_text = "时间"
        name_row_text = "昵称"
        uid_row_text = "QQ（邮箱）"
        cont_row_text = "内容"

        if row0_z == 'Y' or row0_z == 'y':
            if time_list_out:
                time_row_text = get_input("请输入原'时间'可选项标题（留空则默认'时间'）：", "时间")
            if name_list_out:
                name_row_text = get_input("请输入原'昵称'可选项标题（留空则默认'昵称'）：", "昵称")
            if uid_list_out:
                uid_row_text = get_input("请输入原'QQ（邮箱）'可选项标题（留空则默认'QQ（邮箱）'）：", "QQ（邮箱）")
            if cont_list_out:
                cont_row_text = get_input("请输入原'内容'可选项标题（留空则默认'内容'）：", "内容")

        start = time.perf_counter()

        qe = QQtoExcel(QQChat_route=QQChat_route, file_path=workdirs, sheet_name=sheet_name,
                       time_list_out=time_list_out,
                       name_list_out=name_list_out, uid_list_out=uid_list_out,
                       cont_list_out=cont_list_out, time_row_text=time_row_text,
                       name_row_text=name_row_text, uid_row_text=uid_row_text,
                       cont_row_text=cont_row_text, out_type=out_type)
        qe.toExcel()

        end = time.perf_counter()
        print("导出完成\n耗时：", end - start, "秒")

        out_txt = input("是否退出（Y/N）：")
        if out_txt == "Y" or out_txt == "y":
            break

    sys.exit(0)
