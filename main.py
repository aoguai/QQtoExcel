import os
import argparse
import sys

from QQtoExcel import QQtoExcel

workdir = os.path.dirname(os.path.realpath(sys.argv[0]))

if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument("QQChat_route", help="导出的QQ聊天记录文件路径")
    parser.add_argument("--file_path", "-f", default=workdir + "\\out\\", help="保存转换后的Excel表格目录")
    parser.add_argument("--sheet_name", "-sn", default="test", help="工作表的名字")
    parser.add_argument("--time_list_out", "-to", action="store_false", default=True, help="是否导出时间（时间、昵称、UID、内容不得全为否）")
    parser.add_argument("--name_list_out", "-no", action="store_false", default=True, help="是否导出昵称（时间、昵称、UID、内容不得全为否）")
    parser.add_argument("--uid_list_out", "-uo", action="store_false", default=True, help="是否导出UID（时间、昵称、UID、内容不得全为否）")
    parser.add_argument("--cont_list_out", "-co", action="store_false", default=True, help="是否导出内容（时间、昵称、UID、内容不得全为否）")
    parser.add_argument("--time_row_text", "-tt", default="时间", help="时间列表头的标题")
    parser.add_argument("--name_row_text", "-nt", default="昵称", help="昵称列表头的标题")
    parser.add_argument("--uid_row_text", "-ut", default="QQ（邮箱）", help="UID列表头的标题")
    parser.add_argument("--cont_row_text", "-ct", default="内容", help="内容列表头的标题")
    parser.add_argument("--out_type", "-ot", type=int, default=0, help="导出模式，0为按好友导出，非0按分组导出")

    args = parser.parse_args()

    qe = QQtoExcel(QQChat_route=args.QQChat_route, file_path=args.file_path, sheet_name=args.sheet_name, time_list_out=args.time_list_out,
                   name_list_out=args.name_list_out, uid_list_out=args.uid_list_out, cont_list_out=args.cont_list_out, time_row_text=args.time_row_text,
                   name_row_text=args.name_row_text, uid_row_text=args.uid_row_text, cont_row_text=args.cont_row_text, out_type=args.out_type)
    qe.toExcel()

