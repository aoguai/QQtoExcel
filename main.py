import os
import sys
import argparse

from QQtoExcel import QQtoExcel

WORKDIR = os.path.dirname(os.path.realpath(sys.argv[0]))

if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument("qq_chat_route", help="导出的QQ聊天记录文件路径")
    parser.add_argument("--file_path", "-f", default=WORKDIR + "\\out\\", help="保存转换后的Excel表格目录")
    parser.add_argument("--sheet_name", "-sn", default="[消息对象]",
                        help="工作表的名字(留空则使用[消息对象]，同时可用的参数：[消息分组]、[消息对象]、[序号]作为规则宏)")
    parser.add_argument("--rule_string", "-rs", default="",
                        help="自定义文件名规则(留空则使用默认规则，同时可用的参数：[消息分组]、[消息对象]、[序号]作为规则宏)")
    parser.add_argument("--time_list_out", "-to", action="store_false", default=True, help="是否导出时间（时间、昵称、UID、内容不得全为否）")
    parser.add_argument("--name_list_out", "-no", action="store_false", default=True, help="是否导出昵称（时间、昵称、UID、内容不得全为否）")
    parser.add_argument("--uid_list_out", "-uo", action="store_false", default=True, help="是否导出UID（时间、昵称、UID、内容不得全为否）")
    parser.add_argument("--cont_list_out", "-co", action="store_false", default=True, help="是否导出内容（时间、昵称、UID、内容不得全为否）")
    parser.add_argument("--cont_nil_out", "-cno", action="store_true", default=False,
                        help="是否过滤无意义内容（仅在导出内容时候起作用）")
    parser.add_argument("--time_row_text", "-tt", default="时间", help="时间列表头的标题")
    parser.add_argument("--name_row_text", "-nt", default="昵称", help="昵称列表头的标题")
    parser.add_argument("--uid_row_text", "-ut", default="QQ（邮箱）", help="UID列表头的标题")
    parser.add_argument("--cont_row_text", "-ct", default="内容", help="内容列表头的标题")
    parser.add_argument("--multi_sheet_export", "-ms", type=int, default=0,
                        help="多工作表导出模式（0：非多工作表导出, 1：按消息对象导出, 2：按消息分组导出）")
    parser.add_argument("--out_type", "-ot", type=int, default=0, help="导出模式（0为按好友导出，非0按消息分组导出）")

    args = parser.parse_args()

    qe = QQtoExcel(qq_chat_route=args.qq_chat_route, file_path=args.file_path, sheet_name=args.sheet_name,
                   rule_string=args.rule_string, time_list_out=args.time_list_out,
                   name_list_out=args.name_list_out, uid_list_out=args.uid_list_out, cont_list_out=args.cont_list_out,
                   cont_nil_out=args.cont_nil_out, multi_sheet_export=args.multi_sheet_export,
                   time_row_text=args.time_row_text,
                   name_row_text=args.name_row_text, uid_row_text=args.uid_row_text, cont_row_text=args.cont_row_text,
                   out_type=args.out_type)
    qe.toExcel()
