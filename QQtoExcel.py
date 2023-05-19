import openpyxl
from tqdm import tqdm
from init import *


class QQtoExcel:
    """
    将QQ聊天记录导出为Excel文件的类。

    """

    def __init__(self, qq_chat_route: str = os.path.join(WORKDIR, '全部消息记录.txt'), file_path: str = WORKDIR + "\\out\\",
                 sheet_name: str = "test", time_list_out: bool = True, name_list_out: bool = True,
                 uid_list_out: bool = True, cont_list_out: bool = True,
                 cont_nil_out: bool = False,
                 time_row_text: str = "时间", name_row_text: str = "昵称", uid_row_text: str = "QQ（邮箱）",
                 cont_row_text: str = "内容", out_type: int = 0):
        """
        初始化QQtoExcel类。

         :param qq_chat_route: QQ聊天记录路径
         :type qq_chat_route: str
         :param file_path: 导出目录
         :type file_path: str
         :param sheet_name: 工作表的名字
         :type sheet_name: str
         :param time_list_out: 是否导出时间
         :type time_list_out: bool
         :param name_list_out: 是否导出昵称
         :type name_list_out: bool
         :param uid_list_out: 是否导出Uid
         :type uid_list_out: bool
         :param cont_list_out: 是否导出内容
         :type cont_list_out: bool
         :param cont_nil_out: 是否过滤无意义内容
         :type cont_nil_out: bool
         :param time_row_text: "时间"列表头的标题
         :type time_row_text: str
         :param name_row_text: "昵称"列表头的标题
         :type name_row_text: str
         :param uid_row_text: "QQ（邮箱）"列表头的标题
         :type uid_row_text: str
         :param cont_row_text: "内容"列表头的标题
         :type cont_row_text: str
         :param out_type: 导出模式，0为按好友导出，非0按分组导出
         :type out_type: int
        """

        self.qq_chat_route = qq_chat_route
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.time_list_out = time_list_out
        self.name_list_out = name_list_out
        self.uid_list_out = uid_list_out
        self.cont_list_out = cont_list_out
        self.cont_nil_out = cont_nil_out
        self.out_type = out_type
        self.row = {}  # 表头字典
        self.row_key_list = []  # 表头标题列表，用以确保输出有序
        if self.time_list_out:
            self.row['time_list'] = time_row_text
            self.row_key_list.append('time_list')
        if self.name_list_out:
            self.row['name_list'] = name_row_text
            self.row_key_list.append('name_list')
        if self.uid_list_out:
            self.row['uid_list'] = uid_row_text
            self.row_key_list.append('uid_list')
        if self.cont_list_out:
            self.row['cont_list'] = cont_row_text
            self.row_key_list.append('cont_list')

    def get_QQChat_record(self):
        """
        读取 QQ 聊天记录并生成对象。

        :return: 生成的对象
        :rtype: generator
        """
        object_file_name_list = []  # 文件名称列表（分组_备注）
        object_list = []  # 消息对象列表

        time_list = []
        name_list = []
        uid_list = []
        cont_list = []

        with open(self.qq_chat_route, encoding="utf-8") as f:
            text = f.read()

        pattern = r"消息分组:\s*(.+)[\s\S]*?消息对象:\s*(.+)(?:[=\n\r]*((?:\d{1,4}-(?:1[0-2]|0?[1-9])-(?:0?[1-9]|[1-2]\d|30|31)) (?:(?:[0-1]\d|2[0-3]|[0-9]):[0-5]\d:[0-5]\d))[ ]?(.*)[\n\r]([\s\S]*?)[\n\r][\n\r])+"
        message_pattern = re.compile(
            r'(\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{1,2}:\d{1,2})(.+?([(|<](.*)[)|>])?\n)([\s\S]*?)(?=\n\s*(\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{1,2}:\d{1,2})|$)')

        for match in re.finditer(pattern, text):
            if len(match.group(2)) < 42:
                if self.out_type == 0:
                    object_file_name_list.append(
                        data_win_file(data_clean(match.group(1)) + "_" + data_clean(match.group(2))))
                else:
                    out_z_path_name = data_win_file(data_clean(match.group(1)))
                    out_z_path = os.path.join(self.file_path, out_z_path_name)
                    if not os.path.exists(out_z_path):
                        os.mkdir(out_z_path)
                    object_file_name_list.append(out_z_path_name + "\\" + data_win_file(data_clean(match.group(2))))

                match_text = match.group()
                match_text_list = message_pattern.findall(match_text)
                if match_text_list:
                    for j in match_text_list:
                        if self.cont_list_out:
                            cont_text = data_clean(j[4])
                            if self.cont_nil_out:
                                cont_text = re.sub(f'\[(图片|语音|表情|QQ红包)]', "", cont_text)
                            if len(cont_text) > 0 or not self.cont_nil_out:
                                cont_list.append(cont_text)
                            else:
                                continue
                        if self.time_list_out:
                            time_list.append(j[0])
                        if self.name_list_out:
                            name_list.append(
                                data_clean(j[1].replace(j[2], '').replace('\\n', '')))
                        if self.uid_list_out:
                            uid_list.append(j[3])
                        else:
                            print("如果您选择了过滤无意义内容，则不能不导出内容")
                            break
                object_list.append([time_list, name_list, uid_list, cont_list])
                # 清空列表
                time_list = []
                name_list = []
                uid_list = []
                cont_list = []

        return object_file_name_list, object_list

    # 创建Excel文件
    def toExcel(self):
        """
        创建Excel文件并导出数据。

        :return: None
        :rtype: None
        """

        object_file_name_list, object_list = self.get_QQChat_record()
        if len(self.row) <= 0:
            print("请选择至少一项输出内容")
            return

        files_path = []  # 输出目录列表
        for i in range(len(object_file_name_list)):
            files_path.append(os.path.join(self.file_path, object_file_name_list[i] + '.xlsx'))

        for i in tqdm(range(len(files_path)), '导出中'):
            time_list = object_list[i][0]
            name_list = object_list[i][1]
            uid_list = object_list[i][2]
            cont_list = object_list[i][3]

            # 创建workbook和sheet对象
            workboot = openpyxl.Workbook()
            worksheet = workboot.active
            worksheet.title = self.sheet_name  # 设置工作表的名字

            # 写入表头
            for j in range(len(self.row_key_list)):
                worksheet.cell(1, j + 1, self.row[self.row_key_list[j]])

            # 写入内容
            if self.time_list_out:
                for k in range(len(time_list)):
                    worksheet.cell(k + 2, self.row_key_list.index("time_list") + 1, time_list[k])
            if self.name_list_out:
                for k in range(len(name_list)):
                    worksheet.cell(k + 2, self.row_key_list.index("name_list") + 1, name_list[k])
            if self.uid_list_out:
                for k in range(len(uid_list)):
                    worksheet.cell(k + 2, self.row_key_list.index("uid_list") + 1, uid_list[k])
            if self.cont_list_out:
                for k in range(len(cont_list)):
                    worksheet.cell(k + 2, self.row_key_list.index("cont_list") + 1, cont_list[k])

            # 保存Excel文件
            with open(files_path[i], 'wb') as file:
                workboot.save(file)
