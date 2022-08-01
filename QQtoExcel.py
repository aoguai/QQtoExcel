import openpyxl
from tqdm import tqdm
from init import *


class QQtoExcel:
    """
    parameters
    ----------
        QQChat_route : str
            QQ聊天记录路径
        file_path : str
            导出目录
        sheet_name : str
            工作表的名字
        time_list_out : bool
            是否导出时间
        name_list_out : bool
            是否导出昵称
        uid_list_out : bool
            是否导出Uid
        cont_list_out : bool
            是否导出内容
        time_row_text : str
            "时间"列表头的标题
        name_row_text : str
            "昵称"列表头的标题
        uid_row_text : str
            "QQ（邮箱）"列表头的标题
        cont_row_text : str
            "内容"列表头的标题
        out_type : int
            导出模式，0为按好友导出，非0按分组导出
    """
    QQChat_route = os.path.join(workdir, '全部消息记录.txt')  # 默认QQ聊天记录路径
    file_path = workdir + "\\out\\"  # 默认导出目录
    sheet_name = "test"  # 默认工作表的名字
    row = {}  # 表头字典
    row_key_list = []  # 表头标题列表，用以确保输出有序
    time_list_out = True  # 是否导出时间
    name_list_out = True  # 是否导出昵称
    uid_list_out = True  # 是否导出UID
    cont_list_out = True  # 是否导出内容
    out_type = 0  # 导出模式，0为按好友导出，非0按分组导出

    def __init__(self, QQChat_route, file_path, sheet_name, time_list_out, name_list_out, uid_list_out, cont_list_out,
                 time_row_text="时间", name_row_text="昵称", uid_row_text="QQ（邮箱）", cont_row_text="内容", out_type=0):
        self.QQChat_route = QQChat_route
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.time_list_out = time_list_out
        self.name_list_out = name_list_out
        self.uid_list_out = uid_list_out
        self.cont_list_out = cont_list_out
        self.out_type = out_type

        # 写入Excel标题
        # 默认，{"time_list": "时间", "name_list": "昵称", "uid_list": "QQ（邮箱）", "cont_list": "内容"}
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
        object_file_name_list = []  # 文件名称列表（分组_备注）
        object_list = []  # 消息对象列表

        time_list = []  # 消息时间列表
        name_list = []  # 昵称列表
        uid_list = []  # QQ or 邮箱列表
        cont_list = []  # 内容列表

        f = open(self.QQChat_route, encoding="utf-8")
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
            # 获取 保存文件名。格式：分组_昵称
            if self.out_type == 0:
                object_file_name_list.append(data_win_file(data_clean(
                    result[2 + (4 * i)].replace('\n', '').replace('\r', '').replace(' ', '').replace('消息分组:',
                                                                                                     ''))) + "_" + data_win_file(
                    data_clean(
                        result[3 + (4 * i)].replace('\n', '').replace('\r', '').replace(' ', '').replace('消息对象:', ''))))
            else:
                out_z_path_name = data_win_file(data_clean(
                    result[2 + (4 * i)].replace('\n', '').replace('\r', '').replace(' ', '').replace('消息分组:', '')))
                out_z_path = os.path.join(self.file_path, out_z_path_name)
                if not os.path.exists(out_z_path):
                    # print(out_z_path)
                    os.mkdir(out_z_path)  # 创建分组目录
                object_file_name_list.append(out_z_path_name + "\\" + data_win_file(
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
                    if self.time_list_out:
                        time_list.append(j[0])
                    if self.name_list_out:
                        name_list.append(data_clean(j[1].replace(j[2], '').replace('()', '').replace('<>', '')))
                    if self.uid_list_out:
                        uid_list.append(j[2])
                    if self.cont_list_out:
                        cont_list.append(data_clean(j[3][1:]))
            else:
                m = pattern2.findall(result[4 + (4 * i)])
                if len(m) >= 0:
                    for j in m:
                        if self.time_list_out:
                            time_list.append(j[0])
                        if self.name_list_out:
                            name_list.append(data_clean(j[1].replace(j[2], '').replace('()', '').replace('<>', '')))
                        if self.uid_list_out:
                            uid_list.append('')
                        if self.cont_list_out:
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
    def toExcel(self):
        object_file_name_list, object_list = self.get_QQChat_record()
        if len(self.row) <= 0:
            print("请选择至少一项输出内容")
            return

        files_path = []  # 输出目录列表
        for i in range(len(object_file_name_list)):
            files_path.append(os.path.join(self.file_path, object_file_name_list[i] + '.xls'))

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

            workboot.save(files_path[i])
            workboot.close()
