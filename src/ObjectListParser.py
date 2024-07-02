import math
import os
from typing import List

import pandas as pd


class ObjectListParser:
    """
    解析 object_list 并将其转换为 Pandas DataFrame 的类。
    """

    def __init__(self, object_list: List[List[List[str]]],
                 columns: List[str] = ["时间", "昵称", "UID", "内容"]):
        """
        初始化 ObjectListParser 类。

        :param object_list: 需要解析的对象列表
        :param columns: DataFrame 的列名
        """
        self.object_list = object_list
        self.columns = columns

    def to_dataframe(self) -> pd.DataFrame:
        """
        将 object_list 转换为 Pandas DataFrame。

        :return: 转换后的 DataFrame
        :rtype: pd.DataFrame
        """
        # 创建一个空的 DataFrame
        df = pd.DataFrame(columns=self.columns)

        # 遍历 object_list，将数据添加到 DataFrame 中
        for obj in self.object_list:
            temp_df = self._convert_to_df(obj)
            # 将临时 DataFrame 添加到主 DataFrame 中
            df = pd.concat([df, temp_df], ignore_index=True)

        return df

    def to_single_dataframe(self, index: int) -> pd.DataFrame:
        """
        将 object_list 中的单个元素转换为 Pandas DataFrame。

        :param index: object_list 的索引
        :return: 转换后的 DataFrame
        :rtype: pd.DataFrame
        """
        if index < 0 or index >= len(self.object_list):
            raise IndexError("索引超出范围")

        obj = self.object_list[index]
        return self._convert_to_df(obj)

    def to_dataframe_list(self) -> List[pd.DataFrame]:
        """
        将 object_list 中的每个元素转换为 Pandas DataFrame，并返回 DataFrame 列表。

        :return: DataFrame 列表
        :rtype: List[pd.DataFrame]
        """
        dataframe_list = []
        for obj in self.object_list:
            dataframe_list.append(self._convert_to_df(obj))
        return dataframe_list

    def _convert_to_df(self, obj: List[List[str]]) -> pd.DataFrame:
        """
        辅助方法：将单个消息对象转换为 Pandas DataFrame。

        :param obj: 单个消息对象
        :return: 转换后的 DataFrame
        :rtype: pd.DataFrame
        """
        if len(obj) < 1:
            raise ValueError("每个消息对象必须至少包含一个子列表")

        # 移除空子列表
        filtered_obj = [col for col in obj if col]

        # 动态创建临时 DataFrame
        temp_data = {self.columns[i]: filtered_obj[i] for i in range(len(filtered_obj))}
        return pd.DataFrame(temp_data)


    def export_to_file(self, dfs, file_path, sheet_names=None, max_rows_per_sheet=1000000, max_sheets_per_file=255):
        """
        将 DataFrame 导出为 Excel 或 CSV 文件。如果行数超过 max_rows_per_sheet，则将超出部分分到新的工作表中。
        如果用户输入多个 DataFrame，则作为多张工作表导出为 Excel，同时 sheet_names 也需要是一个列表。
        每个 .xlsx 文件中的工作表数量不超过 max_sheets_per_file，且不超过 255。

        :param dfs: 需要导出的 DataFrame 或 DataFrame 列表
        :param file_path: 导出文件的路径
        :param sheet_names: 工作表名称或名称列表
        :param max_rows_per_sheet: 每个工作表的最大行数，默认为 1000000，不得超过 1048576
        :param max_sheets_per_file: 每个 .xlsx 文件中的最大工作表数量，默认为 255，不得超过 255
        """
        # 检查每个工作表的行数限制是否在有效范围内
        if max_rows_per_sheet > 1048576 or max_rows_per_sheet < 1:
            raise ValueError("max_rows_per_sheet 不能超过 1048576 或 小于 1")
        # 检查每个文件中工作表数量的限制是否在有效范围内
        if max_sheets_per_file > 255 or max_sheets_per_file < 1:
            raise ValueError("max_sheets_per_file 不能超过 255 或 小于 1")

        # 如果输入的 dfs 不是列表，则将其转换为列表
        if not isinstance(dfs, list):
            dfs = [dfs]

        # 如果未指定 sheet_names，则默认为每个 DataFrame 一个名称
        if sheet_names is None:
            sheet_names = ["Sheet" + str(i + 1) for i in range(len(dfs))]
        # 如果 sheet_names 不是列表，则将其转换为包含单个名称的列表
        elif not isinstance(sheet_names, list):
            sheet_names = [sheet_names]

        # 检查 DataFrames 和 sheet_names 的数量是否一致
        if len(dfs) != len(sheet_names):
            raise ValueError("DataFrames 的数量和 sheet_names 的数量必须一致")

        # 判断文件类型并进行相应的导出操作
        if file_path.endswith('.xlsx'):
            format = 'xlsx'
        elif file_path.endswith('.csv') and len(dfs) == 1:
            format = 'csv'
        else:
            raise ValueError("不支持的文件格式或多个 DataFrame 仅支持 .xlsx 格式文件")

        # 如果是 xlsx 格式文件
        if format == 'xlsx':
            # 计算需要多少个文件
            num_files = math.ceil(sum([math.ceil(len(df) / max_rows_per_sheet / max_sheets_per_file) for df in dfs]))

            # 获取 ExcelWriter 相关
            file_name = os.path.splitext(file_path)[0]
            file_index = 0


            writer = None
            def next_writer():
                nonlocal writer
                nonlocal file_index
                if writer is not None:
                    writer.close()
                if file_index < num_files:
                    file_name_suffix = f"_{file_index}" if file_index > 0 else ""
                    writer = pd.ExcelWriter(f"{file_name}{file_name_suffix}.xlsx")
                    file_index += 1

            # 遍历每个 DataFrame 和对应的 sheet_name
            sheet_index = 0
            for df, sheet_name in zip(dfs, sheet_names):
                next_writer()
                num_rows = len(df)
                num_sheets = math.ceil(num_rows / max_rows_per_sheet)
                start_row = 0
                while True:
                    end_row = min(start_row + max_rows_per_sheet, num_rows)
                    sheet_df = df.iloc[start_row:end_row]

                    # 写入文件
                    sheet_name_suffix = f"_{sheet_index + 1}" if num_sheets > 1 else ""
                    sheet_df.to_excel(writer, sheet_name=f"{sheet_name}{sheet_name_suffix}", index=False)

                    sheet_index += 1
                    if end_row == num_rows:
                        break
                    start_row = end_row

                    # 处理当前文件用尽的情况
                    if sheet_index == max_sheets_per_file:
                        sheet_index = 0
                        next_writer()

            # 关闭最后一个 Writer
            if writer is not None:
                writer.close()

        elif format == 'csv':
            # 对于 CSV 格式，将 DataFrame 写入单个文件中的多个工作表
            df = dfs[0]
            num_sheets = math.ceil(len(df) // max_rows_per_sheet) + 1
            for i in range(num_sheets):
                start_row = i * max_rows_per_sheet
                end_row = (i + 1) * max_rows_per_sheet
                sheet_df = df.iloc[start_row:end_row]
                # 在追加模式下写入文件，确保每个工作表都正确写入
                sheet_df.to_csv(file_path, mode='a' if i > 0 else 'w', index=False, header=i == 0)

    def export_dfs_to_excel(self, dfs: List[pd.DataFrame], file_path: str, sheet_names: List[str],
                            max_rows_per_sheet: int = 1000000, max_sheets_per_file: int = 255):
        """
        将多个 DataFrame 导出为一个或多个 Excel 文件中的多个工作表。
        如果行数超过 max_rows_per_sheet，则将超出部分分到新的工作表中。
        如果工作表数量超过 max_sheets_per_file，则创建新的文件。

        :param dfs: 需要导出的 DataFrame 列表
        :param file_path: 导出文件的路径
        :param sheet_names: 工作表名称列表
        :param max_rows_per_sheet: 每个工作表的最大行数，默认为 1000000
        :param max_sheets_per_file: 每个文件中的最大工作表数量，默认为 255
        """
        if not file_path.endswith('.xlsx'):
            raise ValueError("文件格式必须为 .xlsx")

        if len(dfs) != len(sheet_names):
            raise ValueError("DataFrame 列表和工作表名称列表长度必须一致")

        # 基本文件名和扩展名
        base_file_name, file_extension = os.path.splitext(file_path)
        file_index = 0
        sheet_index = 0

        writer = None

        def next_writer():
            nonlocal writer
            nonlocal file_index
            if writer is not None:
                writer.close()
            file_name_suffix = f"_{file_index}" if file_index > 0 else ""
            writer = pd.ExcelWriter(f"{base_file_name}{file_name_suffix}{file_extension}")
            file_index += 1

        next_writer()

        for df, sheet_name in zip(dfs, sheet_names):
            num_rows = len(df)
            num_sheets = math.ceil(num_rows / max_rows_per_sheet)
            start_row = 0

            for sheet_num in range(num_sheets):
                if sheet_index >= max_sheets_per_file:
                    next_writer()
                    sheet_index = 0

                end_row = min(start_row + max_rows_per_sheet, num_rows)
                sheet_df = df.iloc[start_row:end_row]

                sheet_name_suffix = f"_{sheet_num + 1}" if num_sheets > 1 else ""
                sheet_df.to_excel(writer, sheet_name=f"{sheet_name}{sheet_name_suffix}", index=False)

                start_row = end_row
                sheet_index += 1

        if writer is not None:
            writer.close()

# if __name__ == '__main__':
#     # 示例使用方法
#     object_list = [
#         [
#             ["2024-06-01 12:00:00", "2024-06-01 12:05:00", "2024-06-01 12:05:00", "2024-06-01 12:05:00", "2024-06-01 12:05:00", "2024-06-01 12:05:00", "2024-06-01 12:05:00", "2024-06-01 12:05:00", "2024-06-01 12:05:00", "2024-06-01 12:05:00", "2024-06-01 12:05:00"],  # time_list for first message object
#             [f"昵称{i * 2}", f"昵称{i * 2 + 1}", f"昵称{i * 2 + 2}", f"昵称{i * 2 + 3}", f"昵称{i * 2 + 4}", f"昵称{i * 2 + 5}", f"昵称{i * 2 + 6}", f"昵称{i * 2 + 7}", f"昵称{i * 2 + 8}", f"昵称{i * 2 + 9}", f"昵称{i * 2 + 10}"],  # name_list for first message object
#             [f"UID{i * 2}", f"UID{i * 2 + 1}", f"UID{i * 2 + 2}", f"UID{i * 2 + 3}", f"UID{i * 2 + 4}", f"UID{i * 2 + 5}", f"UID{i * 2 + 6}", f"UID{i * 2 + 7}", f"UID{i * 2 + 8}", f"UID{i * 2 + 9}", f"UID{i * 2 + 10}"],  # uid_list for first message object
#             [f"内容{i * 2}", f"内容{i * 2 + 1}", f"内容{i * 2 + 2}", f"内容{i * 2 + 3}", f"内容{i * 2 + 4}", f"内容{i * 2 + 5}", f"内容{i * 2 + 6}", f"内容{i * 2 + 7}", f"内容{i * 2 + 8}", f"内容{i * 2 + 9}", f"内容{i * 2 + 10}"],  # uid_list for first message object
#         ] for i in range(51)
#     ]
#
#     # 创建解析器对象
#     # parser = ObjectListParser(object_list)
#
#     # 转换整个 object_list 为 DataFrame
#     # df = parser.to_dataframe()
#     # print("整个 object_list 转换后的 DataFrame:")
#     # print(df)
#
#     # 导出整个 object_list 转换后的 DataFrame 到 Excel 文件
#     # parser.export_to_file(df, "output.xlsx", "全部消息")
#
#     # 导出单个消息对象转换后的 DataFrame 到 Excel 文件
#     # parser.export_to_file([df] * 5, "output.xlsx", [f"{i + 1}" for i in range(5)])
#
#     # 导出单个消息对象转换后的 DataFrame 到 Excel 文件(多工作簿)
#     # parser.export_dfs_to_excel([df] * 5, "output.xlsx", [f"{i + 1}" for i in range(5)])