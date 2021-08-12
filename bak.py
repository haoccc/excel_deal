import math
import time
import pandas as pd
import os
import numpy as np
import xlsxwriter


def pd_toexcel(data, filename):  # pandas库储存数据到excel
    dfData = {  # 用字典设置DataFrame所需数据
        '图号-孔数': data[0],
        '标签内容': data[1],
        '长*宽=数量': data[2]
    }
    # if not os.path.exists(filename):
    #     os.mkdir(filename)
    df = pd.DataFrame(dfData)  # 创建DataFrame
    df.to_excel(filename, index=False)  # 存表，去除原始索引列（0,1,2...）


class change_table():
    pass


class ResultTable():
    pass


class CaculateTable():
    def __init__(self, file_path):
        self.data = extract_data(file_path)
        self.length_data = len(self.data[0])
        self.data[2] = [" " if math.isnan(self.data[2][i]) else str(self.data[2][i]) for i in range(self.length_data)]
        self.data[3] = [" " if math.isnan(self.data[3][i]) else str(self.data[3][i]) for i in range(self.length_data)]
        # self.file_path = file_path
        self.small_kong, self.big_kong, self.small_corner, self.big_corner = cal_kong_angle(self.data)
        self.single_price = [150 if i != " " else 100 for i in self.data[8]]
        self.total_price = [0 for i in range(self.length_data)]
        self.title = ['序', '长边尺', '宽边尺', 'A', 'B', '图号', '数量', '孔', '大孔', '小孔', '标签内容', '备注', '单价', '单位面积', '除加工费价格',
                      '单件小圆角个数',
                      '单件大圆角个数', '总价']
        self.sequence_number = [i + 1 for i in range(self.length_data)]

        # dfData = {  # 用字典设置DataFrame所需数据
        #     '序': self.sequence_number,
        #     '长边尺': data[0],
        #     '宽边尺': data[1],
        #     'A': data[2],
        #     'B': data[3],
        #     '图号': data[4],
        #     '数量': data[5],
        #     '孔': data[6],
        #     '大孔': self.big_kong,
        #     '小孔': self.small_kong,
        #     '标签内容': data[7],
        #     '备注': data[8],
        #     '单价': self.single_price,
        #     '单件小圆角个数': self.small_corner,
        #     '单件大圆角个数': self.big_corner,
        #     '总价': self.total_price
        # }
        # if not os.path.exists(filename):
        #     os.mkdir(filename)
        # df = pd.DataFrame(dfData)  # 创建DataFrame
        # df.to_excel(self.file_path, index=False)  # 存表，去除原始索引列（0,1,2...）

    def write_with_formula(self, file_name):
        random_list = [0 for i in range(self.length_data)]  # 填充空位
        data = [(self.sequence_number[i], self.data[0][i], self.data[1][i], self.data[2][i], self.data[3][i],
                 self.data[4][i], self.data[5][i], self.data[6][i], self.big_kong[i], self.small_kong[i],
                 self.data[7][i],
                 self.data[8][i], self.single_price[i], random_list[i],
                 random_list[i], self.small_corner[i], self.big_corner[i], self.total_price[i])
                for i
                in range(self.length_data)]  # 数据转成一行一行的，方便写入

        # print(data)
        workbook = xlsxwriter.Workbook(file_name)
        worksheet = workbook.add_worksheet("Sheet1")
        worksheet.write_row(0, 0, self.title)  # 写入title
        for row in range(1, self.length_data + 1):
            worksheet.write_row(row, 0, data[row - 1])
            formula_area = "if(B{}*C{}/1000000<=0.1, 0.1, B{}*C{}/1000000)".format(row + 1, row + 1, row + 1,
                                                                                   row + 1)  # 面积公式
            formula_price1 = "=M{}*N{}*G{}".format(row + 1, row + 1, row + 1)
            formula_price_all = "=round(O{}+(J{}+5*I{}+0.5*P{}+Q{})*G{},2)".format(row + 1, row + 1, row + 1, row + 1,
                                                                                   row + 1, row + 1)
            # print(formula_x)
            worksheet.write_formula(row, 13, formula_area)  # 写入面积
            worksheet.write_formula(row, 14, formula_price1)  # 写入除加工费价格
            worksheet.write_formula(row, 17, formula_price_all)  # 写入总价
        workbook.close()


def cal_kong_angle(data):
    """
    计算大孔、小孔、大小圆角数量
    :param data:
    :return:
    """
    little_kong = []
    for i in data[6]:
        # print(int(i))
        if i == ' ':
            little_kong.append(0)
        elif int(i) == 17:
            little_kong.append(19)
        elif int(i) == 13:
            little_kong.append(10)
        else:
            little_kong.append(int(i))

    big_kong = []
    for i in data[6]:
        # print(int(i))
        if i == ' ':
            big_kong.append(0)
        elif int(i) == 17:
            big_kong.append(8)
        elif int(i) == 13:
            big_kong.append(3)
        else:
            big_kong.append(int(i))

    big_corner = [0 for x in range(len(data[0]))]
    small_corner = [4 for x in range(len(data[0]))]

    return little_kong, big_kong, small_corner, big_corner


def change_origin(data, filename):
    """
    修改原始表格
    :param data:
    :param filename:
    :return:
    """
    sequence_number = [i + 1 for i in range(len(data[0]))]
    dfData = {  # 用字典设置DataFrame所需数据
        '序': sequence_number,
        '长边尺': data[0],
        '宽边尺': data[1],
        'A': data[2],
        'B': data[3],
        '图号': data[4],
        '数量': data[5],
        '孔': data[6],
        '标签内容': data[7],
        '备注': data[8]
    }
    # if not os.path.exists(filename):
    #     os.mkdir(filename)
    df = pd.DataFrame(dfData)  # 创建DataFrame
    df.to_excel(filename, index=False)  # 存表，去除原始索引列（0,1,2...）


def extract_data(file):
    """
    提取表格里面的数据，只处理空值
    :param file:
    :return:
    """
    if pd.read_excel(file, header=None)[0][0] != "序号":  # 如果没有表头则从第一行读取
        df = pd.read_excel(file, header=None)
    else:  # 有表头
        df = pd.read_excel(file)

    data_length = df.iloc[:, 1].values  # 长
    data_width = df.iloc[:, 2].values  # 宽
    data_A = df.iloc[:, 3].values  # A
    data_B = df.iloc[:, 4].values  # B

    try:  # 如果这一列均为空
        data_id = df.iloc[:, 5].values  # 图号， 5 代表位于第6列
    except IndexError as e:
        data_id = ['' for i in range(len(data_id))]

    try:
        data_kong_count = df.iloc[:, 7].values  # 孔数量
    except IndexError as e:
        data_kong_count = [0 for i in range(len(data_id))]

    data_kong_count = ["%.0f" % i for i in data_kong_count]  # 转成字符串，没小数点

    data_count = df.iloc[:, 6].values  # 数量
    data_label = df.iloc[:, 8].values  # 标签内容

    try:
        data_remark = df.iloc[:, 9].values  # 备注
    except IndexError as e:
        data_remark = ['' for i in range(len(data_id))]

    data_remark = [str(i) for i in data_remark]  # 转成字符串，没小数点
    for i in range(len(data_id)):
        # 判断空值
        if data_kong_count[i] == "nan":
            data_kong_count[i] = " "
        if data_remark[i] == "nan":
            data_remark[i] = " "
        if data_A[i] == "nan":  # todo 有bug
            data_A[i] = " "
        if data_B[i] == "nan":
            data_B[i] = " "

    # 返回提取出的所有数据，不做任何处理
    return [data_length, data_width, data_A, data_B, data_id, data_count, data_kong_count, data_label, data_remark]


def get_data_for_result(data):
    """
    将提取出的数据用于-->结果
    :param file: 文件路径
    :return:
    """
    flag_change = [0 for i in range(len(data[0]))]
    for i in range(len(data[4])):
        # 去掉 THBL
        if "THBL-" in data[4][i]:
            data[4][i] = data[4][i].split("THBL-")[-1]

        # 交换长宽值
        if data[1][i] > data[0][i]:
            flag_change[i] = 1
            data[1][i], data[0][i] = data[0][i], data[1][i]

    data = sort_data(data)

    data_length = data[0]  # 长
    data_width = data[1]  # 宽
    data_A = data[2]  # A
    data_B = data[3]  # B
    data_id = data[4]  # 图号， 5 代表位于第6列
    data_count = data[5]  # 数量
    data_kong_count = data[6]  # 孔数量
    data_label = data[7]  # 标签内容
    data_remark = data[8]  # 备注

    result = []
    data_0 = [str(data_id[i]) + "-" + data_kong_count[i] + data_remark[i] for i in range(len(data_id))]
    # print("data_0", data_0)
    result.append(data_0)
    result.append(data_label)
    data_2 = [(str(data_length[i]) + "×" + str(data_width[i]) + "=" + str(data_count[i])) if flag_change[i] else (
                str(data_length[i]) + "*" + str(data_width[i]) + "=" + str(data_count[i])) for i in
              range(len(data_count))]
    result.append(data_2)

    return result


def get_data_for_change(data):
    """
    画孔的单
    交换长宽->调整原来的表格
    :param data:
    :return:
    """
    flag_change = [0 for i in range(len(data[0]))]
    data_length = []
    data_width = []
    for i in range(len(data[4])):
        # 交换长宽值
        if data[1][i] > data[0][i]:
            data_length.append('×' + str(data[1][i]))
            data_width.append(data[0][i])
            data[1][i], data[0][i] = data[0][i], data[1][i]
        else:
            data_length.append(data[0][i])
            data_width.append(data[1][i])
    return sort_data(data, data_length, data_width)


def sort_data(data, data_length=None, data_width=None):
    """
    交换数据顺序
    :param data:
    :return:
    """
    change_data = [(data[0][i], i) for i in range(len(data[0]))]  # 用于确认交换位置
    change_data = np.array(change_data)
    change_data = change_data[np.lexsort(change_data[:, ::-1].T)]  # 第一列正序
    change_data = change_data[::-1]

    index_data = [i[1] for i in change_data]  # 数据的序号
    data[0] = [i[0] for i in change_data]
    data[1] = [data[1][i] for i in index_data]
    data[2] = [data[2][i] for i in index_data]
    data[3] = [data[3][i] for i in index_data]
    data[4] = [data[4][i] for i in index_data]
    data[5] = [data[5][i] for i in index_data]
    data[6] = [data[6][i] for i in index_data]
    data[7] = [data[7][i] for i in index_data]
    data[8] = [data[8][i] for i in index_data]

    if data_width and data_length:
        data_width = [data_width[i] for i in index_data]
        data_length = [data_length[i] for i in index_data]
        data[0] = data_length
        data[1] = data_width
    return data


if __name__ == '__main__':
    print("把要处理的文件放在订单这个文件夹里面，结果存在结果文件夹里面")
    if not os.path.exists('./输入订单'):
        os.mkdir('./输入订单')
    if not os.path.exists('./输出结果'):
        os.mkdir('./输出结果')
    if not os.path.exists('./画孔的单'):
        os.mkdir('./画孔的单')
    if not os.path.exists('./算价格的单'):
        os.mkdir('./算价格的单')

    file_list = os.listdir("./输入订单")
    files_path = [os.path.join("./输入订单", i) for i in file_list]
    result_path = os.getcwd() + "/输出结果/"
    files_save_path = [os.path.join("./输出结果", i) for i in file_list]
    files_change_path = [os.path.join("./画孔的单", i) for i in file_list]
    files_cal_path = [os.path.join("./算价格的单", i) for i in file_list]

    for i in range(len(files_path)):
        print("正在处理：", file_list[i])
        data = extract_data(files_path[i])
        data_bak = extract_data(files_path[i])
        pd_toexcel(get_data_for_result(data), files_save_path[i])
        change_origin(get_data_for_change(data_bak), files_change_path[i])
        test = CaculateTable(files_path[i])
        test.write_with_formula(files_cal_path[i])
    print("五秒后自动退出")
    print("处理完成！")
    time.sleep(5)
