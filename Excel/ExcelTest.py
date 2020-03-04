import xlrd
import xlwt
from xlutils.copy import copy
import pandas as pd
import matplotlib.pyplot as plt
from scipy.optimize import leastsq
import numpy as np


def open_sheet(workbook):
    while True:
        try:
            print("当前文档的sheet表格有：")
            # 获得当前excel文件所有sheet的名称
            sheets = workbook.sheet_names()
            for i in range(len(sheets)):
                print(str(i + 1) + "." + sheets[i])
            worksheet_choose = int(input("请选择要操作的sheet：")) - 1
            print(worksheet_choose)
            worksheet = workbook.sheet_by_index(worksheet_choose)
            print("您选择的表格为：")
            # 新建一个空数组
            array = []
            # 获取sheet的行数
            rows = worksheet.nrows
            # 将所有行数的值写到一个数组中
            for i in range(rows):
                array.append(worksheet.row_values(i))
            # 给数组添加标签
            for i in range(len(array)):
                for j in range(len(array[i])):
                    array[i][j] = float(array[i][j])
            table = pd.DataFrame(array)
            # 打印带标签的数组
            print(table)
            break
        except:
            print("输入有误，请重新输入！")
    return worksheet_choose, table


def change_excel(filename, workbook):
    worksheet_choose = open_sheet(workbook)[0]
    new_workbook = copy(workbook)
    new_worksheet = new_workbook.get_sheet(worksheet_choose)
    while True:
        try:
            cell = input("请输入你要修改的单元格行列号及修改内容（例如修改第0行，第0列为test请输入:0,0,test）,结束修改请输入n：")
            if cell == 'n':
                new_workbook.save(filename)
                print("修改完成！")
                break
            else:
                cells = cell.split(',')
                new_worksheet.write(int(cells[0]), int(cells[1]), cells[2])
        except:
            print("格式输入错误,请重新输入！")
    print("修改文件完成,请继续")
    new_workbook.save(filename)


def sum_excel(workbook):
    worksheet_choose = open_sheet(workbook)[0]
    worksheet = workbook.sheet_by_index(worksheet_choose)
    while True:
        try:
            row_col_choose = int(input("请输入选择对行还是列进行求和操作（1.行；2.列）："))
            sum = 0
            if row_col_choose == 1:
                row_choose = int(input("请选择要进行求和的行号："))
                row_values = worksheet.row_values(row_choose)
                for i in range(len(row_values)):
                    sum += float(row_values[i])
                print("第" + str(row_choose) + "行的和为：" + str(sum))
                break
            elif row_col_choose == 2:
                col_choose = int(input("请选择要进行求和的列号："))
                col_values = worksheet.col_values(col_choose)
                for i in range(len(col_values)):
                    sum += float(col_values[i])
                print("第" + str(col_choose) + "列的和为：" + str(sum))
                break
        except:
            print("输入有错误，请重新输入！")
    print("求和操作完成，请继续！")


def mean_excel(workbook):
    table = open_sheet(workbook)[1]
    while True:
        try:
            row_col_choose = int(input("请输入选择对行还是列进行求和操作（1.行；2.列；3.全部行均值；4.全部列均值）："))
            if row_col_choose == 1:
                row_choose = int(input("请选择要进行求均值的行号："))
                print("第" + str(row_choose) + "行的均值为：" + str(table.iloc[row_choose, :].mean()))
            elif row_col_choose == 2:
                col_choose = int(input("请选择要进行求均值的列号："))
                print("第" + str(col_choose) + "列的均值为：" + str(table.iloc[:, col_choose].mean()))
            elif row_col_choose == 3:
                print("所有行的均值为：\n" + str(table.mean(axis=1)))
            elif row_col_choose == 4:
                print("所有列的均值为：\n" + str(table.mean(axis=0)))
            break
        except:
            print("输入有错误，请重新输入！")
    print("均值操作完成，请继续！")


def draw_scatter(xArray, yArray):
    xArray1 = []
    yArray1 = []
    for i in range(len(xArray)):
        xArray1.append(float(xArray[i]))
    for i in range(len(yArray)):
        yArray1.append(float(yArray[i]))
    # 定义绘图图框样式
    plt.figure(figsize=(8, 8), facecolor='grey', frameon='False')
    # 绘制散点图
    plt.scatter(xArray1, yArray1, s=150, c='cyan', edgecolors='blue', marker='o', label='Simple Points', linewidth=3)
    # 添加X轴标记
    plt.xlabel('x')
    # 添加Y轴标记
    plt.ylabel('y')
    # 绘制图例
    plt.legend()
    # 展示图表
    plt.show()


def excel_draw_scatter(workbook):
    worksheet_choose = open_sheet(workbook)[0]
    worksheet = workbook.sheet_by_index(worksheet_choose)
    while True:
        try:
            row_col_choose = int(input("请输入选择对行还是列进行求和操作（1.行；2.列）："))
            if row_col_choose == 1:
                rows = input("请选择要进行求均值的行号,中间用逗号隔开(例如，第0行和第1行输入0,1)：").split(',')
                row_x = worksheet.row_values(int(rows[0]))
                row_y = worksheet.row_values(int(rows[1]))
                draw_scatter(row_x, row_y)
                break
            elif row_col_choose == 2:
                cols = input("请选择要进行求均值的列号,中间用逗号隔开(例如，第0列和第1列输入0,1)：").split(',')
                col_x = worksheet.col_values(int(cols[0]))
                col_y = worksheet.col_values(int(cols[1]))
                draw_scatter(col_x, col_y)
                break
        except:
            print("输入有错误，请重新输入！")
    print("散点图绘制操作完成，请继续！")


def draw_line(xArray, yArray):
    xArray1 = []
    yArray1 = []
    for i in range(len(xArray)):
        xArray1.append(float(xArray[i]))
    for i in range(len(yArray)):
        yArray1.append(float(yArray[i]))
    # 定义绘图图框样式
    plt.figure(figsize=(8, 8), facecolor='grey', frameon='False')
    # 添加X轴标记
    plt.xlabel('x')
    # 添加Y轴标记
    plt.ylabel('y')
    plt.plot(xArray1, xArray1, c='black', linestyle='--', label='Fitting line', linewidth=3)
    # 绘制图例
    plt.legend()
    # 展示图表
    plt.show()


def excel_draw_line(workbook):
    worksheet_choose = open_sheet(workbook)[0]
    worksheet = workbook.sheet_by_index(worksheet_choose)
    while True:
        try:
            row_col_choose = int(input("请输入选择对行还是列进行求和操作（1.行；2.列）："))
            if row_col_choose == 1:
                rows = input("请选择要进行求均值的行号,中间用逗号隔开(例如，第0行和第1行输入0,1)：").split(',')
                row_x = worksheet.row_values(int(rows[0]))
                row_y = worksheet.row_values(int(rows[1]))
                draw_line(row_x, row_y)
                break
            elif row_col_choose == 2:
                cols = input("请选择要进行求均值的列号,中间用逗号隔开(例如，第0列和第1列输入0,1)：").split(',')
                col_x = worksheet.col_values(int(cols[0]))
                col_y = worksheet.col_values(int(cols[1]))
                draw_line(col_x, col_y)
                break
        except:
            print("输入有错误，请重新输入！")
    print("折线图绘制操作完成，请继续！")


# 定义拟合函数
def func(factor, x):
    # 定义函数的两个参数
    k, b = factor
    # 返回拟合函数计算结果
    return k * x + b


# 定义误差函数
def error(factor, x, y, s):
    print(s)  # 打印参数估计次数
    return func(factor, x) - y  # 返回计算数组


# 最小二乘法拟合函数
def least_square(factor, xArray, yArray):
    try:
        s = '最小二乘法拟合参数估计'
        fit_result = leastsq(error, factor, args=(xArray, yArray, s))
        print("拟合完成！")
    except:
        print("拟合出错了,正在重新拟合！")
    return fit_result

# 定义绘图函数
def draw(fit_result, xArray, yArray):
    # 定义绘图图框样式
    plt.figure(figsize=(8, 8), facecolor='grey', frameon='False')
    # 绘制散点图
    plt.scatter(xArray, yArray, s=150, c='cyan', edgecolors='blue', marker='o', label='Simple Points', linewidth=3)
    # 添加X轴标记
    plt.xlabel('x')
    # 添加Y轴标记
    plt.ylabel('y')
    # 在1~100范围内随机生成100个随机数
    draw_x_array = np.linspace(0, 100, 100)
    # 获得一次函数两个参数
    k, b = fit_result[0]
    # 打印拟合函数
    print('y =', k, 'x + ', b)
    # 生成拟合函数
    draw_y_array = k * draw_x_array + b
    # 绘制拟合函数曲线
    plt.plot(draw_x_array, draw_y_array, c='black', linestyle='--', label='Fitting line', linewidth=3)
    # 绘制图例
    plt.legend()
    # 展示图表
    plt.show()


def least_squares(workbook):
    worksheet_choose = open_sheet(workbook)[0]
    worksheet = workbook.sheet_by_index(worksheet_choose)
    xArray = []
    yArray = []
    while True:
        try:
            row_col_choose = int(input("请输入选择对行还是列进行求和操作（1.行；2.列）："))
            if row_col_choose == 1:
                rows = input("请选择要进行求均值的行号,中间用逗号隔开(例如，第0行和第1行输入0,1)：").split(',')
                row_x = worksheet.row_values(int(rows[0]))
                row_y = worksheet.row_values(int(rows[1]))
                for i in range(len(row_x)):
                    xArray.append(float(row_x[i]))
                for i in range(len(row_y)):
                    yArray.append(float(row_y[i]))
            elif row_col_choose == 2:
                cols = input("请选择要进行求均值的列号,中间用逗号隔开(例如，第0列和第1列输入0,1)：").split(',')
                col_x = worksheet.col_values(int(cols[0]))
                col_y = worksheet.col_values(int(cols[1]))
                for i in range(len(col_x)):
                    xArray.append(float(col_x[i]))
                for i in range(len(col_x)):
                    yArray.append(float(col_y[i]))
            factor = [1, 1]
            fit_result = least_square(factor, xArray, yArray)
            draw(fit_result, xArray, yArray)
            break
        except:
            print("输入有错误，请重新输入！")
    print("最小二乘法拟合操作完成，请继续！")


def open_excel():
    while True:
        try:
            # 输入需要打开文件的文件名
            filename = input("请输入要打开文件的文件名：")
            # 判断输入的文件名是否有后缀
            if filename.endswith('.xls'):
                # 有后缀，使文件名保持当前文件名不变
                filename = filename
            else:
                # 不存在后缀，给文件名加上后缀
                filename = filename + '.xls'
            # 打开文件名为filename的xls文件
            workbook = xlrd.open_workbook(filename)
            # 获取excel文件中所有sheet的名称
            sheets = workbook.sheet_names()
            # 打印当前sheet的个数
            print(filename + "文件中共有" + str(len(sheets)) + "个sheet文件")
            # 打印所有的sheet的内容
            for i in range(len(sheets)):
                # 打印正在打开的sheet的名称
                print("第" + str(i + 1) + '个sheet的名称为：' + sheets[i] + "，内容为：")
                # 打开sheet
                worksheet = workbook.sheet_by_index(i)
                # 新建一个空数组
                array = []
                # 获取sheet的行数
                rows = worksheet.nrows
                # 将所有行数的值写到一个数组中
                for i in range(rows):
                    array.append(worksheet.row_values(i))
                # 给数组添加标签
                table = pd.DataFrame(array)
                # 打印带标签的数组
                print(table)
            print(filename + "文件打开完成，请继续操作")
            break
        except:
            print("打开失败，请检查文件名是否输入错误")
    while True:
        # 定义下一步可以进行的操作
        open_excel_operations = ['修改Excel文件内容', '求和', '求平均', '绘制散点图', '绘制折线图', '最小二乘法拟合', '我不想操作了，退出']
        # 打印出来所有可以进行的操作
        for j in range(len(open_excel_operations)):
            print(str(j + 1) + "." + open_excel_operations[j])
        try:
            open_excel_operations_choose = int(input("请输入序号进行下一的操作："))
            if open_excel_operations_choose == 1:
                change_excel(filename, workbook)
            elif open_excel_operations_choose == 2:
                sum_excel(workbook)
            elif open_excel_operations_choose == 3:
                mean_excel(workbook)
            elif open_excel_operations_choose == 4:
                excel_draw_scatter(workbook)
            elif open_excel_operations_choose == 5:
                excel_draw_line(workbook)
            elif open_excel_operations_choose == 6:
                least_squares(workbook)
            elif open_excel_operations_choose == 7:
                break
        except:
            print("输入的序号不符合要求，请重新输入！")


# 保存excel文件
def save_excel(filename, workbook):
    while True:
        try:
            workbook.save(filename)
            break
        except:
            filename = input("不合法的文件名，请重新输入：") + '.xls'


def write_sheet(filename, workbook, worksheet):
    # 定义初始输入行的序号
    row_number = 0
    while True:
        try:
            # 输入每行的的值，或者输入“n”退出输入
            row_value = input("输入第" + str(row_number + 1) + "行的值（不同单元格用,隔开，n结束输入）\n")
            # 输入“n”，退出程序
            if row_value == 'n':
                print("输入完成，请继续操作！")
                break
            # 输入其他内容，写入单元格
            else:
                # 以“,”分割字符串
                rows = row_value.split(',')
                # 将输入的值写入单元格
                for col_number in range(len(rows)):
                    worksheet.write(row_number, col_number, rows[col_number])
                # 输入行号加 1
                row_number += 1
        except:
            # 当前行输入错误，重新进行
            print("输入错误，请重新输入！")
    while True:
        # 定义可进行的操作数组
        write_sheet_operations = ["保存表格并进行其他操作", '不了，我要退出操作']
        # 展示可进行的操作
        for i in range(len(write_sheet_operations)):
            print(str(i + 1) + "." + write_sheet_operations[i])
        try:
            # 输入编号选择接下来进行的操作
            write_sheet_operations_choose = int(input("输入序号选择需要执行的操作："))
            # 判断输入的编号，进行下一步的操作
            if write_sheet_operations_choose == 1:
                # 保存表格
                save_excel(filename, workbook)
                # 下一步继续进行操作
                holden = 0
                # 跳出当前循环
                break
            # 退出程序
            elif write_sheet_operations_choose == 2:
                # 保存表格
                save_excel(filename, workbook)
                # 下一步催出操作
                holden = 1
                # 跳出当前循环
                break
        except:
            # 输入编号不符合要求，重新输入
            print("序号输入错误，请重试！")
    # 返回下一步操作
    return holden


def add_sheet(filename, workbook):
    while True:
        try:
            # 输入表格的名称
            sheet_name = input('请输入表格名称（输入exit结束新建表格）:')
            # 输入为exit时
            if sheet_name == 'exit':
                # 跳出循环
                break
            # 输入不为exit时候
            else:
                # 添加表格
                worksheet = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
        except:
            # 操作错误，重新输入
            print("操作错误，请重新操作！")
    while True:
        # 定义可进行操作的数组
        add_sheet_operations = ['向当前表格写入内容', "保存表格并进行其他操作", '不了，我要退出操作']
        # 展示可进行操作
        for i in range(len(add_sheet_operations)):
            print(str(i + 1) + "." + add_sheet_operations[i])
        try:
            # 输入选择操作的编号
            add_sheet_operations_choose = int(input("输入序号选择需要执行的操作："))
            # 输入操作编号为1时
            if add_sheet_operations_choose == 1:
                # 进行输入操作
                holden = write_sheet(filename, workbook, worksheet)
                break
            # 输入操作编号为2时
            elif add_sheet_operations_choose == 2:
                # 保存excel
                save_excel(filename, workbook)
                # 定义进行其他操作的变量
                holden = 0
                break
            # 退出程序
            elif add_sheet_operations_choose == 3:
                # 保存excel
                save_excel(filename, workbook)
                # 定义退出程序的变量
                holden = 1
                break
        except:
            print("序号输入错误，请重试！")
    # 返回下一步操作的变量
    return holden


def creat_excel():
    # 输入新建文件的文件名
    filename = input("请输入新建文件的文件名(不需要加文件后缀):") + '.xls'
    # 打开新建的工作簿
    workbook = xlwt.Workbook(encoding="utf-8", style_compression=0)
    while True:
        try:
            # 输入添加表格的名称
            sheet_name = input('请输入表格名称:')
            # 添加表格
            worksheet = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
            break
        except:
            # 文件名输入错误
            print("输入的表格名称不合格法，请重新输入！")
    while True:
        # 定义可进行操作的数组
        create_excel_operations = ['继续添加表格', '向当前表格写入内容', "保存表格并进行其他操作", '不了，我要退出操作']
        # 展示要进行操作
        for i in range(len(create_excel_operations)):
            print(str(i + 1) + "." + create_excel_operations[i])
        try:
            # 输入下一步操作的序号
            create_excel_operations_choose = int(input("输入序号选择需要执行的操作："))
            # 输入的序号为1
            if create_excel_operations_choose == 1:
                # 继续添加表格
                holden = add_sheet(filename, workbook)
                # 跳出循环
                break
            # 输入的序号为2
            elif create_excel_operations_choose == 2:
                # 向当前表格写入内容
                holden = write_sheet(filename, workbook, worksheet)
                # 跳出循环
                break
            # 输入的序号为3
            elif create_excel_operations_choose == 3:
                # 保存表格，进行其他操作
                save_excel(filename, workbook)
                # 定义进行其他操作的变量
                holden = 0
                # 跳出循环
                break
            # 输入的序号为4
            elif create_excel_operations_choose == 4:
                # 退出程序
                save_excel(filename, workbook)
                # 定义退出程序的变量
                holden = 1
                # 跳出循环
                break
        except:
            # 输入序号不在操作的范围内，打印提示，并输入其他内容
            print("序号输入错误，请重试！")
    # 返回下一步操作的变量
    return holden


if __name__ == '__main__':
    print("欢迎进入程序……")
    goon = 0
    while goon == 0:
        operations = ['新建Excel文件', '打开Excel文件', '退出操作']
        for i in range(len(operations)):
            print(str(i + 1) + "." + operations[i])
        try:
            operations_choose = int(input("输入序号选择需要执行的操作："))
            if operations_choose == 1:
                goon = creat_excel()
            elif operations_choose == 2:
                open_excel()
            elif operations_choose == 3:
                print("操作完成，退出程序！")
                break
        except:
            print("输入的序号错误，请重试")
