```python
import xlrd
import xlwt
from xlutils.copy import copy
import pandas as pd


def judge_input(con):
    while True:
        # 判断输入的参数是不是要求输入的内容
        if con == 'n' or con == 'o' or con == 'e':
            break
        else:
            con = input("输入错误,请重新输入:")
    return con


def creat_excel():
    # 输入文件的文件名
    filename = input("请输入新建文件的文件名(不需要加文件后缀):") + '.xls'
    # 打开工作簿
    excel = xlwt.Workbook(encoding="utf-8", style_compression=0)
    # 输入添加表格的名称
    sheet_name = input('请输入表格名称:')
    # 添加表格
    sheet = excel.add_sheet(sheet_name, cell_overwrite_ok=True)
    # 判断是否在工作簿中继续添加表格
    while True:
        # 选择是否添加表格
        con = input('是否继续添加表格(y/n):')
        # 如果是,添加表格
        if con == 'y' or con == 'Y':
            # 输入添加表格的名称
            sheet_name = input('请输入表格名称:')
            # 添加表格
            sheet = excel.add_sheet(sheet_name, cell_overwrite_ok=True)
        # 不添加表格,跳出循环
        elif con == 'n' or con == 'N':
            print("新建文件结束,请选择其他操作!\n")
            break
        # 输入的字符串既不是新建也不是不新建
        else:
            print("输入错误,请重新输入!\n")
            continue
    # 保存xls文件
    excel.save(filename)
    # 新建文件完成
    print("新建文件完成,文件名:" + filename + "\n")


def read_excel(filename):
    # 打开文件名为filename的xls文件
    workbook = xlrd.open_workbook(filename)
    # 获取excel文件中所有sheet的名称
    sheets = workbook.sheet_names()
    # 打印所有的sheet的名称
    for i in range(len(sheets)):
        print("第" + str(i + 1) + '个sheet的名称为:' + sheets[i] + '\n')
    # 选择要打开的sheet
    sheet = int(input("输入序号,选择你要打开的表格:")) - 1
    # 打开选择的sheet
    worksheet = workbook.sheet_by_index(sheet)
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
    # 返回打开的worksheet和带标签的数组
    return filename, sheet, worksheet


def write_excel(filename, sheet):
    workbook = xlrd.open_workbook(filename)
    new_workbook = copy(workbook)
    new_worksheet = new_workbook.get_sheet(sheet)
    i = 1
    rows = []
    while True:
        strings = input("请输入第" + str(i) + "行数据(不同单元格之间用','隔开,结束请输入n):\n")
        if strings == 'n' or strings == 'N':
            break
        else:
            if strings =='':
                i += 1
                continue
            else:
                for j in range(len(strings.split(','))):
                    new_worksheet.write(i - 1, j, strings.split(',')[j])
                i += 1
    new_workbook.save(filename)
    print("写入数据完成！")


def change_excel(filename, sheet):
    workbook = xlrd.open_workbook(filename)
    new_workbook = copy(workbook)
    new_worksheet = new_workbook.get_sheet(sheet)
    while True:
        cell = input("请输入修改单元格的行列号及修改内容（例如，第0行第0列修改成test输入--0,0,text）:")
        if cell == 'n' or 'N':
            break
        else:
            temp = cell.split(',')
            if len(temp) == 3:
                new_worksheet.write(int(temp[0]), int(temp[1]), temp[2])
                print("修改成功")
            else:
                print("输入错误，请重新输入！\n")


def draw_scatter(filename, sheet):
    pass


def draw_plot(filename, sheet):
    pass


def excel_sum(filename, sheet):
    pass


def excel_mean(filename, sheet):
    pass


def excel_least_square(filename, sheet):
    pass


if __name__ == '__main__':
    # 选择进行什么操作
    condition = input('选择操作[n(新建),o(打开),e(退出)]:')
    # 判断输入是否正确
    condition = judge_input(condition)
    while True:
        # 输入为n时,新建一个excel文件
        if condition == 'n':
            # 新建文件
            creat_excel()
            # 输入下一步的操作
            condition = input("选择操作,n(继续新建),o(打开文件),e(退出):")
            # 判断输入是否正确
            condition = judge_input(condition)
            # 进行下一次循环
            continue
        elif condition == 'o':
            # 输入打开文件的名称
            filename = input("输入文件名打开文件:\n")
            while True:
                # 判断文件名是不是以xls结尾
                if filename.endswith(".xls"):
                    print("正在打开文件:" + filename + '\n')
                    filename, sheet, worksheet = read_excel(filename)
                    o_condition = input("是否进行下一步操作,y(是)或n(否):")
                    while True:
                        if o_condition == 'y' or o_condition == 'Y' or o_condition == 'n' or o_condition == 'N':
                            break
                        else:
                            o_condition = input("输入错误,请重新输入:")
                            continue
                    if o_condition == 'y' or o_condition == 'Y':
                        operations = ['写入数据', '修改', '绘制散点图', '绘制直线图', '求和', '求平均值', '最小二乘法拟合']
                        for i in range(len(operations)):
                            print(str(i + 1) + '.' + operations[i] + '\n')
                        number = int(input("请输入的操作的序号(例如:1):"))
                        while True:
                            if 1 <= number <= 7:
                                break
                            else:
                                number = int(input("输入错误,请重新输入:"))
                        if number == 1:
                            # 向打开的表格中写入数据
                            write_excel(filename, sheet)
                            break
                        elif number == 2:
                            write_excel(filename, sheet)
                            break
                        elif number == 3:
                            draw_scatter(filename, sheet)
                            print()
                            break
                        elif number == 4:
                            draw_plot(filename, sheet)
                            print()
                            break
                        elif number == 5:
                            excel_sum(filename, sheet)
                            break
                        elif number == 6:
                            print()
                            excel_mean(filename, sheet)
                            break
                        elif number == 7:
                            excel_least_square(filename, sheet)
                            break
                    elif o_condition == 'n' or o_condition == 'N':
                        condition = 'e'
                        break
                else:
                    filename = input("文件名称输入错误,检查文件名是否错误或者以.xls结尾后,重新输入:")
                    continue
            continue
        elif condition == 'e':
            print("退出程序!\n")
            break
    print("完成操作!\n")

```