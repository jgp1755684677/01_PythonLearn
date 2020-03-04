import numpy as np
from scipy.optimize import leastsq
import matplotlib.pyplot as plt
import random


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
    s = '最小二乘法拟合参数估计'
    fit_result = leastsq(error, factor, args=(xArray, yArray, s))
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
    # 在1~100范围内随机生成100个随机点
    draw_x_array = np.linspace(0, 100, 100)
    # 获得一次函数两个参数
    k, b = fit_result[0]
    # 打印拟合函数
    print('y =', k, 'x + ', b)
    # 生成拟合函数
    draw_y_array = k*draw_x_array + b
    # 绘制拟合函数曲线
    plt.plot(draw_x_array, draw_y_array, c='black', linestyle='--', label='Fitting line', linewidth=3)
    # 绘制图例
    plt.legend()
    # 展示图表
    plt.show()


if __name__ == "__main__":
    # 定义两个空数组,作为x,y坐标
    xArray = np.empty(20)
    yArray = np.empty(20)
    # 向两个数组随机赋值1~100的数
    for i in range(20):
        xArray[i] = random.randint(1, 100)
        yArray[i] = random.randint(1, 100)
    # 给两个数组从小到大排序
    xArray.sort()
    yArray.sort()
    # 定义开始拟合函数两个常数
    factor = [1, 1]
    # 调用最小二乘法拟合函数
    fit_result = least_square(factor, xArray, yArray)
    # 调用绘图
    draw(fit_result, xArray, yArray)
