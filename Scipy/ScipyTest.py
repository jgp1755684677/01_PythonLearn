import scipy.stats as st
import random

# 定义空的一个数组
array = []
# 自定义输入数组
Str = input("请输入数字,以空格间隔,回车结束\n")
# 将输入的数字存储到数组中
for i in range(len(Str.split())):
    array.append(int(Str.split()[i]))
# 打印一下数组
print(array)
# 计算并打印偏度
print(st.skew(array))
# 计算并打印峰度
print(st.kurtosis(array))
