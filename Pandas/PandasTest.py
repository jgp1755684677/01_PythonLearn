import numpy as np
import pandas as pd
import random

# 定义一个空的多维数据
grades = np.empty((4, 4))
for i in range(4):
    for j in range(4):
        grades[i][j] = random.randint(40, 100)

# 给数组的行列加上标记
df = pd.DataFrame(grades, columns=['GIS', 'AI', 'GPS', 'RS'], index=['A', 'B', 'C', 'D'])
print('------------打印一下多维数组----------')
print(df)

print('--------------计算 df 第二行的平均值--------------')
secondRow = df.iloc[1, :].mean()  # 计算第二行的平均值
print(secondRow)

print('--------------计算 df 计算所有行的平均值--------------')
allRows = df.mean(axis=1)  # 计算所有行的平均值
print(allRows)

print('--------------计算 df 第三列的平均值--------------')
ThirdCols = df.iloc[:, 2].mean()  # 计算第三列的平均值
print(ThirdCols)

print('--------------计算 df 计算所有列的平均值--------------')
AllCols = df.mean(axis=0)  # 计算所有列的平均值
print(AllCols)
