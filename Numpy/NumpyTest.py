import numpy as np
import random

# 创建一个空的一维数组
vector = np.empty(4)
# 创建一个空的二维数组
matrix = np.empty((2, 4))
for i in range(4):
    # 初始化一维数组
    vector[i] = random.randint(1, 10)
    for j in range(2):
        # 初始化二维数组
        matrix[j][i] = random.randint(10, 20)
# 打印一维数组
print(vector)
# 打印二维数组
print(matrix)
