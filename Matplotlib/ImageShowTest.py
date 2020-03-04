import matplotlib.pyplot as plt

# 打开指定图片
image = plt.imread('test.jpg')  # 返回值:灰度图像(M,N),RGB图像(M,N,3),RGBA图像(M,N,4)
# 指定图片中的红波段
image_red = image[:, :, 0]  # 绿波段是 1, 蓝波段是 2.
# 不做处理,显示整张图片
plt.imshow(image)
# 窗口中显示图片
plt.show()
# 不做处理,显示整张红波段
plt.imshow(image_red)
# 窗口中显示图片
plt.show()
# 将红波段转换成灰度图展示,即ColorMap的样式为Greys_r
plt.imshow(image_red, cmap='Greys_r')
# 窗口中展示灰度图
plt.show()
# 将红波段转换成热力图显示,即ColorMap的样式为hot_r
plt.imshow(image_red, cmap='hot_r')
# 窗口显示热力图
plt.show()
# 上下反转图像
plt.imshow(image, origin='lower')
# 窗口显示上下翻转的图像
plt.show()
# 对图像进行高斯模糊
plt.imshow(image, interpolation='gaussian')
# 窗口显示高斯模糊后的图像
plt.show()
