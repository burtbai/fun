# -*- coding: UTF-8 -*-

from PIL import Image
import numpy as np


def generate(in_path, depth, out_path):
    image = Image.open(in_path)
    a = np.asarray(image.convert('L')).astype('float')
    # 根据灰度变化来楧拟人类视觉的明暗程度
    # depth = 10  # 预设虚拟深度值为10 范用为0 - 100
    grad = np.gradient(a)  # 提取梯度值
    grad_x, grad_y = grad  # 提取× y方向梯度值 解构赋给grad_x, grad_y

    # 利用像素之间的梯度值和虚拟深度值对图像进行重构
    grad_x = grad_x * depth / 100.
    grad_y = grad_y * depth / 100.  # 根据深度调竖 × y 方向梯度值

    A = np.sqrt(grad_x ** 2 + grad_y ** 2 + 1.)
    uni_x = grad_x / A
    uni_y = grad_y / A
    uni_z = 1. / A

    vec_el = np.pi / 2.2  # 光源的俯视角度 弧度值
    vec_az = np.pi / 4.  # 光源的方位角度 弧度值
    dx = np.cos(vec_el) * np.cos(vec_az)  # 光源对x轴影响
    dy = np.cos(vec_el) * np.sin(vec_az)  # 光源对y轴影响
    dz = np.sin(vec_el)  # 光源对z轴影响

    b = 255 * (dx * uni_x + dy * uni_y + dz * uni_z)  # 光源归一化
    b = b.clip(0, 255)  # 为了避免数据越界，将生成辉度值裁剪至0 - 255 区间
    im = Image.fromarray(b.astype("uint8"))  # 图像重构
    im.save(out_path)  # 保存图片


if __name__ == '__main__':
    path = r'/Users/burt/Downloads/test_1.jpeg'
    # path = r'/Users/burt/Downloads/test_2.webp'
    for depth in range(1, 11):
        out_path = r'./output/out_%s.jpeg' % depth
        generate(path, depth, out_path)
