import numpy as np
import cv2

words_lib = 'words_lib'


def image_cat(image_cv):
    height, width = image_cv.shape[:2]
    # print('x, y = %d,%d' % (width, height))

    x0 = []
    temp = 0
    sum_y = 255 * height
    image_sun_y = np.sum(image_cv, axis=0)  # 行求和
    for x in range(width):
        if temp == 0 and image_sun_y[x] != sum_y:
            x0.append(x)
            temp = 1
        if temp == 1 and image_sun_y[x] == sum_y:
            x0.append(x)
            temp = 0

    # print(x0, len(x0))

    image_cell_v = np.hsplit(image_cv, x0)  # 按数组中相邻两个点分割成列

    y0 = []
    temp = 0
    for i in range(1, len(x0), 2):
        image_sun_x = np.sum(image_cell_v[i], axis=1)  # 列求和
        sum_x = (x0[i] - x0[i - 1]) * 255
        for y in range(height):
            if temp == 0 and image_sun_x[y] != sum_x:
                y0.append(y)
                temp = 1
            if temp == 1 and image_sun_x[y] == sum_x:
                y0.append(y)
                temp = 0

    # print(y0, len(y0))

    image_cell = []

    for x in range(0, len(x0), 2):
        # 裁剪坐标为[y0:y1,x0:x1]
        image_cell.append(image_cv[y0[x] - 1:y0[x + 1] + 1, x0[x] - 1:x0[x + 1] + 1])

    return image_cell


def matching(target, template):  # 读取处理对象图片, 模板图片
    result = cv2.matchTemplate(target, template, cv2.TM_SQDIFF_NORMED)  # 执行模板匹配，采用的匹配方式cv2.TM_SQDIFF_NORMED
    cv2.normalize(result, result, 0, 1, cv2.NORM_MINMAX, -1)  # 归一化处理
    # 寻找矩阵（一维数组当做向量，用Mat定义）中的最大值和最小值的匹配结果及其位置
    _, _, min_loc, _ = cv2.minMaxLoc(result)

    if min_loc[0] < 30:
        result = '0'
    elif 49 < min_loc[0] < 75:
        result = '1'
    elif 91 < min_loc[0] < 121:
        result = '2'
    elif 140 < min_loc[0] < 169:
        result = '3'
    elif 188 < min_loc[0] < 217:
        result = '4'
    elif 236 < min_loc[0] < 265:
        result = '5'
    elif 284 < min_loc[0] < 314:
        result = '6'
    elif 333 < min_loc[0] < 362:
        result = '7'
    elif 381 < min_loc[0] < 411:
        result = '8'
    elif 430 < min_loc[0] < 460:
        result = '9'
    elif 479 < min_loc[0] < 496:
        result = '.'
    elif 503 < min_loc[0] < 540:
        result = 'm'
    elif 567 < min_loc[0] < 597:
        result = 'V'
    else:
        result = -1

    return result


def osc_matching(image):
    image_cell = image_cat(image)
    result = ''
    template = cv2.imdecode(np.fromfile(words_lib, dtype=np.uint8), cv2.IMREAD_GRAYSCALE)  # 解决中文路径问题
    for cell in image_cell:
        result += matching(cell, template)

    return result
