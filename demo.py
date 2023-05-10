import win32gui
import numpy as np
from typing import Tuple
import cv2
import win32con
from PIL import ImageGrab
import os
from PIL import Image
import pyautogui
import time
from win32com.client import Dispatch
hwnd_l = {}
speaker = Dispatch('SAPI.SpVoice')


def get_hwnd(hwnd, mouse):
    if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
        hwnd_l.update({hwnd: win32gui.GetWindowText(hwnd)})


def search_and_click_image(screen, image2, sec):
    """
    查找大图中小图标的位置并且点击
    :param screen: 背景
    :param image2: 背景中的图片
    :param sec:    停留时间
    :return:
    """
    #  width, height = image2.shape
    top_left = search_image(screen, image2)
    #  bottom_right = (top_left[0] + width, top_left[1] + height)
    pyautogui.click(top_left[0] + 20, top_left[1] + 20, button='left')
    time.sleep(sec)


def search_image(image1, image2):
    """
    查找图片
    """
    res = cv2.matchTemplate(image1, image2, cv2.TM_SQDIFF)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    return min_loc


def search_image_and_image(image1, image2, sec=0.1):
    """
    查找图中图
    :param image1:
    :param image2:
    :param sec:
    :return:
    """
    top_left = search_image(get_screen(), image1)
    top1_left1 = search_image(image1, image2)
    position = (top_left[0]+top1_left1[0]+image1.shape[0]+20, top_left[1]+top1_left1[1]+20)
    pyautogui.click(*position, clicks=2, button='left')
    time.sleep(sec)


def show_video(image):
    cv2.imshow('video', image)
    cv2.waitKey(0)
    cv2.destroyAllWindows()


def image_transform_gray(file):
    image = np.asarray(Image.open(file))
    image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    return image


def get_screen():
    screen_image = np.asarray(ImageGrab.grab(windows_rect), np.uint8)
    screen_gray = cv2.cvtColor(screen_image, cv2.COLOR_BGR2GRAY)
    return screen_gray


def scroll_and_moveto_mouse(buffer=500, mouse_x=1000, mouse_y=600):
    pyautogui.moveTo(mouse_x, mouse_y)
    pyautogui.scroll(buffer)
    time.sleep(1)


def click_image(image2, sec):
    search_and_click_image(get_screen(), image2, sec)


def dragto_init(duration=2):
    pyautogui.moveTo(1500, 520)
    with pyautogui.hold('w'):
        pyautogui.dragTo(1100, 520, duration=duration, button='left')


def stack_map1(epochs):
    for _ in np.arange(epochs):
        click_image(image_params['训练兵4'], 0.1)
        dragto_init()
        click_image(image_params['打资源1'], 0.1)  # 进入模式选择
        click_image(image_params['单人模式'], 0.1)
        scroll_and_moveto_mouse()
        click_image(image_params['恶有恶报'], 0.1)
        click_image(image_params['进攻'], 2)
        click_image(image_params['蛮王'], 0.1)
        pyautogui.click(150, 500)
        time.sleep(1)
        click_image(image_params['蛮王技能'], 13)
        click_image(image_params['回营'], 4)


def clear_trees(epochs):
    click_image(image_params['去往夜世界'], 1)
    dragto_init()
    tree_list = [image_params['树'], image_params['大树'], image_params['大树林']]
    for i in range(epochs):
        for j in range(4):
            click_image(tree_list[j], 1)
            click_image(image_params['移除'], 11)


def train_army(name: str = '训练黄毛弓箭手'):
    """
    训练兵种：（黄毛，弓箭手）
    :return:
    """
    click_image(image_params['训练兵4'], 0.1)
    dragto_init()
    click_image(image_params['训练兵1'], 0.1)
    click_image(image_params['训练兵2'], 0.1)
    search_image_and_image(image_params[name], image_params['小训练兵'])
    click_image(image_params['训练兵4'], 0.1)


def send_troops(start_pos: Tuple, end_pos: Tuple, num: int = 50):
    pos_l = []
    error = (end_pos[0] - start_pos[0], end_pos[1] - start_pos[1])
    for i in np.arange(num):
        current_pos = (round(start_pos[0] + i / num * error[0]), round(start_pos[1] + i / num * error[1]))
        pos_l.append(current_pos)
    return pos_l


def brush():
    """
    黄毛和弓箭手打资源
    :return:
    """
    king_pos = (1372, 227)
    dragto_init()
    click_image(image_params['打资源1'], 0.1)
    click_image(image_params['打资源2'], 4)
    click_image(image_params['蛮王'], 0.1)
    pyautogui.click(*king_pos, button='left')
    click_image(image_params['女王'], 0.1)
    pyautogui.click(*king_pos, button='left')
    click_image(image_params['闰土'], 0.1)
    pyautogui.click(*king_pos, button='left')
    click_image(image_params['永王'], 0.1)
    pyautogui.click(*king_pos, button='left')
    for t in range(2):
        scroll_and_moveto_mouse(-500)
        click_image(image_params['黄毛'], 0.1)
        for pos_x, pos_y in right_down_l:
            pyautogui.click(pos_x, pos_y, button='left')
        click_image(image_params['弓箭手'], 0.1)
        for pos_x, pos_y in down_left_l:
            pyautogui.click(pos_x, pos_y, button='left')
        scroll_and_moveto_mouse(500)
        click_image(image_params['黄毛'], 0.1)
        for pos_x, pos_y in left_top_l:
            pyautogui.click(pos_x, pos_y, button='left')
        click_image(image_params['弓箭手'], 0.1)
        for pos_x, pos_y in top_right_l:
            pyautogui.click(pos_x, pos_y, button='left')
    time.sleep(3*60-45)
    click_image(image_params['回营'], 4)


def handle_and_get_dict():
    head_path = 'icon'
    pathdir = os.listdir(head_path)
    image_dict = {}
    for end_path in pathdir:
        path = os.path.join(head_path, end_path)
        image = image_transform_gray(path)
        image_dict.update({end_path.split('.')[0]: image})
    return image_dict


def train_brush(epochs):
    for e in range(1, epochs+1):
        speaker.Speak('现在开始第%d回合' % e)
        start = time.time()
        train_army()
        for c in range(1, 10, 1):
            dragto_init()
            time.sleep(60)
        brush()
        speaker.Speak('第{}回合结束,用时{:d}秒'.format(e, round(time.time() - start)))


def electric_dragon():
    king_pos = (1311, 593)
    dragto_init()
    click_image(image_params['打资源1'], 0.1)
    click_image(image_params['打资源2'], 4)
    scroll_and_moveto_mouse(-500)
    click_image(image_params['电龙'], 0.1)
    for pos_x, pos_y in blue_dragon_l:
        pyautogui.click(pos_x, pos_y, button='left')
    click_image(image_params['蛮王'], 0.1)
    pyautogui.click(*king_pos, button='left')
    click_image(image_params['闰土'], 0.1)
    pyautogui.click(*king_pos, button='left')
    click_image(image_params['永王'], 0.1)
    pyautogui.click(*king_pos, button='left')
    click_image(image_params['女王'], 0.1)
    pyautogui.click(*king_pos, button='left')
    click_image(image_params['冰冻'], 0.1)
    pyautogui.click(1109, 378, button='left')
    click_image(image_params['狂暴'], 0.1)
    for pos_x, pos_y in violent_l:
        pyautogui.click(pos_x, pos_y, button='left')
    time.sleep(5)
    click_image(image_params['永王技能'], 0.1)
    time.sleep(160)
    click_image(image_params['回营'], 4)


def train_electric_dragon(epochs):
    for epoch in range(1, epochs+1):
        speaker.Speak('现在开始第%d回合' % epoch)
        start = time.time()
        train_army('电龙一字划')
        speaker.Speak('正在训练雷龙中,请耐心等待55分钟')
        for t in range(1, 56):
            dragto_init(duration=1)
            time.sleep(60)
            speaker.Speak('已经训练:%d分钟了' % t)
        speaker.Speak('准备完毕,准备开干')
        electric_dragon()
        speaker.Speak('第{}回合结束,用时{:d}秒'.format(epoch, round(time.time() - start)))


violent_start_pos = (950, 605)
violent_end_pos = (1395, 277)
blue_dragon_start_pos = (1027, 769)
blue_dragon_end_pos = (1560, 374)
down_center_pos = (937, 833)
down_right_pos = (1657, 292)
down_left_pos = (213, 289)
top_center_pos = (958, 108)
top_left_pos = (222, 660)
top_right_pos = (1660, 649)
win32gui.EnumWindows(get_hwnd, 0)
hwnd_ = win32gui.FindWindow(None, '雷电模拟器')
win32gui.SetForegroundWindow(hwnd_)
win32gui.ShowWindow(hwnd_, win32con.SW_MAXIMIZE)
image_params = handle_and_get_dict()
windows_rect = win32gui.GetWindowRect(hwnd_)
EPOCHS = 30
right_down_l = send_troops(down_right_pos, down_center_pos)
down_left_l = send_troops(down_center_pos, down_left_pos)
left_top_l = send_troops(top_left_pos, top_center_pos)
top_right_l = send_troops(top_center_pos, top_right_pos)
blue_dragon_l = send_troops(blue_dragon_start_pos, blue_dragon_end_pos, num=10)
violent_l = send_troops(violent_start_pos, violent_end_pos, num=5)
train_electric_dragon(10)
