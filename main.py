import os
import sys
import threading
import win32gui
import cv2
import keyboard
import pyautogui
import numpy as np


def resource_path(relative_path):
    """ 获取资源的绝对路径。用于访问打包后的资源文件。"""
    try:
        # PyInstaller 创建的临时文件夹中获取资源路径
        base_path = sys._MEIPASS
    except Exception:
        # 如果未打包，使用原始路径
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# 使用 resource_path 函数来获取图片路径
bofang_image_path = resource_path("bofang.png")
click_image_path = resource_path("click.png")
shengying_image_path = resource_path("shengying.png")
tuichu_image_path = resource_path("tuichu.png")


def find_element(screen, template):
    gray_screen = cv2.cvtColor(screen, cv2.COLOR_BGR2GRAY)
    res = cv2.matchTemplate(gray_screen, template, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(res)
    if max_val > 0.8:
        return max_loc
    return None


def screenshot():
    print("截图中...")
    screen_size = pyautogui.size()
    img = pyautogui.screenshot(region=(0, 0, screen_size[0], screen_size[1]))
    img = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    return img


def compare_screens(screen1, screen2):
    print("比较截图...")
    gray1 = cv2.cvtColor(screen1, cv2.COLOR_BGR2GRAY)
    gray2 = cv2.cvtColor(screen2, cv2.COLOR_BGR2GRAY)
    diff = cv2.absdiff(gray1, gray2)
    _, thresh = cv2.threshold(diff, 30, 255, cv2.THRESH_BINARY)
    difference = np.sum(thresh)
    print("差异值：", difference)
    return difference > 5000  # 降低阈值


def resize_window(title, class_name, x, y, width, height):
    def get_window_handle(title, class_name):
        def callback(hwnd, handles):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                if class_name == win32gui.GetClassName(hwnd):
                    if title == win32gui.GetWindowText(hwnd):
                        handles.append(hwnd)

        handles = []
        win32gui.EnumWindows(callback, handles)
        return handles[0] if handles else None

    hwnd = get_window_handle(title, class_name)
    if hwnd:
        print(f"找到 '{title}' 窗口, 窗口句柄: {hwnd}")
        win32gui.MoveWindow(hwnd, x, y, width, height, True)
        print(f"窗口大小已调整到 {width}x{height}，位置: ({x}, {y})")
    else:
        print(f"未找到 '{title}' 窗口")


# 使用函数调整微信窗口的大小和位置
resize_window("微信", "Chrome_WidgetWin_0", 500, 80, 924, 876)

while 1:
    quit_button_template = cv2.imread("tuichu.png", cv2.IMREAD_GRAYSCALE)
    # 确保图像模板已正确加载
    if quit_button_template is None:
        print("无法加载退出按钮模板图像，请检查文件路径。")
        exit()
    else:
        print("退出按钮模板图像加载成功")
    # 获取图像的坐标 每隔一秒检测一次 一共检测10次
    back_x = 0
    back_y = 0
    count = 0
    for i in range(10):
        print(f"第 {i + 1}/10 次检测")
        current_screen = screenshot()
        count += 1
        # 检测退出按钮
        quit_button_pos = find_element(current_screen, quit_button_template)
        if quit_button_pos:
            # 存为全局变量
            back_x = quit_button_pos[0]
            back_y = quit_button_pos[1]
            # 跳出循环
            break
        else:
            print("退出按钮未出现，继续检测")
            if count == 10:
                print("退出按钮未出现，退出程序")
                exit()

    print("检测结束")
    # 获得点击坐标 E:\python project\click.png 定为全局变量
    click_button_template = cv2.imread("click.png", cv2.IMREAD_GRAYSCALE)
    # 确保图像模板已正确加载
    if click_button_template is None:
        print("无法加载点击按钮模板图像，请检查文件路径。")
        exit()
    # 获取图像的坐标 每隔一秒检测一次 一共检测10次
    # 定义全局变量
    click_x = 0
    click_y = 0
    count_click = 0
    for i in range(10):
        print(f"第 {i + 1}/10 次检测")
        current_screen = screenshot()
        count_click += 1
        # 检测点击按钮
        click_button_pos = find_element(current_screen, click_button_template)
        if click_button_pos:
            print("点击按钮出现，点击点击按钮")
            pyautogui.click(click_button_pos[0], click_button_pos[1])
            print(f"点击点击按钮，坐标：({click_button_pos[0]}, {click_button_pos[1]})")
            # 存为全局变量
            click_x = click_button_pos[0]
            click_y = click_button_pos[1]
            # 跳出循环
            break
        else:
            print("点击按钮未出现，继续检测")
            if count_click == 10:
                print("点击按钮未出现，退出程序")
                exit()
    print("检测结束")
    print("初始截图...")
    initial_screen = screenshot()
    attempts = 10
    for i in range(attempts):
        print(f"尝试 {i + 1}/{attempts}：点击坐标 ({click_x}, {click_y})")
        pyautogui.click(click_x + 30, click_y - 30)
        pyautogui.sleep(0.5)  # 增加等待时间
        print("新截图...")
        new_screen = screenshot()
        if compare_screens(initial_screen, new_screen):
            print("页面发生跳转")
            break
        else:
            print("页面未发生变化，继续尝试")
            initial_screen = new_screen

    # 识别播放图标 E:\python project\bofang.png
    play_button_template = cv2.imread("bofang.png", cv2.IMREAD_GRAYSCALE)
    # 识别声音图标 E:\python project\shengying.png
    sound_icon_template = cv2.imread("shengying.png", cv2.IMREAD_GRAYSCALE)
    # 确保图像模板已正确加载
    if play_button_template is None:
        print("无法加载播放按钮模板图像，请检查文件路径。")
        exit()
    if sound_icon_template is None:
        print("无法加载声音图标模板图像，请检查文件路径。")
        exit()
    # 获取图像的坐标
    for i in range(50):
        print(f"第 {i + 1}/10 次检测")
        current_screen = screenshot()

        # 检测播放按钮
        play_button_pos = find_element(current_screen, play_button_template)
        if play_button_pos:
            print("播放按钮出现，点击播放按钮")
            pyautogui.sleep(0.2)
            pyautogui.click(play_button_pos[0], play_button_pos[1])
            print(f"点击播放按钮，坐标：({play_button_pos[0]}, {play_button_pos[1]})")
            # 跳出循环
            break
        else:
            print("播放按钮未出现，继续检测")
    print("检测结束")
    # 获取声音的坐标 每隔1秒检测一次 一共检测10次shengying.png
    for i in range(10):
        print(f"第 {i + 1}/10 次检测")
        current_screen = screenshot()

        # 检测声音按钮
        sound_icon_pos = find_element(current_screen, sound_icon_template)
        if sound_icon_pos:
            print("声音按钮出现")
            print("即将点击退出按钮")
            # 跳出循环
            break
        else:
            print("声音按钮未出现，继续检测")
    print("检测结束")
    # 点击退出按钮
    # pyautogui.click(back_x+2, back_y+2)
    # 循环点击退出按钮 检测页面是否发生变化
    attempts = 10
    for i in range(attempts):
        print(f"尝试 {i + 1}/{attempts}：点击坐标 ({back_x}, {back_y})")
        pyautogui.click(back_x + 20, back_y + 8)
        print("新截图...")
        new_screen = screenshot()
        if compare_screens(initial_screen, new_screen):
            print("页面发生跳转")
            break
        else:
            print("页面未发生变化，继续尝试")
            initial_screen = new_screen
