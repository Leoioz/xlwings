import pyautogui  #pip install pyautogui
import os
import sys
import math
import time
import cv2 #pip install opencv-python
import numpy
import PIL #pip install pillow
import pymysql

# #pyautogui.rightClick(x=200,y=100,interval=3,button='right',duration=1.0,logScreenshot=1)
# #写在内容前的变量声明
# x,y=100,100#该参数是一个示例坐标参数
# xoffset,yoffset=200,200#该参数是一个示例偏移坐标参数

# #执行延迟中断
# pyautogui.PAUSE=3#pyautogui包暂停中断，只能在特殊动作后暂停
# time.sleep(3)#系统级delay

# #获取屏幕系数
# width,height=pyautogui.size()#pyautogui包size函数传递返回wight，height
# print (width,height)

# #判断坐标系数
# pyautogui.onScreen(x,y)#pyautogui包该函数入参可搭配坐标识别

# #pyautogui鼠标类操作指令，显示、单击、双击、三击、按住拖动，分解动作等等
# pyautogui.displayMousePosition()#中断显示当前鼠标坐标
# pyautogui.moveTo(x,y)
# pyautogui.moveRel(xoffset,yoffset)#pyautogui包移动偏移函数
# pyautogui.dragTo(x,y)#pyautogui包按住拖动函数
# pyautogui.dragRel(xoffset,yoffset)#pyautogui包按住拖动偏移函数
# pyautogui.click(x,y,button='right')#pyautogui包鼠标单击函数，button提供按键修改
# pyautogui.doubleClick(x,y,button='right')#pyautogui包鼠标双击函数，button提供按键修改
# pyautogui.tripleClick(x,y,button='right')#pyautogui包鼠标三击函数，button提供按键修改
# pyautogui.mouseDown(button='right')#pyautogui包鼠标按下不释放函数，click分解，button提供按键修改
# pyautogui.mouseUp(button='right',x=100,y=100)#pyautogui包鼠标拖到XY再释放函数，click分解，button提供按键修改,xy参数不提供传递变量！

# #缓动/渐变函数可以改变光标移动过程的速度和方向。通常鼠标是匀速直线运动，简称花里胡哨PyAutoGUI有30种缓动/pyautogui.ease*?
# pyautogui.moveTo(100, 100, 2, pyautogui.easeInQuad)# 开始很慢，不断加速
# pyautogui.moveTo(100, 100, 2, pyautogui.easeOutQuad)# 开始很快，不断减速
# pyautogui.moveTo(100, 100, 2, pyautogui.easeInOutQuad)# 开始和结束都快，中间比较慢
# pyautogui.moveTo(100, 100, 2, pyautogui.easeInBounce)# 一步一徘徊前进
# pyautogui.moveTo(100, 100, 2, pyautogui.easeInElastic)# 徘徊幅度更大，甚至超过起点和终点

# #消息类提示
# #pyautogui.alert('这个消息弹窗是文字+OK按钮')
# #pyautogui.confirm('这个消息弹窗是文字+OK+Cancel按钮')
# #pyautogui.prompt('这个消息弹窗是让用户输入字符串，单击OK')

# #截屏图像类处理，需要Pillow库依赖;screenshot('截图名称',region=划分截图范围)
# im1 = pyautogui.screenshot()#直接主屏幕全屏
# im2 = pyautogui.screenshot('my_screenshot.png')#默认截图在python工程文件目录，取名进行传递参数处理
# im3 = pyautogui.screenshot(region=(100, 0, 300 ,400))#region划分范围截图
# im4 = pyautogui.screenshot('my_screenshot.png',region=(100, 0, 300 ,400))#region划分范围截图

# # 对图像截屏的内容进行输出保存，可选项
# img_path = r'C:\Users\HNCJ-liaobingzhi\Desktop\2.png'
# im3.save(img_path)#save是python通用函数

#对图像截屏的内容进行操作
# postion=pyautogui.locateOnScreen(im3)#locateonscreen函数可以获得参数在屏幕上的左边，参数可以是im1或者直接指定的文件路径
# postion=pyautogui.locateOnScreen(image=r'C:\Users\HNCJ-liaobingzhi\Desktop\2.png')
# postion=pyautogui.locateOnScreen(image=r'C:\Users\HNCJ-liaobingzhi\Desktop\2.png',grayscale=True)
# postion=pyautogui.center(postion)#center函数将鼠标移动至识别到的图片中间
# left,top,width,height=postion#以上输出坐标为四元组，可传递参数
# postion=pyautogui.locateAllOnScreen(image=r'C:\Users\HNCJ-liaobingzhi\Desktop\2.png')
# print(left,top,width,height,postion)
# pyautogui.moveTo(postion)
# pyautogui.doubleClick(button='left')
# postion=pyautogui.locateAll()
# postion=pyautogui.locateAllOnScreen()
# postion=pyautogui.locateCenterOnScreen()
# postion=pyautogui.locateOnWindow()

# im1 = pyautogui.screenshot()#直接主屏幕全屏
# im1.getpixel((100,200))
# print(im1.getpixel((100,200)))

# secs_between_keys = 0.1
# pyautogui.typewrite('Hello world!\n', interval=secs_between_keys)

UAC_control=r'xxx.png'
Input_code='xxxx\n'

def function_input_code(code):
    secs_between_keys = 0.1
    pyautogui.typewrite(code, interval=secs_between_keys)
    pass

def function_find_and_click(image):
    # pyautogui.hotkey('alt','tab')
    x,y,z,w=pyautogui.locateOnScreen(image)
    print(x,y,z,w)
    # pyautogui.moveTo(x+z/2,y+w/2)
    pyautogui.click(x+z/2,y+w/2,button='left')
    time.sleep(1)
    pass

function_find_and_click()
function_input_code()