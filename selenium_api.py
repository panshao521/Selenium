#--coding:utf-8--
# @Time    : 2021/1/6/006 22:53
# @Author  : panyuangao
# @File    : selenium_api.py
# @PROJECT : Selenium
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait#引入显式等待条件的包
from selenium.webdriver.support import expected_conditions as EC
import win32clipboard
import win32api
import win32con
import traceback
import time,os

# ==================================== 【浏览器操作】 =============================================

def browser(browser_name): # 通过浏览器名称，打开对应的浏览器（chrome、ie、firefox）
    global driver
    if "chrome" in browser_name:
        driver = webdriver.Chrome(executable_path = "e:\\chromedriver")
    elif "ie" in browser_name:
        driver = webdriver.Ie(executable_path="e:\\IEDriverServer")
    elif"firefox" in browser_name:
        driver = webdriver.Firefox(executable_path="e:\\geckodriver")
    return driver

def visit(url): # 访问某个url
    global driver
    driver.get(url)

def back(): # 网页后退
    global driver
    driver.back()

def forward(): # 网页前进
    global driver
    driver.forward()

def refresh(): # 网页刷新
    global driver
    driver.refresh()

def close(): # 关闭当前标签页
    global driver
    driver.close()

def quit(): # 关闭当前浏览器
    global driver
    driver.quit()

def maximize(): #浏览器最大化
    global driver
    driver.maximize_window()

def get_position(): # 获取当前窗口的坐标
    global driver
    window_position = driver.get_window_position()
    print("当前浏览器窗口的坐标：%s" %window_position)
    return window_position

def set_position(x,y): # 将当前窗口移动到坐标（x,y）的位置
    global driver
    driver.set_window_position(x=x,y=y)
    print("已将浏览器窗口移动至坐标（'x': %s, 'y': %s）的位置" %(x,y))

def get_size(): # 获取当前浏览器窗口大小
    global driver
    window_size = driver.get_window_size(windowHandle='current')
    print("当前浏览器窗口大小：%s" %window_size)
    return window_size

def set_size(width,height): # 设置当前浏览器窗口大小（width宽,height高）
    global driver
    driver.set_window_size(width = width, height = height, windowHandle='current')
    print("当前浏览器窗口大小设置：{'width': %s, 'height': %s}" %(width,height))

def get_title(): # 获取浏览器标题栏文字内容
    global driver
    title = driver.title
    print("当前浏览器的标题：%s" %title)

def page_source(): # 获取网页源码内容
    global driver
    page_source = driver.page_source
    return page_source

def get_url(): # 获取当前网页url
    global driver
    current_url = driver.current_url
    print("当前网页的url：%s" %current_url)
    return current_url

def get_window_handle(): #获取当前浏览器窗口句柄
    global driver
    window_handle = driver.current_window_handle
    print("当前浏览器窗口句柄：%s" % window_handle)
    return window_handle

def get_window_handles(): #获取所有窗口句柄
    global driver
    window_handles = driver.window_handles
    print("所有浏览器窗口句柄：%s" % window_handles)
    return window_handles

def switch_to(window_handle): # 通过窗口句柄切换窗口
    driver.switch_to.window(window_handle)
    print("切换浏览器窗口成功，窗口句柄：%s" %window_handle)




# ==================================== 【页面元素定位】 =============================================

def id_find(element_id): # 通过id定位页面元素
    global driver
    try:
        element = driver.find_element_by_id(str(element_id))
        return element
    except NoSuchElementException as e:# 捕获NoSuchElementException异常
        traceback.print_exc()
        raise e

def name_find(element_name): # 通过name定位页面元素
    global driver
    try:
        element = driver.find_element_by_name(element_name)
        return element
    except NoSuchElementException as e:# 捕获NoSuchElementException异常
        traceback.print_exc()
        raise e

def xpath_find(xpath_exp): # 通过xpath表达式定位页面元素
    global driver
    try:
        element = driver.find_element_by_xpath(xpath_exp)
        return element
    except NoSuchElementException as e:# 捕获NoSuchElementException异常
        traceback.print_exc()
        raise e

def link_text_find(link_text): # 通过链接文本定位页面元素
    global driver
    try:
        element = driver.find_element_by_link_text(link_text)
        return element
    except NoSuchElementException as e:# 捕获NoSuchElementException异常
        traceback.print_exc()
        raise e

def element_find(find_type,find_exp): # 通过多种方式定位
    '''
    :param find_type: 定位方式
    :param find_exp: 定位表达式
    :return:返回定位到的元素
    示例
    driver.find_element("id","kw")       #通过id定位
    driver.find_element("name","wd")     #通过名称定位
    driver.find_element("xpath","//input[@id='kw']")    #通过xpath定位
    driver.find_element("link text","使用百度前必读")   #通过链接的全部文字去定位
    driver.find_element("partial link text","前必读")   #通过链接的部分文字去定位
    '''
    global driver
    try:
        element = driver.driver.find_element(find_type,find_exp)
        return element
    except NoSuchElementException as e:# 捕获NoSuchElementException异常
        traceback.print_exc()
        raise e

def css_find(css_exp): # 通过css定位页面元素(很少用)
    global driver
    try:
        element = driver.find_element_by_css_selector(css_exp)
        return element
    except NoSuchElementException as e:# 捕获NoSuchElementException异常
        traceback.print_exc()
        raise e


# ==================================== 【获取元素属性】 =============================================

def get_element_tag_name(element): # 获取元素标签名称
    return element.tag_name

def get_element_size(element): # 获取元素大小
    return element.size

def get_element_text(element): # 获取元素文字信息
    return element.text

def get_element_attribute(element, attribute_name): # 通过特性名称，获取元素特性值
    '''
    :param element: 元素
    :param attribute_name: 元素特性名称，例如：href,id,text,value; value可获取输入框内的值
    :return: 返回特性的值
    与 get_property()功能类似
    '''
    return element.get_attribute(attribute_name)

def get_element_property(element, property_name):  # 通过属性名称，获取元素属性（标签）
    '''
    :param element: 元素
    :param property_name: 属性名称，例如：href,id,text,value；value可获取输入框内的值
    :return: 返回属性值
    与 get_attribute()功能类似，property 是特殊的 attribute
    '''
    return element.get_property(property_name)

def get_css_property(element,css_property): # 获取页面元素的CSS属性值
    return element.value_of_css_property(css_property)


# ==================================== 【元素判断】 =============================================

def is_displayed(element): # 判断元素是否显示（未被隐藏）
    if element.is_displayed():
        return True
    else:
        print("元素%s被隐藏，未显示" %element)
        return False

def is_enabled(element): # 判断元素是否可用(只读也会返回True)
    if element.is_enabled():
        return True
    else:
        print("元素%s不可用" %element)
        return False

def is_selected(element): # 判断元素是否已被选择
    if element.is_selected():
        return True
    else:
        print("元素%s未被选择" %element)
        return False

def isElementPresent(by, value): #判断元素是否存在
    global driver
    try:
        driver.find_element(by=by, value=value)
    except NoSuchElementException as e:
        # 打印异常信息
        print(e)
        # 发生了NoSuchElementException异常，说明页面中未找到该元素，返回False
        return False
    else:
        # 没有发生异常，表示在页面中找到了该元素，返回True
        return True


# ==================================== 【元素操作】 =============================================

def send_keys(element,content_input): # 输入框中输入内容
    element.send_keys(content_input)

def clear(element): # 清空输入框内容
    element.clear()

def click(element): #点击元素
    try:
        element.click()
    except Exception as e:
        print("元素点击失败")
        traceback.print_exc()
        raise e

def double_click(element): #双击元素
    action_chains = ActionChains(driver)
    action_chains.double_click(element).perform()  # 双击动作


# ==================================== 下拉列表（单选） =============================================
def select_list_option(select_list,option_text): # 通过名称，选择列表选项
    '''
    :param select_list: 选择列表
    :param option_text: 列表选项文本
    :return:
    '''
    all_options = select_list.find_elements_by_tag_name("option")
    for option in all_options:
        if option.text == option_text:
            option.click()
            return
    print("未能选中列表中的%s选项" %option_text)

def select_index(select_list,index): # 下拉框单选（通过序号）
    try:
        select_element = Select(select_list)
        select_element.select_by_index(index)  # 用序号来选中的选项
        option_text = select_element.options[index].text #获取选项文本
        assert option_text in select_element.all_selected_options[0].text
        print("成功选择第'%s'个选项，选项文本：%s" %(index+1,option_text))
    except NoSuchElementException as e:
        print("通过序号，下拉框单选失败")
        traceback.print_exc()
        raise e

def select_text(select_list,option_text): # 下拉框单选（通过选项文本）
    try:
        select_element = Select(select_list)
        select_element.select_by_visible_text(option_text)  # 通过选项文本来选中选项
        assert option_text in select_element.all_selected_options[0].text
        print("成功选择'%s'选项" %option_text)
    except NoSuchElementException as e:  # 捕获NoSuchElementException异常
        print("通过选项文本，下拉框单选失败")
        traceback.print_exc()
        raise e

def select_value(select_list,option_value): # 下拉框单选（通过value）
    try:
        select_element = Select(select_list)
        select_element.select_by_value(option_value)  # 通过value来选中选项
    except NoSuchElementException as e:  # 捕获NoSuchElementException异常
        print("通过value，下拉框单选失败")
        traceback.print_exc()
        raise e

# ==================================== 下拉列表（多选） =============================================

def select_indexs(select_list,indexs): # 下拉框多选（通过序号）
    try:
        select_element = Select(select_list)
        for index in indexs:
            select_element.select_by_index(index)
    except NoSuchElementException as e:  # 捕获NoSuchElementException异常
        print("通过序号，下拉框多选失败")
        traceback.print_exc()
        raise e

def select_texts(select_list,option_texts): # 下拉框多选（通过选项文本）
    try:
        select_element = Select(select_list)
        for option_text in option_texts:
            select_element.select_by_visible_text(option_text)
    except NoSuchElementException as e:  # 捕获NoSuchElementException异常
        print("通过选项文本，下拉框多选失败")
        traceback.print_exc()
        raise e

def select_values(select_list,option_values): # 下拉框多选（通过values）
    try:
        select_element = Select(select_list)
        for option_value in option_values:
            select_element.select_by_value(option_value)
    except NoSuchElementException as e:  # 捕获NoSuchElementException异常
        print("通过values，下拉框多选失败")
        traceback.print_exc()
        raise e

# ==================================== 下拉列表（取消多选） =============================================

def deselect_indexs(select_list,indexs): # 下拉框取消多选（通过序号）
    try:
        select_element = Select(select_list)
        for index in indexs:
            select_element.deselect_by_index(index)
    except NoSuchElementException as e:  # 捕获NoSuchElementException异常
        print("通过序号，下拉框取消多选失败")
        traceback.print_exc()
        raise e

def deselect_texts(select_list,option_texts): # 下拉框取消多选（通过选项文本）
    try:
        select_element = Select(select_list)
        for option_text in option_texts:
            select_element.deselect_by_visible_text(option_text)
    except NoSuchElementException as e:  # 捕获NoSuchElementException异常
        print("通过选项文本，下拉框取消多选失败")
        traceback.print_exc()
        raise e

def deselect_values(select_list,option_values): # 下拉框取消多选（通过values）
    try:
        select_element = Select(select_list)
        for option_value in option_values:
            select_element.deselect_by_value(option_value)
    except NoSuchElementException as e:  # 捕获NoSuchElementException异常
        print("通过values，下拉框取消多选失败")
        traceback.print_exc()
        raise e

# ==================================== 下拉列表（全选、取消全选、反选） =============================================

def select_all(select_list): # 全选所有选项
    try:
        select_element = Select(select_list)
        all_options = select_element.options
        all_selected_options = select_element.all_selected_options
        for option in all_options:
            if option not in all_selected_options:
                option.click()
    except NoSuchElementException as e:  # 捕获NoSuchElementException异常
        print("下拉框全选失败")
        traceback.print_exc()
        raise e

def deselect_all(select_list): # 取消所有选项
    try:
        select_element = Select(select_list)
        select_element.deselect_all()
    except NoSuchElementException as e:  # 捕获NoSuchElementException异常
        print("取消所有选项失败")
        traceback.print_exc()
        raise e

def select_invert(select_list): # 反选已选择的选项
    try:
        select_element = Select(select_list)
        all_selected_options = select_element.all_selected_options
        all_options = select_element.options
        select_element.deselect_all()
        for option in all_options:
            if option not in all_selected_options:
                option.click()
    except NoSuchElementException as e:  # 捕获NoSuchElementException异常
        print("下拉框反选失败")
        traceback.print_exc()
        raise e

# ==================================== 单选框/复选框（操作） =============================================

def select_radio(radio_xpath_exp): # 单选框选择
    global driver
    radio = driver.find_element_by_xpath(radio_xpath_exp) # 通过xpath定位单选框
    if not radio.is_selected():
        radio.click()  #点击单选框

def select_check_box(check_box_xpath_exp): # 单选复选框
    global driver
    check_box = driver.find_element_by_xpath(check_box_xpath_exp) # 通过xpath定位复选框
    if not check_box.is_selected(): #如果复选框未被选择，则点击复选框
        check_box.click()

def select_all_check_box(check_boxs_xpath_exp): # 全选复选框
    global driver
    check_boxs = driver.find_elements_by_xpath(check_boxs_xpath_exp) # 通过xpath定位所有复选框
    for check_box in check_boxs:
        if not check_box.is_selected(): #如果复选框未被选择，则点击复选框
            check_box.click()


def select_invert_check_box(check_boxs_xpath_exp): # 反选复选框
    global driver
    check_boxs = driver.find_elements_by_xpath(check_boxs_xpath_exp) # 通过xpath定位所有复选框
    for check_box in check_boxs:
        check_box.click()



# ==================================== 分隔线 =============================================
def switch_frame(frame_obj): # 页面中切换frame
    global driver
    try:
        driver.switch_to.frame(frame_obj)
    except Exception as e:
        traceback.print_exc()
        print("切换frame失败")
        raise e

# ==================================== 分隔线 =============================================


# ==================================== 【断言、截屏】 =============================================

def assert_word(word): # 断言源码中是否包含关键字
    global driver
    try:
        assert word in driver.page_source
    except AssertionError as e:
        print("断言出现错误！")
        traceback.print_exc()
        raise e

def screen_shot_browser(save_path,img_name): # 浏览器截图
    global driver
    if not os.path.exists(save_path): # 如果目录不存在，新建目录
        os.makedirs(save_path)
    img_path = os.path.join(save_path,img_name+".png")
    driver.get_screenshot_as_file(img_path)  # 截屏，并保存到本地

def screen_shot_windows(save_path,img_name): #系统截图（整个桌面）
    global driver
    from PIL import ImageGrab  # 引用pillow包的方法，需要先进行安装 pip install pillow
    if not os.path.exists(save_path): # 如果目录不存在，新建目录
        os.makedirs(save_path)
    img_path = os.path.join(save_path,img_name+".jpg")

    img = ImageGrab.grab()  # 操作系统的截屏（非浏览器）
    img.save(img_path, "jpeg")  # 截屏图片保存本地


# ==================================== 【键盘、鼠标事件】 =============================================

def keyboard_sys(first_key=None,second_key=None,third_key=None):
    VK_CODE = {
        'backspace': 0x08,
        'tab': 0x09,
        'clear': 0x0C,
        'enter': 0x0D,
        'shift': 0x10,
        'ctrl': 0x11,
        'alt': 0x12,
        'pause': 0x13,
        'caps_lock': 0x14,
        'esc': 0x1B,
        'spacebar': 0x20,
        'page_up': 0x21,
        'page_down': 0x22,
        'end': 0x23,
        'home': 0x24,
        'left_arrow': 0x25,
        'up_arrow': 0x26,
        'right_arrow': 0x27,
        'down_arrow': 0x28,
        'select': 0x29,
        'print': 0x2A,
        'execute': 0x2B,
        'print_screen': 0x2C,
        'ins': 0x2D,
        'del': 0x2E,
        'help': 0x2F,
        '0': 0x30,
        '1': 0x31,
        '2': 0x32,
        '3': 0x33,
        '4': 0x34,
        '5': 0x35,
        '6': 0x36,
        '7': 0x37,
        '8': 0x38,
        '9': 0x39,
        'a': 0x41,
        'b': 0x42,
        'c': 0x43,
        'd': 0x44,
        'e': 0x45,
        'f': 0x46,
        'g': 0x47,
        'h': 0x48,
        'i': 0x49,
        'j': 0x4A,
        'k': 0x4B,
        'l': 0x4C,
        'm': 0x4D,
        'n': 0x4E,
        'o': 0x4F,
        'p': 0x50,
        'q': 0x51,
        'r': 0x52,
        's': 0x53,
        't': 0x54,
        'u': 0x55,
        'v': 0x56,
        'w': 0x57,
        'x': 0x58,
        'y': 0x59,
        'z': 0x5A,
        'numpad_0': 0x60,
        'numpad_1': 0x61,
        'numpad_2': 0x62,
        'numpad_3': 0x63,
        'numpad_4': 0x64,
        'numpad_5': 0x65,
        'numpad_6': 0x66,
        'numpad_7': 0x67,
        'numpad_8': 0x68,
        'numpad_9': 0x69,
        'multiply_key': 0x6A,
        'add_key': 0x6B,
        'separator_key': 0x6C,
        'subtract_key': 0x6D,
        'decimal_key': 0x6E,
        'divide_key': 0x6F,
        'F1': 0x70,
        'F2': 0x71,
        'F3': 0x72,
        'F4': 0x73,
        'F5': 0x74,
        'F6': 0x75,
        'F7': 0x76,
        'F8': 0x77,
        'F9': 0x78,
        'F10': 0x79,
        'F11': 0x7A,
        'F12': 0x7B,
        'F13': 0x7C,
        'F14': 0x7D,
        'F15': 0x7E,
        'F16': 0x7F,
        'F17': 0x80,
        'F18': 0x81,
        'F19': 0x82,
        'F20': 0x83,
        'F21': 0x84,
        'F22': 0x85,
        'F23': 0x86,
        'F24': 0x87,
        'num_lock': 0x90,
        'scroll_lock': 0x91,
        'left_shift': 0xA0,
        'right_shift ': 0xA1,
        'left_control': 0xA2,
        'right_control': 0xA3,
        'left_menu': 0xA4,
        'right_menu': 0xA5,
        'browser_back': 0xA6,
        'browser_forward': 0xA7,
        'browser_refresh': 0xA8,
        'browser_stop': 0xA9,
        'browser_search': 0xAA,
        'browser_favorites': 0xAB,
        'browser_start_and_home': 0xAC,
        'volume_mute': 0xAD,
        'volume_Down': 0xAE,
        'volume_up': 0xAF,
        'next_track': 0xB0,
        'previous_track': 0xB1,
        'stop_media': 0xB2,
        'play/pause_media': 0xB3,
        'start_mail': 0xB4,
        'select_media': 0xB5,
        'start_application_1': 0xB6,
        'start_application_2': 0xB7,
        'attn_key': 0xF6,
        'crsel_key': 0xF7,
        'exsel_key': 0xF8,
        'play_key': 0xFA,
        'zoom_key': 0xFB,
        'clear_key': 0xFE,
        '+': 0xBB,
        ',': 0xBC,
        '-': 0xBD,
        '.': 0xBE,
        '/': 0xBF,
        '`': 0xC0,
        ';': 0xBA,
        '[': 0xDB,
        '\\': 0xDC,
        ']': 0xDD,
        "'": 0xDE,
        '`': 0xC0
    }

    #键盘键按下
    def keyDown(keyName):
        win32api.keybd_event(VK_CODE[keyName], 0, 0, 0)

    #键盘键抬起
    def keyUp(keyName):
        win32api.keybd_event(VK_CODE[keyName], 0, win32con.KEYEVENTF_KEYUP, 0)

    if first_key:
        keyDown(first_key)
    if second_key:
        keyDown(second_key)
    if third_key:
        keyDown(third_key)

    if third_key:
        keyUp(third_key)
    if second_key:
        keyUp(second_key)
    if first_key:
        keyUp(first_key)

def getText(): # 获取剪切板内容
    win32clipboard.OpenClipboard()
    text = win32clipboard.GetClipboardData(win32con.CF_TEXT)
    win32clipboard.CloseClipboard()
    return text

def setText(str_text): # 设置剪切板内容
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, str_text)
    win32clipboard.CloseClipboard()

def right_click(element): # 鼠标右键
    global driver
    ActionChains(driver).move_to_element(element).perform() # 模拟鼠标悬停于元素上方
    ActionChains(driver).context_click(element).perform()  # 模拟点击鼠标右键
    time.sleep(0.5)

def left_click(element): # 鼠标左键
    global driver
    ActionChains(driver).click_and_hold(element).perform()  # 按住鼠标左键
    ActionChains(driver).release(element).perform()  # 鼠标的左键，松开

def hovering(element): # 鼠标悬停（某个元素上方）
    global driver
    ActionChains(driver).move_to_element(element).perform() # 模拟鼠标悬停于元素上方


def wait_implicit(times): # 隐式等待，页面中所有需要定位的元素，必须在设定时间内找到，不然会抛异常
                          # 隐性等待对整个driver的周期都起作用，所以只要设置一次即可
    global driver
    driver.implicitly_wait(times)

def wait_visibility_by_element(element,times=10): #指定的时间内，判断元素是否存在，并返回元素
    global driver
    wait = WebDriverWait(driver, times, 0.2)
    element = wait.until(EC.visibility_of(element))
    return element

def wait_visibility_by_xpath(xpath_exp,times=10): #指定的时间内，通过xpath定位，判断元素是否存在，并返回元素
    global driver
    wait = WebDriverWait(driver, times, 0.2)
    element = wait.until(EC.visibility_of(driver.find_element_by_xpath(xpath_exp)))
    return element

def wait_clickable(By,locator_path,times=10): #指定的时间内，判断按钮能否被点击，并返回元素
    '''
    :param By: By.ID By.XPATH By.NAME等

    By所支持的定位器的分类
    CLASS_NAME = 'class name'
    CSS_SELECTOR = 'css selector'
    ID = 'id'
    LINK_TEXT = 'link text'
    NAME = 'name'
    PARTIAL_LINK_TEXT = 'partial link text'
    TAG_NAME = 'tag name'
    XPATH = 'xpath'

    :param locator_path: 值或者表达式
    :param times: 显示等待的时间
    :return:
    '''
    global driver
    wait = WebDriverWait(driver, times, 0.2)
    element = wait.until(EC.element_to_be_clickable((By,locator_path)))
    return element







