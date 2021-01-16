#--coding:utf-8--
# @Time    : 2021/1/9/009 19:44
# @Author  : panyuangao
# @File    : handle_email_126.py
# @PROJECT : Selenium
from selenium_api import *
import random,string

def login(account,password): # 登录126邮箱
    global driver
    driver = browser("chrome")
    visit("https://www.126.com/")
    maximize()
    frame = xpath_find("//iframe[contains(@id,'x-URS-iframe')]") # 定位登录框iframe
    switch_frame(frame) # 切入登录框iframe
    input_box_account = name_find('email') # 定位用户名输入框
    input_box_password = name_find('password') # 定位密码输入框
    send_keys(input_box_account,account) # 输入账号
    send_keys(input_box_password,password) # 输入密码
    button_login = id_find('dologin') # 定位登录按钮
    click(button_login) # 点击登录按钮
    driver.implicitly_wait(10) # 隐式等待10s

def logout(): # 退出126邮箱
    global driver
    my_msg = xpath_find("//span[@id='spnUid']") # 定位显示邮箱信息的位置
    click(my_msg)
    sleep(0.5) # 等待0.5秒，以免“退出”按钮未展示
    button_quit = wait_clickable(By.XPATH,"//div[.='退出']",3) # 显式等待“退出”按钮
    click(button_quit) # 点击退出按钮
    assert "重新登录" in page_source()

def add_contact(): # 添加联系人
    global driver
    address_book = xpath_find("//div[.='通讯录']") # 定位通讯录按钮
    click(address_book) # 点击“通讯录”按钮

    new_contact = wait_clickable('xpath',"//span[.='新建联系人']")
    click(new_contact) # 点击“新建联系人”按钮

    name = wait_clickable("id",'input_N',3) # 隐式等待3s，并返回姓名元素
    email = xpath_find("//div[@id='iaddress_MAIL_wrap']//input")
    phone = xpath_find("//div[@id='iaddress_TEL_wrap']//input")
    button_confirm = xpath_find("//span[.='确 定']")
    star_target = xpath_find("//span[.='设为星标联系人']")

    send_keys(name,"联系人%s" %(str(random.randint(0,1000)))) #填写联系人姓名
    send_keys(email,"%s@126.com" %(''.join(random.sample(string.ascii_letters,5)))) # 填写联系人email
    # click(star_target) # 点击选择，设为“星标联系人”
    send_keys(phone,"155%s" %(str(random.randint(10000000,99999999)))) # 填写手机号码
    click(button_confirm) # 点击“确定”按钮
    try:
        hint = wait_clickable("xpath","//span[.='关 闭']",1) # 隐式等待1s，提示窗口的关闭按钮
        click(hint) # 点击提示窗口的关闭按钮
    except:
        print("首次添加联系人，不会弹出提示窗")

def send_email():
    home_page = xpath_find("//li[@title='首页']") # 点击“首页”，如果不在这个页面，会定位不到“写信”按钮
    click(home_page)
    write_button = xpath_find("//span[.='写 信']")
    click(write_button) # 点击“写信”按钮
    receiver= xpath_find('//input[@aria-label="收件人地址输入框，请输入邮件地址，多人时地址请以分号隔开"]')
    send_keys(receiver,"youngboss2020@126.com") # 收件人邮箱地址
    theme = xpath_find('//input[@maxlength="256"]')
    send_keys(theme,"发给自己的测试邮件") # 填写邮件标题
    enclosure = xpath_find("//input[@type='file']") #定位添加附件按钮
    send_keys(enclosure,"e:\\a.txt") # 添加附件

    iframe = xpath_find('//iframe[@tabindex="1"]') # 定位邮件内容副文本框iframe
    switch_frame(iframe) # 切入iframe
    text_body = xpath_find("//title[.='编辑邮件正文']/../../body")
    send_keys(text_body,123456) #邮件正文写入内容
    switch_default() #切出frame

    send_button = xpath_find('//span[.="发送"]')
    click(send_button) # 点击“发送按钮”
    sleep(5)

if __name__ == '__main__':
    login('youngboss2020', 'Wy123456') #登录邮箱
    add_contact() # 添加联系人
    send_email() # 发送邮件给自己（标题、正文、附件）
    logout() # 退出邮箱
    quit() # 退出浏览器




