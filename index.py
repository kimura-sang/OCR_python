from math import nan
import pyautogui
import pandas as pd
import time
import cv2
import numpy as np
import time
import mss
import easyocr
from fuzzywuzzy import process
import pyperclip
import sys
import multiprocessing
import keyboard
import os
import logging
import pywinauto
import math

# SWAPY will record the title and class of the window you want activated
app = pywinauto.application.Application()
print(app)
t, c = 'NoxPlayer', 'Qt5QWindowIcon'
handle = pywinauto.findwindows.find_windows(title=t)
if len(handle)<1:
    pyautogui.alert(text='noxを起動した後、再度実行してください。')

handle = handle[0]

print(pywinauto.handleprops.classname(handle))
print(pywinauto.handleprops.text(handle))
# SWAPY will also get the window
window = app.connect(handle=handle)
window = app.window(handle=handle)

# this here is the only line of code you actually write (SWAPY recorded the rest)
window.set_focus()
# window.maximize()
# print(window.width())
# manager = multiprocessing.Manager()
# current_row_list = manager.list()
# errors_list = manager.list()
# def _append_run_path():
#     if getattr(sys, 'frozen', False):
#         pathlist = []

#         # If the application is run as a bundle, the pyInstaller bootloader
#         # extends the sys module by a flag frozen=True and sets the app
#         # path into variable _MEIPASS'.
#         pathlist.append(sys._MEIPASS)

#         # the application exe path
#         _main_app_path = os.path.dirname(sys.executable)
#         pathlist.append(_main_app_path)

#         # append to system path enviroment
#         os.environ["PATH"] += os.pathsep + os.pathsep.join(pathlist)

#     logging.error("current PATH: %s", os.environ['PATH'])


# _append_run_path()

pyautogui.PAUSE = 0.1


def screenshot(top, left, swidth, sheight):
    with mss.mss() as sct:
        monitor = {"top": top, "left": left, "width": swidth, "height": sheight}
        im = sct.grab(monitor)
        mss.tools.to_png(im.rgb, im.size, output="monitor-1.png")
        img_rgb = np.array(im)
        img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
        return img_rgb


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.getcwd())
    # print(os.path.abspath(os.getcwd()))

    return os.path.join(os.path.abspath(os.getcwd()), relative_path)


def imagesearch(image, precision=0.7):
    screenWidth, screenHeight = pyautogui.size()
    # print(screenHeight)
    with mss.mss() as sct:
        im = sct.grab(sct.monitors[0])
        img_rgb = np.array(im)
        img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
        template = cv2.imread(image, 0)
        # template = np.array(template)
        # template = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
        if template is None:
            raise FileNotFoundError('Image file not found: {}'.format(image))
        # template.shape[::-1]
        # thresh = cv2.threshold(img_gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)#TM_CCOEFF_NORMED
        
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        # print(max_val)
        if max_val < precision:
            return [-1, -1]
        return max_loc


data = pd.read_excel (resource_path('temp.xlsx'), header=0) 
dict_data = list(data.to_dict(orient='records'))

ready1 = 0
x1, y1, x2, y2 = 0, 0, 0, 0
width, height = 566, 978
dis1 = [702-657, 97-60]
dis2 = [1168-657, 1005-60]
dis3 = [1182-657, 93-60]
pos = pos1 = pos2 = pos3 = pos5 = poscolor = posselect = posselect1 = pospre = pospre1 = posmypage = possetting = poslogout = posenter = pos_repair = [-1, -1]
errors = []
current_row = 0
pres = "北海道,青森県,岩手県,宮城県,秋田県,山形県,福島県,茨城県,栃木県,群馬県,埼玉県,千葉県,東京都,神奈川県,新潟県,富山県,石川県,福井県,山梨県,長野県,岐阜県,静岡県,愛知県,三重県,滋賀県,京都府,大阪府,兵庫県,奈良県,和歌山県,鳥取県,島根県,岡山県,広島県,山口県,徳島県,香川県,愛媛県,高知県,福岡県,佐賀県,長崎県,熊本県,大分県,宮崎県,鹿児島県,沖縄県"
pres = pres.split(",")
print(pres)

def on_move(x, y):
    print ("Mouse moved")


def on_scroll(x, y, dx, dy):
    print ("Mouse scrolled")


def addition(n):
    return n[1]


def logout():
    global posmypage, possetting, poslogout
    while(posmypage[0]<0):
        posmypage = imagesearch(resource_path("mypage.png"))
        time.sleep(0.7)
    # loadingSleep()
    time.sleep(3)
    pyautogui.click(posmypage[0] + 10, posmypage[1] + 10)
    # loadingSleep()
    time.sleep(5)
    while(possetting[0]<0):
        possetting = imagesearch(resource_path("setting.png"))
        time.sleep(0.7)
    pyautogui.click(possetting[0] + 10, possetting[1] + 10)
    time.sleep(0.7)
    while(poslogout[0]<0):
        poslogout = imagesearch(resource_path("logout.png"))
        time.sleep(0.7)
    pyautogui.click(poslogout[0] + 10, poslogout[1] + 10)
    time.sleep(0.7)
    pyautogui.press("tab")
    pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.press("enter")
    poslogged = [-1, -1]
    while(poslogged[0]<0):
        poslogged = imagesearch(resource_path("logged.png"))
        time.sleep(1)
    # loadingSleep()
    time.sleep(7)


# def error_row(row):    
#     highlight = 'background-color: lightcoral;'
#     default = ''
#     print("row")
#     print(row)
#     # must return one string per cell in this row
#     if row['アカウントID'] > row['パスワード']:
#         return [highlight, highlight]
#     else:
#         return [highlight, highlight]


def addError(current_row, type):
    errors.append([current_row, type])
    print("errors")
    print(errors)

def checkError():
    global pos, data, current_row, errors
    poserror = [-1, -1]
    poscount = 0
    while(poserror[0]<0 and poscount<3):
        poserror = imagesearch(resource_path("error.png"))
        poscount = poscount + 1
        print(poserror[0]<0 and poscount<2)
        time.sleep(0.5)
    if(poserror[0]<0):
        return False
    addError(current_row, "login")
    
    time.sleep(0.5)
    pyautogui.press("tab")
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.5)
    pyautogui.moveTo(pos[0] + dis1[0], pos[1] + dis1[1])
    pyautogui.click(pos[0] + dis1[0], pos[1] + dis1[1])
    # with xlsxwriter.Workbook('result.xlsx') as workbook:
    #     worksheet = workbook.get_worksheet_by_name("Sheet1")

    #     data_format1 = workbook.add_format({'bg_color': '#FFC7CE'})
    #     data_format2 = workbook.add_format({'bg_color': '#00C7CE'})

    #     for row in range(0, 10, 2):
    #         worksheet.set_row(row, cell_format=data_format1)
    #         worksheet.set_row(row + 1, cell_format=data_format2)

    return True


def writeExcel(errors, current_row):
    # global errors
    writer = pd.ExcelWriter(resource_path('result.xlsx'), engine='xlsxwriter')

    df = pd.DataFrame(data)

    # Write data to an excel
    df.to_excel(writer, sheet_name="Sheet1", startrow=0, index=False, header=True)
    print(list(df.to_dict(orient='records'))[1])

    # Get workbook
    workbook = writer.book

    # Get Sheet1
    worksheet = writer.sheets['Sheet1']
    worksheet.set_column(0, 0, 30)
    worksheet.set_column(1, 1, 20)
    worksheet.set_column(4, 4, 30)
    worksheet.set_column(5, 5, 20)
    worksheet.set_column(6, 6, 20)
    worksheet.set_column(7, 7, 20)
    worksheet.set_column(8, 8, 20)

    # add format for headers
    error_format_red = workbook.add_format({'bold': True, 'font_color': 'red'})
    error_format_blue = workbook.add_format({'bold': True, 'font_color': 'blue'})
    error_format_stop = workbook.add_format({'bold': True, 'bg_color': 'green'})
    error_format = error_format_red
    # error_format.set_font_name('Bodoni MT Black')
    # error_format.set_font_color('red')
    # Write the column headers with the defined format.
    for error in errors:
        if(error[1]=="login"):
            addcol = 0
            error_format = error_format_red
        elif error[1]=="pre" :
            addcol = 3
            error_format = error_format_blue
        elif error[1]=="shop" :
            addcol = 4
            error_format = error_format_blue
        for col_num, value in enumerate(df.columns[[1]]):
            worksheet.write(error[0], col_num+addcol, list(df.to_dict(orient='records'))[error[0]][value], error_format)
    for col_num, value in enumerate(df.columns[list(range(8))]):
            worksheet.write(current_row, col_num, list(df.to_dict(orient='records'))[current_row][value], error_format_stop)
    writer.close()
    return


def loadingSleep():
    time.sleep(7)
    # pos_loading = [-1, -1]
    # time.sleep(1)
    # while(pos_loading[0]<0):
    #     pos_loading = imagesearch(resource_path("loading.png"))
    #     time.sleep(0.5)
    # pos_loading = [1, 1]
    # while(pos_loading[0]>0):
    #     pos_loading = imagesearch(resource_path("loading.png"))
    #     time.sleep(0.5)
    # print(pos_loading)
    return True

def loadingSleepAgain():
    time.sleep(4)
    # time.sleep(1)
    # pos_loading = [1, 1]
    # while(pos_loading[0]>0):
    #     pos_loading = imagesearch(resource_path("loading.png"))
    #     time.sleep(0.5)
    # print(pos_loading)
    return True

def dragScreen(screenHeight):
    print(pos3)
    pyautogui.moveTo(pos3[0], pos3[1]-40)
    pyautogui.dragTo(pos3[0], pos3[1]-40 - screenHeight*0.8, 1)

def dragScreenUp(screenWidth, screenHeight):
    pyautogui.moveTo(screenWidth/2, screenHeight/2-100)
    pyautogui.dragTo(screenWidth/2, screenHeight-100, 1)

def start(excel_datum):
    if pd.isnull(excel_datum["アカウントID"]) or pd.isnull(excel_datum["パスワード"]) or pd.isnull(excel_datum["機種"]) or pd.isnull(excel_datum["都道府県"]) or pd.isnull(excel_datum["店舗名"]) or pd.isnull(excel_datum["指名"]) or pd.isnull(excel_datum["シメイ"]) or pd.isnull(excel_datum["電話番号"]):
        print("failed")
        return
    print("started")
    screenWidth, screenHeight = pyautogui.size()
    global pos, pos1, pos2, pos3, pos, pres, posselect, posselect1, width, height, dis1, dis2, dis3, pospre, pospre1, posmypage, possetting, poslogout, posenter, pos_repair
    pos_count = 0
    max_count = 7
    while(pos[0]<0 and pos_count<max_count):
        pos = imagesearch(resource_path("temp.png"))
        time.sleep(0.7)
        pos_count += 1
    # if pos_count == max_count:
    #     raise ValueError('A very specific bad thing happened.')
    pyautogui.moveTo(pos[0] + dis1[0], pos[1] + dis1[1])
    pyautogui.click(pos[0] + dis1[0], pos[1] + dis1[1])
    pos_count = 0
    while(pos1[0]<0 and pos_count<max_count):
        pos1 = imagesearch(resource_path("temp1.png"))
        time.sleep(0.7)
        pos_count += 1
    # while(pos2[0]<0):
    #     pos2 = imagesearch("temp2.png")
    if(pos_count==0):
        time.sleep(0.7)
    pyautogui.click(pos1[0] + 400, pos1[1] + 60)
    pyautogui.click(pos1[0] + 400, pos1[1] + 60)
    # for _ in range(30):
    pyautogui.keyDown("backspace")
    time.sleep(2)
    pyautogui.keyUp("backspace")
    pyautogui.write(excel_datum["アカウントID"])
    pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.write(excel_datum["パスワード"])
    while(pos3[0]<0):
        pos3 = imagesearch(resource_path("temp3.png"))
        time.sleep(0.7)
    pyautogui.click(pos3[0] + 10, pos3[1] + 10)
    # if False:
    #     pyautogui.click(pos2[0] + 10, pos2[1] + 60)
    # pyautogui.press("tab")
    # pyautogui.press("tab")
    # pyautogui.press("tab")
    # pyautogui.press("enter")
    # loadingSleep()
    time.sleep(7)
    isError = checkError()
    if(isError==True):
        return

    # poslogged = [-1, -1]
    # pos_count = 0
    # while(poslogged[0]<0 and pos_count<max_count):
    #     poslogged = imagesearch(resource_path("logged.png"))
    #     time.sleep(1)
    #     pos_count += 1
    # time.sleep(3)

    currentMouseX, currentMouseY = pyautogui.position()
    dragScreen(screenHeight)
    time.sleep(0.7)
    pos4 = [-1, -1]
    while(pos4[0]<0):
        pos4 = imagesearch(resource_path("temp4.png"))
        if(pos4[0]<0):
            dragScreen(screenHeight)
        time.sleep(0.7)
    pyautogui.click(pos4[0] + 30, pos4[1] + 70)
    # loadingSleep()
    time.sleep(1)
    pos_play = [-1]
    while(pos_play[0]<0):
        pos_play = imagesearch(resource_path("play.png"))
        time.sleep(0.7)
    pos_count = 0
    while(pos_repair[0]<0 and pos_count<3):
        pos_repair = imagesearch(resource_path("repair.png"))
        time.sleep(0.7)
        pos_count += 1
    if(pos_repair[0]>0):
        pyautogui.alert('ただいま抽選販売のシステムメンテナンスを実施しております。')
        raise ValueError('システムメンテナンス')
    # pyautogui.dragTo(currentMouseX, currentMouseY - screenHeight)
    pyautogui.press("tab")
    time.sleep(0.5)
    pyautogui.press("down")
    time.sleep(0.5)
    pyautogui.press("down")
    time.sleep(0.5)
    pyautogui.press("down")
    # while(pos5[0]<0):
    #     pyautogui.press("tab")
    #     pos5 = imagesearch("temp5.png")
    #     time.sleep(0.5)
    # pyautogui.click(pos5[0] + 30, pos5[1] + 20)
    time.sleep(0.5)
    pyautogui.press("enter")
    # loadingSleep()
    time.sleep(7)
    posalready = imagesearch(resource_path("already.png"))
    if posalready[0]>0:
        print("already registered")
        # pyautogui.moveTo(pos[0]+50, screenHeight*0.8)
        # currentMouseX, currentMouseY = pyautogui.position()
        # pyautogui.dragTo(currentMouseX, 1)
        # time.sleep(0.7)
        # pyautogui.click(currentMouseX, screenHeight*0.85)
        time.sleep(0.5)
        pyautogui.press("tab")
        time.sleep(0.3)
        pyautogui.press("down")
        time.sleep(0.3)
        pyautogui.press("down")
        time.sleep(0.3)
        pyautogui.press("down")
        pyautogui.press("down")
        time.sleep(0.7)
        pyautogui.press("enter")
        # loadingSleep()
        time.sleep(5)
        logout()
        return
    # if(False):
    #     while(poscolor[0]<0):
    #         poscolor = imagesearch("color.png")
    #     pyautogui.click(poscolor[0] + 10, poscolor[1] + 10)
    pyautogui.press("tab")
    if excel_datum["機種"] == "黒":
        print("黒")
        time.sleep(0.7)
        pyautogui.press("down")
        time.sleep(1)
        pyautogui.press("enter")
        # time.sleep(0.7)
        # pyautogui.press("tab")

        # poscolor = [-1, -1]
        # while(poscolor[0]<0):
        #     poscolor = imagesearch("color.png")
        #     print(poscolor)
        # pyautogui.click(poscolor[0] + 10, poscolor[1] + 10)
        # time.sleep(0.7)
        # pyautogui.press("enter")
        # pyautogui.press("enter")
    else:
        # pyautogui.press("tab")
        print("!黒")
    # pyautogui.moveTo(pos[0]+50, screenHeight - 70)
    # currentMouseX, currentMouseY = pyautogui.position()
    # pyautogui.dragTo(currentMouseX, currentMouseY - screenHeight*0.8, 1)
    # while(posselect[0]<0):
    #     posselect = imagesearch("select.png")
    # pyautogui.click(posselect[0] + 10, posselect[1] + 10)
    # while(posselect1[0]<0):
    #     posselect1 = imagesearch("select1.png")
    # pyautogui.click(posselect1[0] + 10, posselect1[1] + 10)
    # while(pospre[0]<0):
    #     pospre = imagesearch("pre.png")
    # pyautogui.click(pospre[0] + 10, pospre[1] + 10)
    time.sleep(0.5)
    pos_option = [-1, -1]
    dragScreen(screenHeight)
    time.sleep(0.7)
    while(pos_option[0]<0):
        pos_option = imagesearch(resource_path("selectoption.png"))
        time.sleep(0.7)
    pyautogui.click(pos_option[0] + 10, pos_option[1] + 10)
    # time.sleep(0.8)
    # pyautogui.press("enter")
    time.sleep(0.8)
    pyautogui.press("tab")
    pyautogui.press("tab")
    time.sleep(0.8)
    pyautogui.press("enter")
    time.sleep(0.8)
    # while(pospre[0]<0):
    #     pospre = imagesearch("pre.png")
    # pyautogui.click(pospre[0] + 10, pospre[1] + 30)
    pyautogui.press("tab")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    # preindex = pres.index(excel_datum["都道府県"])
    preindex = 0
    for x in pres:
        print(x)
        if x.find(excel_datum["都道府県"]) > -1:
            break
        preindex += 1
    print(preindex)
    if(preindex == len(pres)):
        addError(current_row, "pre")
        pyautogui.click(pos[0] + 10, pos[1] + 10)
        time.sleep(0.7)
        pyautogui.click(pos[0] + 10, pos[1] + 10)
        time.sleep(0.7)
        # loadingSleep()
        logout()
        return
    for _ in range(int(preindex+1)):
        pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.press("enter")
    # loadingSleep()
    pos_branch_option = [-1, -1]
    while(pos_branch_option[0]<0):
        pos_branch_option = imagesearch(resource_path("branch_select.png"))
        time.sleep(0.7)
    time.sleep(3)
    pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.press("tab")
    time.sleep(0.7)
    # if excel_datum["機種"] == "黒":
    #     pyautogui.press("tab")
    #     time.sleep(0.7)
    # if excel_datum["機種"] == "黒":
    #     pyautogui.press("tab")
    #     time.sleep(0.7)
    pyautogui.press("enter")
    shop_read_count = 0
    while shop_read_count<7:
        time.sleep(1)
        screenimage = screenshot(80, 700, 380, 940)
        reader = easyocr.Reader(['ja']) 
        bounds = reader.readtext(screenimage)
        shop_lists = list(map(addition, list(bounds)))
        print("here")
        print(shop_lists)

        highest = process.extractOne(excel_datum["店舗名"],shop_lists)

        print(highest)

        if highest[1]>40:
            max_index = shop_lists.index(highest[0])
            print(max_index)

            time.sleep(0.7)
            for _ in range(max_index):
                pyautogui.press("tab")
            time.sleep(0.7)
            pyautogui.press("enter")
            break
        pyautogui.moveTo(pos[0]+50, screenHeight - 90)
        currentMouseX, currentMouseY = pyautogui.position()
        pyautogui.dragTo(currentMouseX, currentMouseY - screenHeight*0.75, 1)
        shop_read_count += 1

    if(shop_read_count >= 7):
        pyautogui.press("esc")
        addError(current_row, "shop")
        time.sleep(0.7)
        pyautogui.click(pos[0] + 10, pos[1] + 10)
        time.sleep(3)
        pyautogui.click(pos[0] + 10, pos[1] + 10)
        time.sleep(2)
        pyautogui.click(pos[0] + 10, pos[1] + 10)
        time.sleep(0.7)
        # loadingSleep()
        logout()
        return


    time.sleep(0.7)
    pyautogui.press("tab")
    pyperclip.copy(excel_datum["指名"])
    time.sleep(0.7)
    pyautogui.hotkey("ctrl", "v")
    # pyautogui.write(str(excel_datum["指名"]))
    time.sleep(0.7)
    pyautogui.press("tab")
    time.sleep(0.7)
    pyperclip.copy(excel_datum["シメイ"])
    pyautogui.hotkey("ctrl", "v")
    # pyautogui.write(str(excel_datum["シメイ"]))
    time.sleep(0.7)
    pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.write(str(excel_datum["電話番号"]))
    pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.write(str(excel_datum["電話番号"]))
    pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.write(excel_datum["アカウントID"])
    pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.write(excel_datum["アカウントID"])
    time.sleep(0.7)
    pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.press("tab")
    # time.sleep(1)
    # pyautogui.press("enter")
    posagree = [-1, -1]
    while posagree[0]<0:
        posagree = imagesearch(resource_path("agree.png"))
        time.sleep(0.5)
    pyautogui.click(posagree[0] + 10, posagree[1] + 10)
    
    time.sleep(0.7)
    pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.press("enter")
    # loadingSleep()
    time.sleep(3)

    # pyautogui.press("tab")
    # pyautogui.press("down")
    # time.sleep(1)
    # pyautogui.press("tab")
    # time.sleep(0.5)
    
    # pyautogui.press("enter")

    # print(pos[0]+50)
    # print(screenHeight - 90)
    # pyautogui.moveTo(pos[0]+50, screenHeight - 90)
    # currentMouseX, currentMouseY = pyautogui.position()
    # pyautogui.dragTo(currentMouseX, currentMouseY - screenHeight*0.75)
    pyautogui.press("tab")
    pyautogui.press("tab")
    time.sleep(0.5)
    posenter = [-1, -1]
    while posenter[0]<0:
        pyautogui.press("tab")
        time.sleep(0.5)
        posenter = imagesearch(resource_path("enter.png"))
        time.sleep(0.5)
    pyautogui.click(posenter[0] + 30, posenter[1] + 20)
    time.sleep(2)
    pyautogui.press("tab")
    pyautogui.press("tab")
    time.sleep(0.7)
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.press("tab")
    pyautogui.press("tab")
    time.sleep(0.5)
    pyautogui.press("down")
    pyautogui.press("down")
    pyautogui.press("down")
    pyautogui.press("enter")
    time.sleep(1)
    logout()
    return


# @process
def myrun():
    # for _ in range(5):
        # current_row_list.append(1)
    # print(current_row_list)
    # return
    screenWidth, screenHeight = pyautogui.size()
    if screenHeight!=1080:
        pyautogui.alert("画面の解像度を1920x1080に設定してください。")
        return
    global current_row, errors
    print("start here")
    # start_index = pyautogui.prompt(text='アプリがログアウトされたのか確認してください。\n開始番号を入力してください', title='確認する' , default='2')
    start_index = 2
    print(start_index)
    if(start_index is None):
        return
    current_row = max(int(start_index) - 2, 0)
    if(current_row>0):
        for _ in range(current_row):
            dict_data.pop(0)
    posback = [1, 1]
    while posback[0]>0:
        posback = imagesearch(resource_path("back.png"))
        print(posback)
        if(posback[0]>100):
            pyautogui.click(posback[0] + 10, posback[1] + 10)
        time.sleep(0.5)
    dragScreenUp(screenWidth, screenHeight)
    logincount = 0
    poslogin = [-1, -1]
    while poslogin[0]<0 and logincount<3:
        print(poslogin)
        poslogin = imagesearch(resource_path("login.png"))
        logincount += 1
        time.sleep(0.5)
    print(logincount)
    if(logincount>2):
        logout()
    time.sleep(1)
    for x in dict_data:
        print(x)
        print(current_row)
        start(x)
        current_row += 1
    writeExcel(errors, current_row)
    print("finished")


def mystop():
    global errors, current_row
    print("interupted")
    writeExcel(errors, current_row)
    # h.terminate()  # sends a SIGTERM
    # pyautogui.alert('操作が終了しました。')


if __name__ == "__main__":  # confirms that the code is under main function
    try:
        myrun()
        # proc = multiprocessing.Process(target=myrun, args=())
        # proc.start()
        # h = myrun
        # h.start()
        # h.start()
        # Terminate the process
    except:
        print("ended with error")
        print(errors)
        print(current_row)
        writeExcel(errors, current_row)
    else:
        print("ended without errors")
        # keyboard.add_hotkey('ctrl+space', lambda: mystop())
        
        # print(x1, y1, x2, y2)
        # screenWidth, screenHeight = x2-x1, y2-y1



    # currentMouseX, currentMouseY = pyautogui.position() # Get the XY position of the mouse.

# with Listener(on_click = on_click) as listener:
#   listener.join()

# screenWidth, screenHeight = pyautogui.size() # Get the size of the primary monitor.

# currentMouseX, currentMouseY = pyautogui.position() # Get the XY position of the mouse.
# pyautogui.alert('This displays some text with an OK button.')
# pyautogui.moveTo(0, 0) # Move the mouse to XY coordinates.


# pyautogui.click()          # Click the mouse.
# pyautogui.click(100, 200)  # Move the mouse to XY coordinates and click it.
# # pyautogui.click('button.png') # Find where button.png appears on the screen and click it.

# pyautogui.move(0, 10)      # Move mouse 10 pixels down from its current position.
# pyautogui.doubleClick()    # Double click the mouse.
# pyautogui.moveTo(500, 500, duration=2, tween=pyautogui.easeInOutQuad)  # Use tweening/easing function to move mouse over 2 seconds.

# pyautogui.write('Hello world!', interval=0.25)  # type with quarter-second pause in between each key
# pyautogui.press('esc')     # Press the Esc key. All key names are in pyautogui.KEY_NAMES

# pyautogui.keyDown('shift') # Press the Shift key down and hold it.
# pyautogui.press(['left', 'left', 'left', 'left']) # Press the left arrow key 4 times.
# pyautogui.keyUp('shift')   # Let go of the Shift key.

# pyautogui.hotkey('ctrl', 'c') # Press the Ctrl-C hotkey combination.

# pyautogui.alert('This is the message to display.') # Make an alert box appear and pause the program until OK is clicked.

# distance = 200
# while distance > 0:
#         pyautogui.drag(distance, 0, duration=0.5)   # move right
#         distance -= 5
#         pyautogui.drag(0, distance, duration=0.5)   # move down
#         pyautogui.drag(-distance, 0, duration=0.5)  # move left
#         distance -= 5
#         pyautogui.drag(0, -distance, duration=0.5)  # move up

# def on_click(x, y, button, pressed):
#     if pressed:
#         print (x, y)
#         global ready1, x1, y1, x2, y2
#         if ready1 < 8 :
#             print(ready1)
#             x1, y1 = x, y
#             ready1 += 1
#         else :
#             print(ready1)
#             x2, y2 = x, y
#             listener.stop()
#             start()

# pospre1 = imagesearch("p"+"prenum"+".png")
    # while pospre1[0]<0:
    #     pyautogui.moveTo(pos[0]+50, screenHeight - 70)
    #     currentMouseX, currentMouseY = pyautogui.position()
    #     pyautogui.dragTo(currentMouseX, currentMouseY - screenHeight*0.8, 1)
    #     pospre1 = imagesearch("p"+"prenum"+".png")