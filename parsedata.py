import pyautogui
import keyboard
from datetime import datetime
import colvir
from colvir import Auto_Colvir
import os
from PIL import Image,ImageOps,ImageEnhance
import time
from datetime import datetime
import pytesseract
from pytesseract import Output
from settings import datasend, image_dict
import json
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
import json
from openpyxl.styles import Color, PatternFill, Border, Alignment,Font, NamedStyle
from openpyxl.styles import colors
from docx import Document
import re
from decimal import Decimal


time.sleep(4)

pytesseract.pytesseract.tesseract_cmd = datasend.tesseract_dir
colvir = Auto_Colvir()


keyboard.press_and_release('alt+1')
keyboard.write("договоры с ТСП")
keyboard.press_and_release('enter')

colvir.window_find("Фильтр")
clear_form = colvir.find_image(image_dict['clear_form'], 0.7)

colvir.click_text("клиента", 0, width_scr_div=1, height_scr_div=1, resize=2)
time.sleep(0.5)
pos = pyautogui.position()
time.sleep(0.5)
pyautogui.moveTo(pos.x+220, pos.y)
pyautogui.click()

colvir.window_find("Фильтр")
time.sleep(0.5)
clear_form = colvir.find_image(image_dict['clear_form'], 0.7)


colvir.fill_form('', '', 'КОСВИГАРАНТ ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ')
time.sleep(0.5)
ok_button = colvir.find_image(image_dict['ok_button'], 0.7)
readdata = colvir.find_image(image_dict['readdata'], 0.7, click=False, ignore=True)


while readdata:
    time.sleep(2)
    readdata = colvir.find_image(image_dict['readdata'], 0.6, click=False, ignore=True)
    time.sleep(2)
    if readdata == None:
        time.sleep(1)
        keyboard.press_and_release('enter')
        break
    
time.sleep(2)

colvir.window_find("Фильтр")
colvir.fill_form('', '', '','', '', 'MT.BO.3.1', 0.5, '', 0.5, 'Зарегистрирован в ПЦ', 0.5)
ok_button = colvir.find_image(image_dict['ok_button'], 0.7)

time.sleep(2)

colvir.window_find("Договоры с ТСП")
keyboard.press_and_release('enter')

colvir.window_find("Договор с ТСП")
colvir.click_text("начала", 0, width_scr_div=1, height_scr_div=1, resize=3, language='rus')
beginning_date = colvir.copy_selection(50,0)

time.sleep(0.5)

colvir.click_text("Продукт", 0, width_scr_div=1, height_scr_div=1, resize=3, language='rus')
agr_number = colvir.copy_selection(90,40)
agr_number = "ACQ-" + agr_number

time.sleep(0.5)

colvir.click_text("Подключенные", 0, width_scr_div=1, height_scr_div=1, resize=3)
email = colvir.find_text("@", 0, width_scr_div=1, height_scr_div=1, resize=6, language='eng')
email = email.replace("|", "")

time.sleep(1)

keyboard.press_and_release('esc')

time.sleep(1)

colvir.find_image(image_dict['acc_agr'], 0.8)

debt = colvir.find_image(image_dict['6733'], 0.6, click=False)

print(debt)

pyautogui.moveTo(debt.left+350, debt.top+30)

pos = pyautogui.position()

im_center = pyautogui.screenshot("scren1.png", region=(pos.x, pos.y, 150, 25))
resize = 2
im_center = im_center.resize((im_center.width*resize, im_center.height*resize))
sharpness = ImageEnhance.Sharpness(im_center)
im_center = sharpness.enhance(1.5)
res = pytesseract.image_to_data(ImageOps.grayscale(im_center), lang='rus', output_type=Output.DICT)

amount = None

for item in res['text']:
    if item != '':
        print(item)
        amount = Decimal(item)

print(str(amount))

keyboard.press_and_release('esc')


