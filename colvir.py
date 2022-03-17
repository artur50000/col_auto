import pyautogui
import sys
import os
import time
import cv2
#from settings import folder
from PIL import Image,ImageOps,ImageEnhance
import pytesseract
from pytesseract import Output
import keyboard
import functools
from listexceptions import NotFoundException
from logsettings import logger
from settings import datasend
from tkinter import Tk

image_path = datasend.screenshot

class Auto_Colvir():
    path = image_path
    
    def retry_scope(number_tries, screen_path):
        def retry(func):
            @functools.wraps(func)
            def inner(*args, **kwargs):
                for i in range(0, number_tries):
                    try:
                        #things I need to do
                        result = func(*args, **kwargs)
                        return result
                    except NotFoundException as e:
                        if i < number_tries-1:
                            print(e)
                            print("Try #{} failed with NotFoundException: Sleeping for 2 secs before next try:".format(i))
                            time.sleep(1)
                            continue
                        else:
                            pyautogui.screenshot(f"{screen_path}screen.png")
                            logger.exception("process terminated")
                            sys.exit("process terminated")
                    break
            return inner
        return retry


    @retry_scope(number_tries=5, screen_path=path)
    def window_find(self, word, ignore=False):
        current_window = pyautogui.getActiveWindow()
        if current_window and current_window.title == word:
            return True
        elif ignore:
            time.sleep(0.5)
            return
        else:
            raise NotFoundException(word)


    @retry_scope(number_tries=5, screen_path=path)
    def find_image(self, image, conf, click=True, region_to_find=None, ignore=False):
        image_to_find = None

        if region_to_find:
            image_to_find = pyautogui.locateOnScreen(image, region=(region_to_find.left, region_to_find.top, region_to_find.width, region_to_find.height),confidence=conf)
        else:
            image_to_find = pyautogui.locateOnScreen(image, confidence=conf)
        if image_to_find and click:
            pyautogui.click(image_to_find, interval=0.5)
            return True
        elif image_to_find and click==False:
            return image_to_find
        elif not image_to_find and ignore:
            return
        else:
            raise NotFoundException(image)


    def word_dict(self, width_scr_div, height_scr_div, resize, language='rus'):
        time.sleep(0.5)
        current_window = pyautogui.getActiveWindow()
        x = current_window.left
        y = current_window.top
        width = current_window.width
        height = current_window.height
        im_center = pyautogui.screenshot(region=(x, y, width//width_scr_div, height//height_scr_div))
        im_center = im_center.resize((im_center.width*resize, im_center.height*resize))
        sharpness = ImageEnhance.Sharpness(im_center)
        im_center = sharpness.enhance(1.5)
        res = pytesseract.image_to_data(ImageOps.grayscale(im_center), lang=language, output_type=Output.DICT)
        #print(res['text'])

        return res, x, y


    @retry_scope(number_tries=5, screen_path=path)
    def click_text(self, word, shift, width_scr_div=1, height_scr_div=1, resize=1, language='rus'):
        res, x, y = self.word_dict(width_scr_div, height_scr_div, resize, language=language)  

        for item in res['text']:

            if word in item:
                i = res['text'].index(item)
                x1 = res["left"][i]
                y1 = res["top"][i]
                w = res["width"][i]
                h = res["height"][i]
                pyautogui.moveTo(x+x1/resize+shift,y+y1/resize)
                pyautogui.click()
                return True
        else:
            raise NotFoundException(word)


    def fill_form(self, *args):
        for item in args:
            if isinstance(item, str) and not item=="del" and len(item)>0:
                keyboard.write(item)
            elif isinstance(item, float):
                time.sleep(item)
            elif item == "del":
                keyboard.press_and_release('del')
            elif item =='':
                keyboard.press_and_release('tab')

    
    def window_close(self, image):
        while True:
            close_wind = pyautogui.locateOnScreen(image, confidence=0.6)
            pyautogui.click(close_wind, interval=0.5)
            time.sleep(1)
            if close_wind == None:
                break

    
    @retry_scope(number_tries=5, screen_path=path)
    def find_text(self, word, shift, width_scr_div=1, height_scr_div=1, resize=1, language='rus'):
        res, x, y = self.word_dict(width_scr_div, height_scr_div, resize, language=language)  

        for item in res['text']:

            if word in item:
                i = res['text'].index(item)
                x1 = res["left"][i]
                y1 = res["top"][i]
                w = res["width"][i]
                h = res["height"][i]
                #pyautogui.moveTo(x+x1/resize+shift,y+y1/resize)
                #pyautogui.click()
                return item
        else:
            raise NotFoundException(word)


    def copy_selection(self, shift_x, shift_y):
        time.sleep(0.5)
        pos = pyautogui.position()
        time.sleep(0.5)
        pyautogui.moveTo(pos.x+shift_x, pos.y+shift_y)
        pyautogui.click(clicks=2, interval=0.25)
        time.sleep(0.5)
        keyboard.press_and_release('ctrl+c')
        copied_text = Tk().clipboard_get()
        return copied_text
                    
