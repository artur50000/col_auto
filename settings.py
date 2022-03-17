from dataclasses import dataclass
from logsettings import logger
import os
from PIL import Image

@dataclass
class SendData:
    tesseract_dir: str
    picfolder: str
    screenshot: str

tesseract_dir = r'C:\Users\sotsenko\AppData\Local\Programs\Tesseract-OCR\tesseract'
picfolder = "D:/scripts/informcorp/pics/"
screenshot = "D:/scripts/informcorp/"

datasend = SendData(tesseract_dir, picfolder, screenshot)
image_dict = {}
image_dict = {file[:-4]:Image.open(f"{picfolder}{file}")
              for file in os.listdir(picfolder)
              if file.endswith(".png")
              or file.endswith("PNG")}