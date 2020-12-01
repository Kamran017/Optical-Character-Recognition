import cv2
import numpy as np
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import pandas as pd
 
#download tesseract exe and import it 
 pytesseract.pytesseract.tesseract_cmd = r"LOCATION OF TESSERACT EXE FILE"

#path of pdf file
path = "PATH OF YOUR PDF FILE"
#convert pdf pages to image format
pages = convert_from_path(path ,500)


#save pages 
counter=0
for page in pages:
    page.save(str(counter) + '.jpg', 'JPEG')
    counter+=1

#open txt file for writing texts from images
f = open("result.txt", "a")
#this loop iterate over pages cropped from pdf 
for i in range(counter):
    text = pytesseract.image_to_string(str(i) + ".jpg")#get text from image
    f.write(text)#write text to file
    
f.close()#close text file


#read txt file to pandas dataframe and convert to the excel file
df = pd.read_csv("result.txt", error_bad_lines=False,encoding='cp1252')
df.to_excel('filename.xlsx')#convert txt file to excel file


    










