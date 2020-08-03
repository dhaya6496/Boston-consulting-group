import os
import io
import pytesseract
import logging
from wand.image import Image as wi
import gc
import re
import pandas as pd
from PIL import Image
from logging.handlers import RotatingFileHandler

name_list = []
director_list = []
date_list = []

source = 'C:/Users/lenovo/Desktop/Boston consulting group/Board_resolution/Input'
lst = os.listdir(source)
for pdf in (lst):
    file_path = source + "/" + pdf
    def Get_text_from_image(pdf_path):
        pdf=wi(filename=pdf_path,resolution=300)
        pdfImg=pdf.convert('jpeg')
        imgBlobs=[]
        
        for img in pdfImg.sequence:
            page=wi(image=img)
            imgBlobs.append(page.make_blob('jpeg'))
        
        for imgBlob in imgBlobs:
            im=Image.open(io.BytesIO(imgBlob))
            text=pytesseract.image_to_string(im,lang='eng')
        return (text)
    text = Get_text_from_image(file_path)
    try:
        logger = logging.getLogger(f'C:/Users/lenovo/Desktop/Boston consulting group/Board_resolution/Logs/{pdf}.log')
        hdler = RotatingFileHandler(f'C:/Users/lenovo/Desktop/Boston consulting group/Board_resolution/Logs/{pdf}.log', maxBytes=1000)
        logger.addHandler(hdler)
        logger.setLevel(logging.INFO)
        formatter = logging.Formatter("%(asctime)s  %(message)s","%Y-%m-%d %H:%M")
        hdler.setFormatter(formatter)
        logger.info(f'Opened {file_path}') 

        name = re.search('Mr. [A-Za-z]* [A-Za-z]*|Mr. [A-Za-z]*([\s\S]*?)[A-Z] [A-za-z]*|Mr. [A-Za-z]*\n[A-Za-z]*(.*)? Muthoot|Mr. [A-Z]*.*[A-Z]{8}|Mr. [A-Za-z]* [A-Za-z]*',text).group()
        position = re.search('change in designation|appointment|resignation',text).group()
        director_string = re.search('from Designated([\s\S]*?)\nDirector.|Additional(.*)?[A-Za-z]{7}|from(.*)directorships|as(.*)?director|from([\s\S]*?)CEO',text).group()
        
        name_of_applicant = name+' '+position+' '+director_string

        director_type = re.search('Mr. [A-Za-z]*(.*)? Company|Mr. [A-Za-z]*(.*)? Whole-time Director|Mr. [A-Z]*(.*)?Director|Mr.[A-Za-z]*(.*)?CEO',text).group()
    
        date = re.search('[A-Z]* \d{2}, \d{4}| [A-Z]{4} \d{2},\n\d{4}| \d [A-Z]{2} \d{4}|\d(.*)[A-Z]{2}, \d{4}|[0-9]{2}\.[0-9]{2}\.\d{4}',text).group() 
        logger.info(f'Closed {file_path}')
    except AttributeError:
      print('File not redable')
      logger.info(f' file not readable')
      logger.info(f'Closed {file_path}')
    
    name_list.append(name_of_applicant)
    director_list.append(director_type)
    date_list.append(date)

df= pd.DataFrame({'Appointed/resigned/becoming': name_list,
'Director, whole-time director, full-time director, designated director, non-designated director':director_list,
'Date on Board resolution':date_list
})

df.to_excel('C:/Users/lenovo/Desktop/Boston consulting group/Board_resolution/Board resolution Output.xlsx',index=False)
