import tabula
import pandas as pd
import os
import glob
import csv
from xlsxwriter.workbook import Workbook

df = tabula.read_pdf(r"C:/Users/balay/OneDrive/Masaüstü/OCR/tasacion 040.pdf", 
                     pages='all')
tabula.convert_into(r"C:/Users/balay/OneDrive/Masaüstü/OCR/tasacion 040.pdf", 
                    r"C:/Users/balay/OneDrive/Masaüstü/OCR/tasacion 040.csv" , 
                    output_format="csv",pages='all', stream=True)





for csvfile in glob.glob(os.path.join('.', '*.csv')):
    workbook = Workbook(csvfile[:-4] + '.xlsx')
    worksheet = workbook.add_worksheet()
    with open(csvfile, 'rt', encoding='cp1252') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)  
    workbook.close()
