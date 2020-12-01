import tabula
import pandas as pd
import os
import glob
import csv
from xlsxwriter.workbook import Workbook

df = tabula.read_pdf(r"LOCATION OF PDF FILE FOR READING", 
                     pages='all')
tabula.convert_into(r"LOCATION OF PDF FILE FOR CONVERTING", 
                    r"LOCATION OF CSV FILE FOR SAVING" , 
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
