# importing required modules
import glob
import os
from datetime import datetime
from pathlib import Path
import PyPDF2
import re
import pandas as pd
import glob
import uuid
from pathlib import Path
from pdf2image import convert_from_path
import pytesseract
import easygui
import time
from rich.progress import track
import warnings
warnings.filterwarnings('ignore')
#Keep your path here

pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'

poppler_path = 'C:/poppler-0.68.0/bin'

def pdf_extraction(text):
    global leave,month,year

    date = re.findall(r"[\d]{1,4}[/-][\d]{1,4}[/-][\d]{1,4}.+", text)
    # print(date)
    date = str(date).replace('lea ve', 'leave').replace("Lea ve","leave").replace('leav','leave').replace("LEA VE","leave").replace('le ave','leave').replace("Leave",'leave').replace('LEAVE','leave')
    # print(date)
    leave = re.findall('leave.', date)
    if leave:
        leave = len(leave)
        # print(leave)
    else:
        leave = re.findall('leave', text)
        if leave:
            leave = len(leave)
        else :
            leave = 0
            # print(leave)

    pattern = r'(aug|jan|May|Jan|Nov|Aug|Jun|Sep|Oct|Feb|Dec|Jul|Apr|Mar|July|)'
    file = str(Path(filename).stem)
    month = re.findall(pattern, file)
    if month:
        month = str(month)
        month = month.replace('[', '').replace("]", "").replace("'", "").replace(",", "").rstrip().lstrip()
        month = month.replace("aug", "August").replace("jan", "January").replace("Jan", "January").replace("Nov", "November") \
            .replace("Aug", "August").replace("Jun", "June").replace("Sep", "September").replace("Oct", "October").replace("Feb", "February").replace(
            "Dec", "December").replace('Jul',"July").replace("Apr","April").replace("Mar",'March').replace('Augustust','August').replace('Januaryuary','January').rstrip().lstrip()

        # print(month)


    year = re.findall(r'[\d]{1,4}',file)
    if year:
        year = str(year).replace('[','').replace("]","").replace("'","")
        # print(year)



def ocr_extract():
    # print("OCR ")
    #change this also
    image_path = f"{directory2}/Dump_images"
    if not os.path.exists(image_path):
        os.makedirs(image_path)

    images = convert_from_path(filename, poppler_path=poppler_path, output_file=str(uuid.uuid4()))
    # print(len(images))

    for i in range(len(images)):
        images[i].save(f'{image_path}/img{i}.jpg')

    imgfile_list = glob.glob(image_path + "/*")
    # print("Images",len(imgfile_list))
    if len(imgfile_list) > 1:
        count = 1
        for imgfile in imgfile_list:
            text = pytesseract.image_to_string(imgfile, lang='eng')
            # print(text)
            pdf_extraction(text)
            data = {'Year': [year],
                    'Month': [month],
                    'Leave Count': [leave]}
            df = pd.DataFrame(data)
            # print(df)
            df.to_excel(
                f'{directory2}/{Path(Dump_Excel).stem}/{Path(filename).stem}{count}.xlsx', index=False)
            count+=1

        path = f"{directory2}/{Path(Dump_Excel).stem}"
        file_list = glob.glob(path + "/*.xlsx")
        excl_list = []
        for file in file_list:
            excl_list.append(pd.read_excel(file))
        new_df = pd.DataFrame()
        excl_merged = pd.concat(excl_list, ignore_index=True)
        # print("MERGED EXCEL ")
        # print(excl_merged)
        excel_name = str(Path(filename).stem).split("_")[1:]
        if len(excel_name) <= 2:
                excel_name[0] = month 
        else:
            excel_name.pop(0) 
            excel_name[0]= month
        excel_name = " ".join(excel_name)
        excl_merged.to_excel(f'{directory2}/{Path(Dump_PDF).stem}/{excel_name}.xlsx', index=False,
                             engine='xlsxwriter')
        # print("___________________________________________________________________")
        for file in file_list:
            os.remove(file)
    else:
        for imgfile in imgfile_list:
            text = pytesseract.image_to_string(imgfile, lang='eng')
            # print(text)
            pdf_extraction(text)
            data = {'Year': [year],
                    'Month': [month],
                    'Leave Count': [leave]}
            df = pd.DataFrame(data)
            # print(df)
            excel_name = str(Path(filename).stem).split("_")[1:]
            if len(excel_name) <= 2:
                    excel_name[0] = month 
            else:
                excel_name.pop(0) 
                excel_name[0]= month
            excel_name = " ".join(excel_name)
            df.to_excel(f'{directory2}/{Path(Dump_PDF).stem}/{excel_name}.xlsx', index=False, engine='xlsxwriter')
            # print("___________________________________________________________________")
    for file in imgfile_list:
        os.remove(file)

#change this path till RP1
directory1 = easygui.diropenbox()

for directory2 in (glob.iglob(f'{directory1}/*')):
    pdf_count = [file for file in glob.iglob(f'{directory2}/*.pdf')]
    description = Path(directory2).stem +" "+"PDF Processing..."
    for i,filename in zip(track(range(len(pdf_count)), description=description),glob.iglob(f'{directory2}/*.pdf')):

        pdfFileObj = open(filename, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        pages = pdfReader.numPages
        # print("pages",pages)
        Dump_Excel = f'{directory2}/Dump_Excel'
        if not os.path.exists(Dump_Excel):
            os.makedirs(Dump_Excel)
        Dump_PDF = f'{directory2}/Dump_PDF'
        if not os.path.exists(Dump_PDF):
            os.makedirs(Dump_PDF)
        pageObj = pdfReader.getPage(0)
        pgtext = pageObj.extractText()
        if pages > 1 and len(pgtext)>500:
            for i in range(pages):
                pageObj = pdfReader.getPage(i)
                text = pageObj.extractText()
                # print(text)
                # print(len(text))
                if len(text)>500 or len(pgtext)>500:
                    pdf_extraction(text)
                    data = {'Year': [year],
                            'Month': [month],
                            'Leave Count': [leave]}
                    df = pd.DataFrame(data)
                    # print(df)
                    df.to_excel(
                        f'{directory2}/{Path(Dump_Excel).stem}/{Path(filename).stem}{i}.xlsx', index=False)
                else:
                    ocr_extract()

            path = f"{directory2}/{Path(Dump_Excel).stem}"
            file_list = glob.glob(path + "/*.xlsx")
            excl_list = []
            for file in file_list:
                excl_list.append(pd.read_excel(file))
            excl_merged = pd.concat(excl_list, ignore_index=True)
            excel_name = str(Path(filename).stem).split("_")[1:]
            if len(excel_name) <= 2:
                    excel_name[0] = month 
            else:
                excel_name.pop(0) 
                excel_name[0]= month
            excel_name = " ".join(excel_name)
            excl_merged.to_excel(f'{directory2}/{Path(Dump_PDF).stem}/{excel_name}.xlsx', index=False, engine='xlsxwriter')
            # print("MERGED EXCEL ")
            # print(excl_merged)
            # print("___________________________________________________________________")

            for file in file_list:
                os.remove(file)
                #
        else:
            pageObj = pdfReader.getPage(0)
            text = pageObj.extractText()
            # print(text)
            # print(len(text))
            if len(text)>500:
                pdf_extraction(text)
                data = {'Year': [year],
                        'Month': [month],
                        'Leave Count': [leave]}
                df = pd.DataFrame(data)
                # print(df)
                excel_name = str(Path(filename).stem).split("_")[1:]
                if len(excel_name) <= 2:
                    excel_name[0] = month 
                else:
                    excel_name.pop(0) 
                    excel_name[0]= month

                excel_name = " ".join(excel_name)
                df.to_excel(f'{directory2}/{Path(Dump_PDF).stem}/{excel_name}.xlsx', index=False, engine='xlsxwriter')
                # print("___________________________________________________________________")

            else:
                ocr_extract()

for directory2 in glob.iglob(f'{directory1}/*'):
    file_list = list()
    for filename in glob.iglob(f'{directory2}/Dump_PDF/*'):
       
        file_list.append(Path(filename).stem)
        
    sorted_dates = sorted(file_list, key = lambda x: datetime.strptime(x, '%B %Y'))
    excl_list = []
    for dates in sorted_dates:
        excl_list.append(pd.read_excel(f"{directory2}/Dump_PDF/{dates}.xlsx"))

    total_counts = sum([count[2] for counts in excl_list for count in counts.values if count[2]!= 0])
    
    data = {'Year': [''],
            'Month': ['Total'],
            'Leave Count': [total_counts]}
    df = pd.DataFrame(data)
    excl_list.append(df)
    excl_merged = pd.concat(excl_list, ignore_index=True)
    
    excl_merged.to_excel(f'{directory1}/{Path(directory2).stem}_Attendance_Report.xlsx', index=False, engine='xlsxwriter')