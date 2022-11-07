import win32com.client as win32
from tkinter import filedialog
import re

## CREATING EXCEL INSTANCE
excel = win32.Dispatch('Excel.Application')
excel.Visible = True
print('Choose Excel file...')
workbook_name = filedialog.askopenfilename()
workbook = excel.Workbooks.Open(Filename=workbook_name)

## CREATING POWERPOINT INSTANCE
ppt = win32.Dispatch('Powerpoint.Application')
ppt.Visible = True
presentation_name = filedialog.askopenfilename()
presentation = ppt.Presentations.Open(FileName=presentation_name)

## SELECT SHEETS AND EXTRACT SLIDE NUMBER FROM SHEET
# for sheet in workbook.Sheets:
#     if 'RD' not in sheet.Name:
#         print(sheet.Name)
#         name_split = sheet.Name.replace('-','').split()
#         for i in name_split:
#             if i.isdigit():
#                 slide_nmbr = int(i)
#                 break
#         print(sheet.Name, slide_nmbr)
#         del slide_nmbr

## INSERT ON SLIDE NUMBER
slides = presentation.Slides
i=0
for namedrange in workbook.Names:
    i += 1
    excel.Range(namedrange.Name).CopyPicture()
    shape = slides[i].Shapes.Paste()
    shape.left, shape.top = 10, 50







# for namedrange in workbook.Names:
#     print(namedrange.Name)
#     excel.Range(namedrange.Name).CopyPicture()
#     new_slide = presentation.Slides.Add(Index = 2, Layout = 12)
#     shape =new_slide.Shapes.Paste()
#     shape.left, shape.top = 10, 50
#     # shape = new_slide.Shapes.PasteSpecial(DataType = 10, Link = False)


## SAVE AND CLOSE
# presentation.SaveAs(filedialog.asksaveasfilename())
# excel.Quit()
# ppt.Quit()
