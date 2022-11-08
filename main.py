import win32com.client as win32
from tkinter import filedialog

## CREAT EXCEL INSTANCE
excel = win32.Dispatch('Excel.Application')
excel.Visible = True
print('Choose Excel file...')
workbook_name = filedialog.askopenfilename()
workbook = excel.Workbooks.Open(Filename=f'"{workbook_name}"')

## CREAT POWERPOINT INSTANCE
ppt = win32.Dispatch('Powerpoint.Application')
ppt.Visible = True
presentation_name = filedialog.askopenfilename()
presentation = ppt.Presentations.Open(FileName=f'"{presentation_name}"')
slides = presentation.Slides

## LOOP THROUGH NAMED RANGES AND INSERT IN PRESENTATION
for namedrange in workbook.Names:
    rangename = namedrange.Name
    rangeref = namedrange.Value
    sheetname = rangeref.replace('=','').split('!')[0]
    if 'RD' not in sheetname:
        name_split = sheetname.replace('-','').split()
        for i in name_split:
            if i.isdigit():
                slide_nmbr = int(i)
                break
    excel.Range(rangename).CopyPicture()
    shape = slides[slide_nmbr-1].Shapes.Paste()
    shape.left, shape.top = 25, 160
    del slide_nmbr

## LOOP THROUGH SHEETS, FIND GRAPHS AND INSERT IN PRESENTATION
for sheet in workbook.Sheets:
    if 'RD' not in sheet.Name:
        print(sheet.Name)
        name_split = sheet.Name.replace('-','').split()
        for i in name_split:
            if i.isdigit():
                slide_nmbr = int(i)
                break
        charts = sheet.ChartObjects()
        for chart in charts:
            chart.CopyPicture()
            shape = slides[slide_nmbr-1].Shapes.Paste()
            shape.left, shape.top = 25, 160
        del slide_nmbr








## INSERT ON SLIDE NUMBER
# slides = presentation.Slides
# i=0
# for namedrange in workbook.Names:
#     i += 1
#     excel.Range(namedrange.Name).CopyPicture()
#     shape = slides[i].Shapes.Paste()
#     shape.left, shape.top = 10, 50



# for sheet in workbook.Sheets:
#     print(sheet.Name)
#     for rng in sheet.Names:
#         print(rng.Name)


# new_slide = presentation.Slides.Add(Index = 2, Layout = 12)


## SAVE AND CLOSE
# presentation.SaveAs(filedialog.asksaveasfilename())
# excel.Quit()
# ppt.Quit()
