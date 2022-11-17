import win32com.client as win32
from tkinter import filedialog

## CONFIGURATION
def_rng_pos = {'left': 15, 'width': 320, 'top': 100}
def_chrt_pos = {'left': 350, 'width': 320, 'top': 120}
rng_positions = {
    'bigdertien': {'left':15, 'width': 670, 'top': 130},
    'smalldertien': {'left':220, 'width': 270, 'top': 75}
}
chrt_positions = {
    'Decompositie': {'left':15, 'width': 670, 'top': 130},
}

## CREAT EXCEL INSTANCE
excel = win32.Dispatch('Excel.Application')
excel.Visible = True
print('Choose Excel file...')
workbook_name = filedialog.askopenfilename()
workbook = excel.Workbooks.Open(Filename=f'{workbook_name}')

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
    shape.left, shape.top, shape.width = rng_positions.get(rangename, def_rng_pos)['left'], rng_positions.get(rangename, def_rng_pos)['top'], rng_positions.get(rangename, def_rng_pos)['width']
    del slide_nmbr

## LOOP THROUGH SHEETS, FIND GRAPHS AND INSERT IN PRESENTATION
for sheet in workbook.Sheets:
    if 'RD' not in sheet.Name:
        name_split = sheet.Name.replace('-','').split()
        for i in name_split:
            if i.isdigit():
                slide_nmbr = int(i)
                break
        charts = sheet.ChartObjects()
        for chart in charts:
            chart.CopyPicture()
            shape = slides[slide_nmbr-1].Shapes.Paste()
            shape.left, shape.top, shape.width = chrt_positions.get(chart.Name, def_chrt_pos)['left'], chrt_positions.get(chart.Name, def_chrt_pos)['top'], chrt_positions.get(chart.Name, def_chrt_pos)['width']
        del slide_nmbr

## SAVE AND CLOSE
presentation.SaveAs(filedialog.asksaveasfilename())
excel.Quit()
# ppt.Quit()
