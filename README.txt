ooxml: Python interface for working with OOXML files such as xlsx, docx, pptx.

For now only reading xlsx works.

Here's how to use it to read xlsx file:

>>> from ooxml.spreadsheet import Spreadsheet
>>> workbook=Spreadsheet('book1.xlsx')             #read in book1.xlsx
>>> workbook.sheet_names 
['Sheet1', 'Sheet3', 'Sheet2']
>>> sheet1 = workbook.sheet(1)                     #index start from 1, not 0
>>> row1 = sheet1.row(1)                           #you can get a row
>>> row1
<spreadsheet.Workrow object at 0xb7a8898c>
>>> row1.cell(1).value                             #access a cell value in a row
'1'
>>> sheet1.cell(2,1).value                         #access by sheet.cell(x,y), x is row, y is column
'2'
>>> sheet2 = workbook.sheet('Sheet2')              #use name instead of index to get a sheet
>>> cell = sheet2.cell(5,5)
>>> sheet2.save_csv('my.csv')                      #save content to csv

Writing xlsx, reading/writing docx, pptx will be added later.
