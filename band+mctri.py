import re
import pandas as pd
import os
import win32com.client as win32
from win32com.client import constants
import time
from colorset import colors

paths = pd.read_csv('filepaths.txt', index_col=False, header=None, delimiter='\s=|\s=\s|=\s|=', engine='python')
paths_list = list(paths[1])
file, file1, file2, text, text1, dir= paths_list

filename= file
filename1= file1
filename2= file2
text_file= text
text_file1= text1

directory = dir

path= os.path.join(directory, filename)
path1= os.path.join(directory, filename1)
path2 = os.path.join(directory, filename2)
text_path = os.path.join(directory, text_file)
text_path1= os.path.join(directory, text_file1)

df= pd.DataFrame()
print('creating dataframe.....\n')
df1= pd.DataFrame()


print('accessing to text files.....')
text = open(text_path, 'r', encoding='utf-8')
text1 = open(text_path1, 'r', encoding='utf-8')

print('reading text file.....\n')
contents= text.read()
contents1= text1.read()


print('compiling pattern to be matched.....\n')
p1 = re.compile(r'([A-Z]+-)(\d+)(-)(\d+)(\s+)(\w+)')  # MCTRI
p3= re.compile(r'(SCCONC)(\s+)([\d])')
p= re.compile(r'([A-Z]+-)(\d+)(-)(\d+)(\s+)(\w+)(\s+[A-Z]+)(\s+)(\w+)')   # BAND


print(colors.fg.red+'getting the mctri-cellname matches.....'+colors.disable)
m1 = p1.findall(contents)
print(m1)
print(colors.fg.red+'getting the mctri value matches.....'+colors.disable)
m3 = p3.findall(contents)
print(m3)
print(colors.fg.red+'getting the band matches.....'+colors.disable)
m= p.findall(contents1)
print(m)


rxotrx, tg, dash, trx, space, cell = map(list, zip(*m1))
scconc, space2, mctri = map(list, zip(*m3))
rxotx, tg2, dash2, trx2, space3, cell2, all, space4, band = map(list, zip(*m))

print(colors.fg.red+'assigning value to the dataframes....'+colors.disable)
time.sleep(0.5)
df['TG'] = tg
df['TRX Number'] = trx
df['CELL Name'] = cell
df['MCTRI'] = mctri
print(colors.fg.lightgreen+df+colors.disable)
time.sleep(0.5)
df1['TG'] = tg2
df1['TRX Number'] = trx2
df1['CELL Name'] = cell2
df1['BAND'] = band
print(colors.fg.lightgreen+df1+colors.disable)
time.sleep(0.5)



print(colors.fg.red+'\nexporting excel files of mctri & band.....')
df.to_excel(path, index=False)
df1.to_excel(path1, index=False)
print(colors.fg.red+'mctri & band excel created....')
print(colors.fg.green+'proceeding towards making the final combined excel.....')


print('pausing the script for 3 seconds.....')
time.sleep(3)
print('creating excel instance......\n')
excel = win32.gencache.EnsureDispatch('Excel.Application')


print('getting accessed to the excel files.....\n')
band_mctri_xl = excel.Workbooks.Open(path2)
mctri_xl = excel.Workbooks.Open(path)
band_xl = excel.Workbooks.Open(path1)

excel.Visible = True


print('activating mctri.xlsx.....\n')
mctri_xl.Activate()
mctri = mctri_xl.Worksheets("Sheet1")

cell1 = mctri.Range("A2:D2")
cell2 = cell1.End(Direction=constants.xlDown)
mctri.Range(cell1, cell2).Select()
print('copying mctri.xlsx values.....\n')
excel.Selection.Copy()


print('activating Band+mctri.xlsx.....\n')
band_mctri_xl.Activate()
mctri_final = band_mctri_xl.Worksheets("Sheet1")
print('pasting the mctri values.....\n')
mctri_final.Paste(mctri_final.Range("A2"))


print('activating band.xlsx.....\n')
band_xl.Activate()
band = band_xl.Worksheets("Sheet1")

cell3= band.Range("A2:D2")
cell4 = cell3.End(Direction=constants.xlDown)
band.Range(cell3, cell4).Select()
print('copying band values....\n')
excel.Selection.Copy()
print('again activating Band+mctri.xlsx.....\n')
band_mctri_xl.Activate()
band_final = band_mctri_xl.Worksheets("Sheet3")
print('pasting the band values.....\n')
band_final.Paste(band_final.Range("A2"))
print('your final combined excel is ready......')

