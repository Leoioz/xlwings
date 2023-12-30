import pyautogui as py
import xlwings as xw
import pandas as pd
import time,os,sys

#初始化一下
xw.App(visible=False,add_book=False)
xw.App.screen_updating=False
#然后打开指定的app
book_read_path = py.prompt(title='文件路径',text='选择要读取数据的文件路径',default='')
print("读取数据的文件路径app:",book_read_path)
book_write_path = py.prompt(title='文件路径',text='选择要保存数据的文件路径',default='')
print("保存数据的文件路径app:",book_write_path)
actived_read_book = xw.Book(book_read_path)
actived_write_book = xw.Book(book_write_path)
#顺便查一下当前打开的app的pid
app_key = xw.apps.keys()
print("查一下当前打开的app的pid:",app_key)
#查一下打开的这个app里面有多少个book
books = xw.books
print("查一下打开的这个app里面有多少个book:",books)
#查一下打开的这个app里面激活的book
actived_sheet = xw.sheets.active
print("查一下打开的这个app里面激活的book:",actived_sheet)
#指定要操作的book
chese_book=py.prompt(title='原表要操作那个工作簿',text='原表要操作那个工作簿',default='')
chese_book=int(chese_book)
actived_read_sheet = actived_read_book.sheets(chese_book)
actived_write_sheet = actived_write_book.sheets(actived_write_book.sheets.add())
#选中需要的格子，并读取数据
#给定起始单元格，计算行列
actived_range_start_row = py.prompt(title='读取文件起始单元格',text='你的第一个单元格,例如F2',default='')
actived_range_column = actived_read_sheet.range(actived_range_start_row).column
actived_range_start_row1 = actived_read_sheet.range(actived_range_start_row).row

#给定结束单元格，计算行
actived_range_end_row = py.prompt(title='读取文件结束单元格',text='你的最后一个单元格,例如F999',default='')
actived_range_end_row1 = actived_read_sheet.range(actived_range_end_row).row

actived_range_count_row = actived_range_end_row1 - actived_range_start_row1
print("起始单元格行号：",actived_range_start_row1)
print("起始单元格列号：",actived_range_column)
print("结束单元格行号：",actived_range_end_row1)
print("总差：",actived_range_count_row)

#range_value = actived_read_sheet.range(actived_range_start_row).expand(mode='down')
range_value = xw.Range(actived_range_start_row).expand(mode='down')

#先拿到总行总列
rows = py.prompt(title='要读多少行',text='要读多少行',default='')
rows = int(rows)
cols = py.prompt(title='要读多少列',text='要读多少列',default='')
cols = int(cols)
special_cols = py.prompt(title='需要指定读第几列吗？',text='需要指定读第几列吗？',default='0')
special_cols = int(special_cols)
write_row = py.prompt(title='需要指定写第几行吗？',text='需要指定写第几行吗？',default='0')
write_row = int(write_row)

try:
    for row in range(1,rows):
        devices = {}
        if special_cols==0:
            actived_read_sheet[row,cols].value
        else:
            devices=actived_read_sheet[row,special_cols].value
            den_devices=actived_read_sheet[row,special_cols-3].value
            range_value=devices.split(sep=",")
            den_range_value=den_devices
            print(range_value)
            range_value_count = len(range_value)
            range_value_count = int(range_value_count)
            if range_value_count==0:
                print("单元格内容缺失")
            elif range_value_count==1:
                actived_write_sheet.range(write_row,2).options(transpose=True).value=range_value[0]
                write_row=write_row+1
            elif range_value_count>1:
                for step in range(0,range_value_count):
                    actived_write_sheet.range(write_row,2).options(transpose=True).value=range_value[step]
                    actived_write_sheet.range(write_row,1).options(transpose=True).value=den_range_value
                    write_row=write_row+1
except:
    pa=1


actived_read_book.save()
actived_read_book.close()
actived_write_book.save()
actived_write_book.close()
xw.App.quit()