import re
import string
import xlwings as xw


#初始化一下
xw.App(visible=False,add_book=False)
xw.App.screen_updating=False


actived_book = xw.Book(r"C:\Users\HNCJ-liaobingzhi\Desktop\访问链接记录.xlsx")
pattern = r"源网络地址:*(.*?)\n"

actived_sheet = actived_book.sheets(1)


end_row = actived_sheet.used_range.last_cell.row
intend=int(end_row)
print(end_row)

for row in range(2,intend):
    range_data = actived_sheet.range(row,6).value
    if range_data == None:
        actived_sheet.range(row,7).value="空"
    else:
        jieguo = re.search(pattern,range_data)
        if jieguo == None:
            actived_sheet.range(row,7).value="空"
        else:
            actived_sheet.range(row,7).value = str(jieguo.group(1))
        jieguo =None


actived_book.save()