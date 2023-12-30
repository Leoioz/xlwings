#读取txt文本里面的内容# 打开文件
import os,sys
import xlwings as xw
import pyautogui as py

def rw_excel():
    #初始化一下
    xw.App(visible=False,add_book=False)
    xw.App.screen_updating=False

    #提示录入需要访问包含txt路径的excel文件，并直接激活该excel
    excel_read_path = py.prompt(title='文件路径',text='需要访问包含txt路径的excel文件',default=r'C:\Users\HNCJ-liaobingzhi\Desktop\1.xlsx')
    print("读取数据的文件路径excel:",excel_read_path)
    actived_excel = xw.Book(excel_read_path)

    #查询当前打开的excel的pid
    excel_key = xw.apps.keys()
    print("当前打开的excel的pid:",excel_key)

    #查一下打开的这个excel里面有多少个sheet
    sheets = xw.sheets
    print("打开的这个excel里面有多少个sheet:",sheets)

    #提示要激活该excel的那个sheet
    actived_sheet=py.prompt(title='激活工作簿？',text='原表要激活那个工作簿？',default='sheet1')
    actived_sheet1 = actived_excel.sheets(actived_sheet)

    #查一下打开的这个excel里面激活的sheet
    actived_sheet = xw.sheets.active
    print("打开的这个excel里面激活的sheet:",actived_sheet)


    actived_sheet_data=actived_sheet1.range('A1').expand(mode='down').value

    active_sheet2 = actived_excel.sheets.add()

    write_row =1

#C:\Users\HNCJ-liaobingzhi\Desktop\1.xlsx
    for path_data in actived_sheet_data:
        path_data_spl = [path_data.split('//')]
        file_name = [name[5] for name in path_data_spl]
        with open(path_data, 'r') as file:
        # 打开文件并读取所有行
            txt_data = file.readlines()
            # 将每一行去掉换行符，并转换为列表
            txt_data = [line.strip() for line in txt_data]
            # 打印结果
            print(txt_data)
            for td in txt_data:
                active_sheet2.range(write_row,2).options(transpose=True).value = td
                active_sheet2.range(write_row,1).options(transpose=True).value = file_name
                write_row = write_row + 1

    pass

def read_txt(txt_file):
    """
    Reads the contents of a text file and prints it to the console.

    Parameters:
        txt_file (str): The path to the text file.
        mode (str): The mode in which to open the file, e.g. 'r' for reading.
        codeing (str): The encoding of the text file.

    Returns:
        None

    """
    with open(txt_file, 'r') as file:
        # 打开文件并读取所有行
        txt_data = file.readlines()
        # 将每一行去掉换行符，并转换为列表
        txt_data = [line.strip() for line in txt_data]
        # 打印结果
        return txt_data


if __name__ == '__main__':
    rw_excel()
    
