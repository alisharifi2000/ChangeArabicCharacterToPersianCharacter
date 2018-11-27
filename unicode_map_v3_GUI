#update:
#add two options for users.
#change english number to persian number.
#change persian number to english number.

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import os
import tkinter.filedialog
from tkinter import *
import io
import xlsxwriter


global df
global df1
global persian_symbol
global arabic_symbol
global persian_unicode
global arabic_unicode
global persian_number
global english_number


def Change_Character(character):  # change arabic symbol to persian symbol .
    i = 0
    for items in arabic_symbol:
        if character == items:
            changed = persian_symbol[i]
            return changed
        i = i + 1
    return character

def Change_Numbers(character):  # change english number to persian number .
    i = 0
    for items in english_number:
        if character == items:
            changed = persian_number[i]
            return changed
        i = i + 1
    return character

def Change_Numbers1(character):  # change persian number to english number .
    i = 0
    for items in persian_number:
        if character == items:
            changed = english_number[i]
            return changed
        i = i + 1
    return character

def Seprate(word):  # separate word unchanged & put them in list of characters.
    list = []
    for characters in word:
        list.insert(len(list), characters)
    print(list)
    return list


def combine(list):  # combine changed charaters to make word.
    word = ''
    for i in range(0, len(list)):
        word = word + list[i]
    return word


def format_of_file(path1):  # find the input file format
    path = path1.split("/")
    n = len(path)
    format_file = path[n - 1].split(".")
    print("format : " + format_file[1])
    return (format_file[1])

def name_of_input_file(path1):#find name of input file
    path = path1.split("/")
    n = len(path)
    format_file = path[n - 1].split(".")
    print("format : " + format_file[0])
    return (format_file[0])

def selection():  # handle the order of radio_button that user choose.
    option = var.get()
    if (option == 1):
        option_action1()
    elif(option == 2):
        option_action2()
    elif(option == 3):
        option_action3()


def option_action1():#if user choice option "change arbic word & numbers to persian"
    global entry
    global root
    path = tkinter.filedialog.askopenfilename()
    if (os.path.exists(path)):
        file_format = format_of_file(path)
        if (file_format == "xlsx"):
            edit_excel_file(path)
        elif (file_format == "txt"):
            edit_text_file(path)
        else:
            print("فرمت فایل انتخابی اشتباه میباشد .")
            root1 = Tk()
            root1.title("خطا در فرمت فایل")
            label3 = Label(root1, text="خطا در انتخاب فرمت فایل ورودی  ").pack()
            label3 = Label(root1, text="برنامه فقط از فرمت های زیر پشتیبانی میکند:").pack()
            label3 = Label(root1, text=".xlsx").pack()
            label3 = Label(root1, text=".txt").pack()
            label3 = Label(root1, text="لطفا دوباره تلاش کنید.").pack()
            root1.mainloop()
        print("path : " + str(path))
    else:
        print("ادرس فایل وارد شده اشتباه میباشد . لطفا دوباره برنامه را راه اندازی کنید.")

def option_action2():#if user choice option "change english numbers to persian"
    global entry
    global root
    path = tkinter.filedialog.askopenfilename()
    if (os.path.exists(path)):
        file_format = format_of_file(path)
        if (file_format == "xlsx"):
            edit_excel_file1(path)
        elif (file_format == "txt"):
            edit_text_file1(path)
        else:
            print("فرمت فایل انتخابی اشتباه میباشد .")
            root1 = Tk()
            root1.title("خطا در فرمت فایل")
            label3 = Label(root1, text="خطا در انتخاب فرمت فایل ورودی  ").pack()
            label3 = Label(root1, text="برنامه فقط از فرمت های زیر پشتیبانی میکند:").pack()
            label3 = Label(root1, text=".xlsx").pack()
            label3 = Label(root1, text=".txt").pack()
            label3 = Label(root1, text="لطفا دوباره تلاش کنید.").pack()
            root1.mainloop()
        print("path : "+str(path))
    else:
        print("ادرس فایل وارد شده اشتباه میباشد . لطفا دوباره برنامه را راه اندازی کنید.")

def option_action3():  # if user choice option "change persian numbers to english"
        global entry
        global root
        path = tkinter.filedialog.askopenfilename()
        if (os.path.exists(path)):
            file_format = format_of_file(path)
            if (file_format == "xlsx"):
                edit_excel_file2(path)
            elif (file_format == "txt"):
                edit_text_file2(path)
            else:
                print("فرمت فایل انتخابی اشتباه میباشد .")
                root1 = Tk()
                root1.title("خطا در فرمت فایل")
                label3 = Label(root1, text="خطا در انتخاب فرمت فایل ورودی  ").pack()
                label3 = Label(root1,text="برنامه فقط از فرمت های زیر پشتیبانی میکند:").pack()
                label3 = Label(root1, text=".xlsx").pack()
                label3 = Label(root1, text=".txt").pack()
                label3 = Label(root1, text="لطفا دوباره تلاش کنید.").pack()
                root1.mainloop()

            print("path : " + str(path))
        else:
            print("ادرس فایل وارد شده اشتباه میباشد . لطفا دوباره برنامه را راه اندازی کنید.")

def edit_excel_file(path):#if file format xlsx & user choose option 1
    name = name_of_input_file(path)
    open_file = pd.read_excel(path, sheetname='Sheet1')
    workbook = xlsxwriter.Workbook(name + '_edit.xlsx')
    worksheet = workbook.add_worksheet()
    size = open_file.shape
    nrow = size[0]
    ncol = size[1]
    for i in range(0, nrow):
        for j in range(0, ncol):
            word = str(open_file.iloc[i, j])
            if not (word is None):
                list = Seprate(word)
                list1 = []
                for items in list:
                    c = Change_Character(items)
                    list1.insert(len(list1), c)
                change_word = combine(list1)
                print(change_word)
                worksheet.write(i, j, change_word)
    workbook.close()

def edit_excel_file1(path):#if file format xlsx & user choose option 2
    name = name_of_input_file(path)
    open_file = pd.read_excel(path, sheetname='Sheet1')
    workbook = xlsxwriter.Workbook(name + '_edit.xlsx')
    worksheet = workbook.add_worksheet()
    size = open_file.shape
    nrow = size[0]
    ncol = size[1]
    for i in range(0, nrow):
        for j in range(0, ncol):
            word = str(open_file.iloc[i, j])
            if not (word is None):
                list = Seprate(word)
                list1 = []
                for items in list:
                    c = Change_Numbers(items)
                    list1.insert(len(list1), c)
                change_word = combine(list1)
                print(change_word)
                worksheet.write(i, j, change_word)
    workbook.close()

def edit_excel_file2(path):#if file format xlsx & user choose option 3
    name = name_of_input_file(path)
    open_file = pd.read_excel(path, sheetname='Sheet1')
    workbook = xlsxwriter.Workbook(name + '_edit.xlsx')
    worksheet = workbook.add_worksheet()
    size = open_file.shape
    nrow = size[0]
    ncol = size[1]
    for i in range(0, nrow):
        for j in range(0, ncol):
            word = str(open_file.iloc[i, j])
            if not (word is None):
                list = Seprate(word)
                list1 = []
                for items in list:
                    c = Change_Numbers1(items)
                    list1.insert(len(list1), c)
                change_word = combine(list1)
                print(change_word)
                worksheet.write(i, j, change_word)
    workbook.close()

def edit_text_file(path):#if file format txt & user choose option 1
    read_file = io.open(path, 'r', encoding='utf8')
    name = name_of_input_file(path)
    write_file = io.open(name + '_edit.txt', 'w', encoding='utf8')
    for line in read_file:
        for words in line:
            word = str(words)
            list = Seprate(word)
            list1 = []
            for items in list:
                c = Change_Character(items)
                list1.insert(len(list1), c)
            change_word = combine(list1)
            write_file.write(change_word)
            write_file.flush()
            print(change_word)
        write_file.write("\n")

def edit_text_file1(path):#if file format txt & user choose option 2
    read_file = io.open(path, 'r', encoding='utf8')
    name = name_of_input_file(path)
    write_file = io.open(name + '_edit.txt', 'w', encoding='utf8')
    for line in read_file:
        for words in line:
            word = str(words)
            list = Seprate(word)
            list1 = []
            for items in list:
                c = Change_Numbers(items)
                list1.insert(len(list1), c)
            change_word = combine(list1)
            write_file.write(change_word)
            write_file.flush()
            print(change_word)
        write_file.write("\n")

def edit_text_file2(path):#if file format txt & user choose option 2
    read_file = io.open(path, 'r', encoding='utf8')
    name = name_of_input_file(path)
    write_file = io.open(name + '_edit.txt', 'w', encoding='utf8')
    for line in read_file:
        for words in line:
            word = str(words)
            list = Seprate(word)
            list1 = []
            for items in list:
                c = Change_Numbers1(items)
                list1.insert(len(list1), c)
            change_word = combine(list1)
            write_file.write(change_word)
            write_file.flush()
            print(change_word)
        write_file.write("\n")


df = pd.read_excel('maping.xlsx', sheetname='Sheet1')
df1 = pd.read_excel('maping2.xlsx',sheetname='Sheet1')

arabic_unicode = df['Unicode_code_arabic']
persian_unicode = df['Unicode_code_farsi']
arabic_symbol = df['arabic_symbol']
persian_symbol = df['persian_symbol']
english_number = df1['english_number']
persian_number = df1['persian_number']

root = Tk()
root.title("برنامه ویرایش فایل های ورودی")
label3 = Label(root, text="خوش آمدید به برنامه ویرایش فایل های ورودی  ").pack()
label2 = Label(root,
               text="-------------------------------------------------------------------------------------------------------------- ").pack()
label3 = Label(root, text="راهنما ").pack()
label2 = Label(root,
               text="-------------------------------------------------------------------------------------------------------------- ").pack()
label3 = Label(root, text="برای انجام عملیات های خود در این برنامه کافی است ").pack()
label4 = Label(root, text="یکی از گزینه های موجود را پر کنید").pack()
label3 = Label(root, text="سپس آدرس فولدر مورد نظر خود را در صفحه ی باز شده انتخاب نمایید .").pack()
label3 = Label(root, text="پس از انجام عملیات فایل جدیدی در همان آدرس ، با نام  .").pack()
label3 = Label(root, text=" اسم فایل + edit  ").pack()
label3 = Label(root, text=" ساخته میشود .").pack()
label3 = Label(root, text="****************************************************************").pack()
label3 = Label(root, text=":عملیات های قابل اجرا با این برنامه عبارتند از ").pack()
label3 = Label(root, text=" تبدیل کاراکتر ها و اعداد عربی به کاراکتر ها و اعداد فارسی ").pack()
label3 = Label(root, text=" نبدیل اعداد انگلیسی به اعداد فارسی ").pack()
label3 = Label(root, text="").pack()
label3 = Label(root, text="فرمت های پشتیبانی شده در این برنامه عبارتند از  ").pack()
label3 = Label(root, text=".xlsx").pack()
label3 = Label(root, text=".txt").pack()
label2 = Label(root,
               text="-------------------------------------------------------------------------------------------------------------- ").pack()
var = IntVar()
Radiobutton(root, text="تبدیل کاراکتر ها و اعداد عربی به فارسی", variable=var, value=1, command=selection).pack()
Radiobutton(root, text="تبدیل اعداد انگلیسی به اعداد فارسی", variable=var, value=2, command=selection).pack()
Radiobutton(root, text="تبدیل اعداد فارسی به اعداد انگلیسی", variable=var, value=3, command=selection).pack()
label = Label(root)
label.pack()

root.mainloop()
