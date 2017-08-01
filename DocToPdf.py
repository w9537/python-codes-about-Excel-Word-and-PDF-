import sys
import os
import comtypes.client
import tkinter.filedialog
from os.path import splitext
from tkinter import *
import tkinter.messagebox

root = Tk()

root.title("Button Test")


wdFormatPDF = 17


def callback1():
    global fpath
    fpath = tkinter.filedialog.askdirectory()

def callback2():
    global fpath2
    fpath2 = tkinter.filedialog.askdirectory()

def callback3():
    for root, dirs, files in os.walk(fpath):
        for file in files:
            if file.endswith('.doc') or file.endswith('.docx'):
                file1 = splitext(file)
                out_file_path = fpath2 + '/' + file1[0]
                filename = os.path.join(root, file)
                in_file = os.path.abspath(filename)
                out_file = os.path.abspath(out_file_path)
                word = comtypes.client.CreateObject('Word.Application')
                doc = word.Documents.Open(in_file)
                doc.SaveAs(out_file, FileFormat=wdFormatPDF)
                doc.Close()
                word.Quit()


Button(root, text="word所在文件夹", fg="blue",bd=2,width=28,command=callback1).pack()

Button(root, text="pdf输出文件夹", fg="blue",bd=2,width=28,command=callback2).pack()

Button(root, text="转换", fg="blue",bd=2,width=28,command=callback3).pack()

#Button(root, text="停止", fg="blue",bd=2,width=28,command=callback4).pack()

root.mainloop()

#root.mainloop()