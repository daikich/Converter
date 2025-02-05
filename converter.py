#import
import tkinter as tk
import subprocess
import fitz
import time
path = r"output"
from docx2pdf import convert
from pdf2docx import Converter
from tkinter import filedialog

#window
root = tk.Tk()
root.title("変換")
root.geometry("350x225")

#変数
content="output"

#def
def file_p():
    global file_path
    root1 = tk.Tk()
    root1.withdraw()
    file_path = filedialog.askopenfilename(title="ファイルを選択してください")
    if file_path:
        print(file_path)
        label5.config(text="PATH:"+file_path)
        label_info.config(text="info:変換するファイルを選択しました")
    else:
        print("パス指定失敗")
        label5.config(text="PATH:ファイルを選択してください")
        label_info.config(text="info:パス指定失敗")

def create_rectangle(canvas, x1, y1, x2, y2):
    canvas.create_rectangle(x1, y1, x2, y2, outline="black", width=1)

def word2pdf():
    label_info.config(text="wordをpdfに変換しています")
    convert(file_path, "output/"+content+".pdf")
    label_info.config(text="変換できました")
    op()
    quit_me(root)
    
def pdf2word():
    label_info.config(text="pdfをwordに変換しています")
    cv = Converter(file_path)
    label_info.config(text="変換できました")
    cv.convert("output/"+content+".docx", start=0, end=None)
    cv.close()
    op()
    quit_me(root)
    
def pdf2txt():
    label_info.config(text="pdfをtxtに変換しています")
    txt_path="output/"+content+".txt"
    pdf_path=file_path
    pdf_document = fitz.open(pdf_path)
    with open(txt_path, 'w', encoding='utf-8') as text_file:
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text = page.get_text()
            text_file.write(text)
            text_file.write("\n")
    label_info.config(text="変換できました")
    op()
    quit_me(root)

def quit_me(root_window):
        root_window.quit()
        root_window.destroy()

def get_txt():
    global content 
    content = t_entry.get()
    print("ファイル名:"+content)
    label_info.config(text="info:出力ファイル名が変更されました")

def op():
    subprocess.Popen(['explorer', path], shell=True)

#canvas
canvas = tk.Canvas(root, width=400, height=300)
canvas.pack()
create_rectangle(canvas, 15, 180, 335, 128)
#label1
label1 = tk.Label(text="選択してください")
label1.place(x=140,y=119)
#label2
label2 = tk.Label(text="現在の出力ファイル名")
label2.place(x=10,y=55)
#label3
label3 = tk.Label(text=".(拡張子)")
label3.place(x=210,y=80)
#label4
label4 = tk.Label(text="変換するファイル")
label4.place(x=10,y=10)
#label5
label5 = tk.Label(text="PATH:")
label5.place(x=20,y=30)
#label_info
label_info = tk.Label(text="info:")
label_info.place(x=10,y=190)
#text_box
t_entry = tk.Entry(root, width=30)
t_entry.insert(tk.END,"output")
t_entry.place(x=20,y=80)
#button1
Button1 = tk.Button(root,text="word→pdf",command=word2pdf)
Button1.place(x=25,y=142)
#button2
Button2 = tk.Button(root,text="pdf→word",command=pdf2word)
Button2.place(x=145,y=142)
#button3
Button3 = tk.Button(root,text="pdf→txt",command=pdf2txt)
Button3.place(x=270,y=142)
#button4
Button4 = tk.Button(root, text="ファイル名変更", command=get_txt)
Button4.place(x=265,y=74)
#button5
Button5 = tk.Button(root, text="参照", command=file_p)
Button5.place(x=280,y=20)

#rootmainloop
root.mainloop()