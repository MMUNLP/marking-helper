import tkinter as tk
from tkinter import *
from glob import glob
import os,sys,io
import pathlib
from tkinter import messagebox,filedialog
from prepare_upload import generate_feedbacks_for_upload,check_exist,match_name_list
from threading import Thread
from tkinter import ttk


buffer = io.StringIO()
sys.stdout = sys.stderr = buffer

"""
Variables
"""
GRADE_FILE=''
FEEDBACK_FOLDER_PATH=''
PDF_FOLDER_PATH=''


"""
GUI
"""
root = Tk()
root.geometry("400x300")
root.title('Moodle Helper ver. 1.1')

"""
Select Grading Sheet
"""

frame = Frame(root)
my_label0 = Label(frame, text='Select Grading Sheet (csv)')
my_label0.pack(side="top") 

e1 = Entry(frame,width=50)
e1.pack(side="left",fill=Y,expand=1) 

def select_grade_sheet():
    e1.delete(0, tk.END)
    GRADE_FILE = filedialog.askopenfilename(filetypes=[("csv file","*.csv")])
    if '.csv' not in GRADE_FILE:
        return messagebox.showerror(title='Error', message='Invalid CSV File. Please reselect and submit!')
    e1.insert(END, GRADE_FILE) # add this
    

my_button=Button(frame,text="Browse",command=select_grade_sheet)
my_button.pack(side="left",fill=Y,expand=1) 
frame.pack()

"""
Input Path box
"""
frame = Frame(root)
my_label0 = Label(frame, text='Select Feedback Folder (docx or pdf)')
my_label0.pack(side="top") 

e2 = Entry(frame,width=50)
e2.pack(side="left",fill=Y,expand=1) 

def select_feedback_folder():
    e2.delete(0, tk.END)
    FEEDBACK_FOLDER_PATH = filedialog.askdirectory()+'/'
    e2.insert(END, FEEDBACK_FOLDER_PATH) # add this

my_button=Button(frame,text="Browse",command=select_feedback_folder)
my_button.pack(side="left",fill=Y,expand=1) 
frame.pack()

"""
Output Path box
"""
frame = Frame(root)
my_label0 = Label(frame, text='Output path')
my_label0.pack(side="top") 

e3 = Entry(frame,width=50)
e3.pack(side="left",fill=Y,expand=1) 

def select_output_folder():
    e3.delete(0, tk.END)
    PDF_FOLDER_PATH = filedialog.askdirectory()+'/'
    e3.insert(END, PDF_FOLDER_PATH) # add this

my_button=Button(frame,text="Browse",command=select_output_folder)
my_button.pack(side="left",fill=Y,expand=1) 
frame.pack()


"""
Output Path box
"""
frame = Frame(root)
my_button2=Button(frame,text="Zip your feedbacks", bg = 'red',fg='white',padx=10,pady=10)

def convert_files():
    GRADE_FILE = e1.get()
    FEEDBACK_FOLDER_PATH = e2.get()
    PDF_FOLDER_PATH = e3.get()
    # empty or invalid path error handling
    if any([check_exist(GRADE_FILE),check_exist(FEEDBACK_FOLDER_PATH),
        check_exist(PDF_FOLDER_PATH)])==False:
        return messagebox.showerror(title='Error', message='Invalid Directory/File. Please reselect and submit!')

    
    if len(glob(FEEDBACK_FOLDER_PATH+'/'+f'*.docx'))==0 and len(glob(FEEDBACK_FOLDER_PATH+'/'+f'*.pdf'))==0:
        return messagebox.showerror(title='Error', message='No pdf or docx in your feedback directory. Please reselect and submit!')
    
    if match_name_list(csv_file=GRADE_FILE,input_feedback_path=FEEDBACK_FOLDER_PATH)==False:
        return messagebox.showerror(title='Error', message='You selected a wrong grading sheet or \nnumber of your feedback files do not match the grading. \n Please recheck your selections.')

    my_button2.config(text="Converting...",bg = 'grey')
    generate_feedbacks_for_upload(input_grade_file=GRADE_FILE,
        input_feedback_path=FEEDBACK_FOLDER_PATH,
        output_path=PDF_FOLDER_PATH)
    my_button2.config(text="Zip your feedbacks", bg = 'red',fg='white')
    pass

def run():
    Thread(target=convert_files).start()
    pass

my_button2.config(command=run)
my_button2.pack(side="bottom",fill=Y,expand=1) 
frame.pack()


"""
Process box
"""
frame = Frame(root)
textbox=Text(frame,bg='black',fg='white',height=5, width=40)
textbox.pack()

def redirector(inputStr):
    textbox.insert("end", inputStr+'\n')
    textbox.see('end')

sys.stdout.write = redirector
sys.stderr.write = redirector

frame.pack()



"""
Author box
"""
footer_frame = tk.Frame(root)
my_label = Label(footer_frame, text= 'Xia Cui (x.cui@mmu.ac.uk)')
footer_frame.pack(side="bottom", fill="x")
my_label.pack(side="right")

root.mainloop()