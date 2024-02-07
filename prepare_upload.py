import pandas as pd
import os,sys,time
import zipfile
from pathlib import Path
import docx
from glob import glob
from docx2pdf import convert
import shutil


GENERATED_FOLDER_NAME = 'generated_files'
FEEDBACK_PDF_FOLDER_NAME = 'feedbacks-pdf'
FEEDBACK_FOLDER_NAME = 'feedbacks'
PDF_FOLDER_NAME = 'pdf'
graded_file = 'test_files/Grading sheet-moodle.csv'
output_file = 'test_files/Grading sheet.csv'
output_path = 'test_files'

def create_folder(path):
    # Check whether the specified path exists or not
    isExist = os.path.exists(path)
    if not isExist:
        # Create a new directory because it does not exist
        os.makedirs(path)
        print(f"{path} is created!")
        return 1
    else:
        print(f"{path} exists.")
        return -1
pass

def check_exist(path):
    return os.path.exists(path)

def match_name_list(csv_file,input_feedback_path):
    df = pd.read_csv(csv_file)
    student_names = df['Surname/Name'].tolist()
    student_names = [reverse_name(s) for s in student_names]
    if len(glob(input_feedback_path+'/'+f'*.docx'))==len(student_names) or\
        len(glob(input_feedback_path+'/'+f'*.pdf'))==len(student_names):
        return True
    else:
        return False
    pass

def get_names_ids(path=graded_file):
    df = pd.read_csv(path)
    submission_ids = df['Submission id'].tolist()
    student_names = df['Surname/Name'].tolist()
    student_names = [reverse_name(s) for s in student_names]
    return student_names,submission_ids

def fill_in_marks(df,input_path=graded_file,output_path=output_file):
    origin = pd.read_csv(input_path)
    student_names = df['Student Name'].tolist()
    df['Surname/Name'] = [reverse_name(s) for s in student_names]
    new_df = pd.merge(origin, df, on='Surname/Name')
    new_df['Grade'] = new_df['Total'].tolist()
    new_df = new_df[origin.columns]
    new_df.to_csv(output_file,index=False)
    pass


def generate_feedbacks_for_upload(input_grade_file=graded_file,
        input_feedback_path=FEEDBACK_FOLDER_NAME,
        output_path=PDF_FOLDER_NAME):
    pdf_feedback_path = output_path+'/'+FEEDBACK_PDF_FOLDER_NAME
    student_names,submission_ids = get_names_ids(input_grade_file)
    
    if len(glob(input_feedback_path+'/'+f'*.docx'))>0:
        print('Converting docx to pdf...')
        create_folder(pdf_feedback_path)
        convert(input_feedback_path+'/',pdf_feedback_path+'/')
        print('Finished.')
    elif len(glob(input_feedback_path+'/'+f'*.pdf'))>0:
        pdf_feedback_path = input_feedback_path
    else:
        print("no docx or pdf in your folder")
    output_pdf_path = output_path+'/'+PDF_FOLDER_NAME 
    create_folder(output_pdf_path)
    for student_name,submission_id in zip(student_names,submission_ids):
        student_feedback = glob(pdf_feedback_path+'/'+f'*{student_name}*.pdf')[0]
        output_pdf_file = output_pdf_path+'/'+submission_id+'.pdf'
        shutil.copy(student_feedback,output_pdf_file)
    zip_all_feedbacks(input_path=output_pdf_path,output_path=output_path+'/feedback')
    pass


def reverse_name(string):
    s = string.split()[::-1]
    l = []
    for i in s:
        # appending reversed words to l
        l.append(i)
    return " ".join(l)

def zip_all_feedbacks(input_path=PDF_FOLDER_NAME,output_path=output_path):
    shutil.make_archive(output_path, 'zip', input_path)
    pass

def generate_sendoff_marksheet(input_path=output_file,output_path=output_path):
    df = pd.read_csv(input_path)
    student_ids = [int(email.replace('@stu.mmu.ac.uk','')) for email in df["Username"].tolist()]
    student_names = df['Surname/Name'].tolist()
    submission_times = df['Submission time'].tolist()
    grades = df['Grade'].tolist()
    results = {
        'MMU ID':student_ids,
        'Student Name':student_names,
        'Submission Time':submission_times,
        'Grade':grades
    }
    new_df = pd.DataFrame(results)
    new_df.to_excel(output_path+'/'+Path(input_path).stem+'.xlsx', index=False)
    pass


if __name__ == '__main__':
    # generate_feedbacks_for_upload()
    # generate_sendoff_marksheet()
    print('hello world')