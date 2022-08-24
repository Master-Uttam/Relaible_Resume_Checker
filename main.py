import shutil
import time
from PyPDF2 import PdfReader
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.feature_extraction.text import CountVectorizer
import docx2txt
import os


def check_file(resume_file):
    desc_file = r"C:\Users\Admin\PycharmProjects\Reliable\description.docx"
    job_description = docx2txt.process(desc_file)
    content = [job_description, resume_file]
    cv = CountVectorizer()
    matrix = cv.fit_transform(content)
    similarity_matrix = cosine_similarity(matrix)

    percent = int(similarity_matrix[0][1] * 100)
    print('Resume Matching : \t' + str(percent) + '%')
    return percent


def paste_folder(transfer_file):
    dst = r'C:\Users\Admin\PycharmProjects\Reliable\Filtered_Resume'
    shutil.copy(transfer_file, dst)


if __name__ == '__main__':
    print("Step 1: Go to description.docx and save the description.")
    time.sleep(2)
    os.startfile(r'C:\Users\Admin\PycharmProjects\Reliable\description.docx')
    print("Step 2: Go to Resume_folder and paste all the resumes")
    time.sleep(2)
    os.startfile(r'C:\Users\Admin\PycharmProjects\Reliable\Resume_folder')
    percentage = int(input("Step 3: Enter comparison percentage between resume & description?\t"))
    resume_folder = r"C:\Users\Admin\PycharmProjects\Reliable\Resume_folder"
    resume_list = os.listdir(resume_folder)
    print(resume_list)

    for i in range(len(resume_list)):
        resumes_path = os.path.join(resume_folder, resume_list[i])
        print(f"\nExecuting this file: {resume_list[i]}")

        try:

            if '.pdf' in resumes_path:
                reader = PdfReader(resumes_path)
                number_of_pages = len(reader.pages)
                page = reader.pages[0]
                resume_pdf_txt = page.extract_text()
                d = check_file(resume_pdf_txt)
                if percentage <= int(d):
                    paste_folder(resumes_path)
                    print(f"{resume_list[i]} is transferred to Filtered_Resume \n")

            elif '.docx' in resumes_path:
                resume_docx_txt = docx2txt.process(resumes_path)
                t = check_file(resume_docx_txt)
                if percentage <= int(t):
                    paste_folder(resumes_path)
                    print(f"{resume_list[i]} is transferred to Filtered_Resume \n")

            else:
                print("This file is not supported")

        except:
            print(f"Error {resume_list[i]} \n")

    print("Execution is Done.")
    os.startfile("Filtered_Resume")
