import PyPDF2 as ppdf
import docx
import re
from openpyxl import Workbook

cur_dir = str(input("Enter Input File Directory: "))
# cur_dir = 'D:\Billu Conversion'
file_name = str(input("Enter Input File Name: "))
# file_name = '_Mcq'
# file_name = 'MCQ 40'
# file_name = '8th-December-Friday-Admission'
# file_name = 'বাংলা বহুনির্বাচনী _ ক্রিয়ার কাল ,ক্রিয়ার ভাব ও অনুজ্ঞা'
# file_name = 'Medical Admission Quiz'
# file_name = 'অধ্যায় - ০২ আমার পথ'
file_name = "\\" + file_name
# file_type = str(input("Enter 1 for pdf or 2 for docx: "))
# file_type = 1
file_type = 2
file_type = '.pdf' if file_type == 1 else '.docx'
# lang = input("Enter 1 for english or 2 for bangla: ")
# lang = 1
# lang = 2
answer_cat = input("Enter 1 for answer only bold or 2 for answer separate: ")
# answer_cat = 1
source_flag = input("Enter 1 to extract sources or 2 to remove sources: ")
# source_flag = 1
# answer_cat = "Answer:"
# answer_cat = "Ans:"
file = cur_dir + file_name + file_type
out_file = cur_dir + file_name + '.xlsx'
wb = Workbook()
sh = wb.active
sh.title = 'MCQs'
sh["A1"] = 'Question'
sh["B1"] = 'Option 1'
sh["C1"] = 'Stat 1'
sh["D1"] = ''
sh["E1"] = 'Option 2'
sh["F1"] = 'Stat 2'
sh["G1"] = ''
sh["H1"] = 'Option 3'
sh["I1"] = 'Stat 3'
sh["J1"] = ''
sh["K1"] = 'Option 4'
sh["L1"] = 'Stat 4'
# sh["M1"] = ''
sh["M1"] = 'Source'

if file_type == '.pdf':
    lines = ''
    f1 = open(file, 'rb')
    f1_reader = ppdf.PdfReader(f1)
    f1_page_num = len(f1_reader.pages)
    for i in range(0, f1_page_num):
        f1_obj = f1_reader.pages[i]
        lines += f1_obj.extract_text()
    # print(lines)
elif file_type == '.docx':
    lines = ''
    doc = docx.Document(file)
    ans_bolds = []
    bolds = []
    fullText = []
    for para in doc.paragraphs:
        for run in para.runs:
            if run.bold:
                bolds.append((run.text).strip())
        fullText.append((para.text).strip())
    lines = '\n'.join(fullText)
    # print(lines)
    prev_bold = ''
    for index, bold in enumerate(bolds):
        if bold == prev_bold:
            bolds[index] = ""
        prev_bold = bold
    while "" in bolds:
        bolds.remove("")
    # print(bolds)
    for bold in bolds:
        if ')' in bold[0:3] or '.' in bold[0:3] or '।' in bold[0:3]:
            ans_bolds.append(bold.strip())
    # print(ans_bolds)
else:
    raise ValueError("Invalid file. Please input correct directory, name and type.")

# lines = re.sub(r'\[.+\]', '', lines)
if source_flag != 1:
    lines = re.sub(r'\[.+', '', lines)
if answer_cat == 1:
    lines = re.sub(r'উত্তর:.+', '', lines)
lines_list = lines.split('\n')
lines_list = [i.strip() for i in lines_list if i]
lines_list = [i.strip() for i in lines_list if i]

q_count = 2
if answer_cat == 1:
    len_q = 5
else:
    len_q = 6
for index, line in enumerate(lines_list):
    if index % len_q == 0:
        q_count = str(q_count)
        ans_flag = 0
        o1_stat = 0
        o2_stat = 0
        o3_stat = 0
        o4_stat = 0
        if source_flag == 1:
            src_line = re.search(r'\[.+', line)
            if src_line:
                src = src_line.group()
            else:
                src = ''
            q = re.sub(r'\[.+', '', line)
        else:
            q = line
        # o1 = lines_list[index + 1]
        try:
            o1 = lines_list[index+1]
        except IndexError:
            print('Error near line: ')
            print(line)
        # o2 = lines_list[index + 2]
        try:
            o2 = lines_list[index+2]
        except IndexError:
            print('Error near line: ')
            print(line)
        # o3 = lines_list[index + 3]
        try:
            o3 = lines_list[index+3]
        except IndexError:
            print('Error near line: ')
            print(line)
        # o4 = lines_list[index + 4]
        try:
            o4 = lines_list[index+4]
        except IndexError:
            print('Error near line: ')
            print(line)
        if answer_cat == 1:
            ans = ans_bolds[int(q_count)-2]
        else:
            # ans = lines_list[index + 5]
            try:
                ans = lines_list[index+5]
            except IndexError:
                print('Error near line: ')
                print(line)
            if ':' in ans:
                ans_flag = 1
                ans = ans.split(':')[-1]
        ans = ans.strip()
        if 'a' in ans[0:3] or 'A' in ans[0:3] or 'ক' in ans[0:3]:
            if ')' in ans[0:3] or '.' in ans[0:3] or '।' in ans[0:3] or ans_flag == 1 or len(ans) == 1:
                o1_stat = 1
        elif 'b' in ans[0:3] or 'B' in ans[0:3] or 'খ' in ans[0:3]:
            if ')' in ans[0:3] or '.' in ans[0:3] or '।' in ans[0:3] or ans_flag == 1 or len(ans) == 1:
                o2_stat = 1
        elif 'c' in ans[0:3] or 'C' in ans[0:3] or 'গ' in ans[0:3]:
            if ')' in ans[0:3] or '.' in ans[0:3] or '।' in ans[0:3] or ans_flag == 1 or len(ans) == 1:
                o3_stat = 1
        elif 'd' in ans[0:3] or 'D' in ans[0:3] or 'ঘ' in ans[0:3]:
            if ')' in ans[0:3] or '.' in ans[0:3] or '।' in ans[0:3] or ans_flag == 1 or len(ans) == 1:
                o4_stat = 1
        else:
            o1_stat = 'Exception'
            o2_stat = 'Exception'
            o3_stat = 'Exception'
            o4_stat = 'Exception'
        q = re.sub(r'^\S+[).।]', '', q)
        q = re.sub(r'Question:', '', q)
        q = q.strip()
        o1 = re.sub(r'^\S+[).।]', '', o1)
        o1 = o1.strip()
        o2 = re.sub(r'^\S+[).।]', '', o2)
        o2 = o2.strip()
        o3 = re.sub(r'^\S+[).।]', '', o3)
        o3 = o3.strip()
        o4 = re.sub(r'^\S+[).।]', '', o4)
        o4 = o4.strip()
        sh["A" + q_count] = q
        sh["B" + q_count] = o1
        sh["C" + q_count] = o1_stat
        sh["D" + q_count] = ''
        sh["E" + q_count] = o2
        sh["F" + q_count] = o2_stat
        sh["G" + q_count] = ''
        sh["H" + q_count] = o3
        sh["I" + q_count] = o3_stat
        sh["J" + q_count] = ''
        sh["K" + q_count] = o4
        sh["L" + q_count] = o4_stat
        if source_flag == 1:
            sh["M" + q_count] = src
        else:
            sh["M" + q_count] = ''
        q_count = int(q_count)
        q_count += 1

wb.save(out_file)   # Creating output excel file
print(f'Generated mcq excel : {out_file}')
