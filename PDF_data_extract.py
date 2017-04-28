############################################################################################################
#This script extract data from the pdf and text file 
# script created by - Pulin Kulshrestha
# Date - 3/9/2017
#
############################################################################################################

import os
import re
import xlsxwriter
import datetime

#Time log
print(datetime.datetime.now())



# Set path for files and folders
current_dir = os.path.dirname(os.path.realpath(__file__))
input_folder = '\\Input\\'
output_folder = '\\Output\\'
temp_folder = '\\Temp\\'
negative ='\\Negative_Tone\\'
output_excel_name = 'output_file.xlsx'
python_path = 'C:\\Users\\Pulin05\\AppData\\Local\\Programs\\Python\\Python35-32\\Scripts\\'

#tagged pdf pattern
Letter_date = '(<P MCID="[0-9]+">)(\s*[A-z]+ [0-9]+, [0-9]...)'
headers_pattern = '(<P MCID="[0-9]+">)([A-z 0-9\.]+)(<\/P><P MCID="[0-9]+"> <\/P><P MCID="[0-9]+">[0-9]+\. )'
pdf_name_pattern = '[A-z 0-9]+'
company_name = '(Re:| Re: |Re:\t)([A-z 0-9/.]+)'
filed_date = '(<P MCID="[0-9]+">Filed )([A-z]+ [0-9]+, [0-9]...)'
issue_pattern= '(<P MCID="[0-9]+"> </P><P MCID="[0-9]+">)([0-9]+)(\.)'
issue_pattern2 = '<[A-Z ]+="[0-9]+">[0-9]+\.'
limited_to_pattern = 'limited to'
closing_letter_pattern = 'completed|completion'
mcid_pattern = '(<P MCID="[0-9]+">)'

#html text pattern
Letter_date_html = '(font-size:11px">)(\s*[A-z]+ [0-9]+, [0-9]...)'
filed_date_html = '(font-size:11px">Filed)(\s*[A-z]+ [0-9]+, [0-9]...)'
issue_pattern_html = '(font-size:11px">)([0-9]+\.)(</span><span)'

# text file pattern

Letter_date_text ='([A-z]+ [0-9]+, [0-9]...)'
issue_pattern_text ='(\n)([0-9]+\.)(\s)'
filed_date_text ='(Filed )([A-z]+ [0-9]+, [0-9]...)'

#Fetch list of folders from the input folder
folder_list = os.listdir(os.path.join(current_dir + input_folder))
print (folder_list)

# open Ecel workbook and write
workbook = xlsxwriter.Workbook(os.path.join(current_dir + output_folder + output_excel_name))
worksheet = workbook.add_worksheet()

# Widen column to make the text clearer
worksheet.set_column('A:K', 20)


# Add a bold format to use to highlight cells
bold = workbook.add_format({'bold': True})

# Text with formatting
worksheet.write('A1', 'CIK', bold)
worksheet.write('B1', 'Folder_Name', bold)
worksheet.write('C1', 'File_Name', bold)
worksheet.write('D1', 'Company_Name', bold)
worksheet.write('E1', 'Letter_date', bold)
worksheet.write('F1', 'Filed_date', bold)
worksheet.write('G1', 'Number_of_Issues', bold)
worksheet.write('H1', 'Word_Count', bold)
worksheet.write('I1', 'Tone_Count', bold)
worksheet.write('J1', 'Limited_to_flag', bold)
worksheet.write('K1', 'Closing_flag', bold)

# Start from the first cell below the headers.
row = 1
col = 0

# File Traversing
folder_c = 0
sub_folder_c = 0

# counter for word count
#numword = 0

for folder in folder_list:
    for sub_folder in os.listdir(os.path.join(current_dir + input_folder + folder_list[folder_c])):
        sub_folder_list =os.listdir(os.path.join(current_dir + input_folder + folder_list[folder_c]))
        for files in os.listdir(os.path.join(current_dir + input_folder + folder_list[folder_c],sub_folder_list[sub_folder_c])):
            if files.endswith(".txt"):
                #print (folder_list[folder_c],sub_folder_list[sub_folder_c],files)
                input_text_file = open(os.path.join(current_dir + input_folder + folder_list[folder_c],sub_folder_list[sub_folder_c], files), 'r',encoding="utf8")
                file_data_text = input_text_file.read()
                file_data_lower_text = file_data_text.lower()
                wordlist = file_data_lower_text.split()
                numword = 0
                numword += len(wordlist)
                negative_count = 0
                neg_list = open(os.path.join(current_dir + negative + 'negative_list.txt'),'r',encoding='utf-8-sig')
                neg_words=neg_list.read()
                negative_words= neg_words.split('\n')
                for word in wordlist:
                    if word in negative_words:
                        negative_count += 1
                worksheet.write(row, col, folder_list[folder_c])
                worksheet.write(row, col+1, sub_folder_list[sub_folder_c])
                worksheet.write(row, col+2, files)
                if len(re.findall(company_name, file_data_text))!= 0:
                    comp_name_text = re.findall(company_name, file_data_text)[0][1]
                    worksheet.write(row, col+3, comp_name_text.strip())
                else:
                    worksheet.write(row, col+3,'Company Name Not Found')
                if len(re.findall(Letter_date_text, file_data_text)) !=0:
                    Letter_dt_text = re.findall(Letter_date_text, file_data_text)[0]
                    worksheet.write(row, col+4, Letter_dt_text.strip())
                else:
                    worksheet.write(row, col+4,'No Letter Date') 
                if len(re.findall(filed_date_text, file_data_text)) != 0:
                    worksheet.write(row, col+5, re.findall(filed_date_text, file_data_text)[0][1])
                else:
                    worksheet.write(row, col+5,'No Filed Date')
                if len(re.findall(issue_pattern_text, file_data_text)) !=0:
                       worksheet.write(row, col+6, str(len(re.findall(issue_pattern_text, file_data_text))))
                else:
                       worksheet.write(row, col+6,0)
                worksheet.write(row, col+7, numword)
                worksheet.write(row, col+8, negative_count)
                if len(re.findall(limited_to_pattern, file_data_text)) != 0:
                    worksheet.write(row, col+9,1)
                else:
                    worksheet.write(row, col+9,0)
                if len(re.findall(closing_letter_pattern, file_data_text)) != 0:
                    worksheet.write(row, col+10,1)
                else:
                    worksheet.write(row, col+10,0)                    
                row += 1
            else:
                os.system(python_path + 'pdf2txt.py -o ' + os.path.join(current_dir + temp_folder) + 'temp_tag.txt -t tag ' + os.path.join(current_dir + input_folder + folder_list[folder_c],sub_folder_list[sub_folder_c], files ))
                os.system(python_path + 'pdf2txt.py -o ' + os.path.join(current_dir + temp_folder) + 'temp.txt ' + os.path.join(current_dir + input_folder + folder_list[folder_c],sub_folder_list[sub_folder_c], files ))
                input_file = open(os.path.join(current_dir + temp_folder + 'temp_tag.txt'), 'r',encoding="utf8")
                input_file_no_tag = open(os.path.join(current_dir + temp_folder + 'temp.txt'), 'r',encoding="utf8")
                file_data = input_file.read()
                file_data_no_tag = input_file_no_tag.read()
                file_data_lower = file_data_no_tag.lower()
                wordlist = file_data_lower.split()
                numword = 0
                numword += len(wordlist)
                negative_count = 0
                neg_list = open(os.path.join(current_dir + negative + 'negative_list.txt'),'r',encoding='utf-8-sig')
                neg_words=neg_list.read()
                negative_words= neg_words.split('\n')
                for word in wordlist:
                    if word in negative_words:
                        negative_count += 1
                #print(re.findall(company_name, file_data))
                worksheet.write(row, col, folder_list[folder_c])
                worksheet.write(row, col+1, sub_folder_list[sub_folder_c])
                worksheet.write(row, col+2, files)
                if len(re.findall(company_name, file_data))!= 0 and len(re.findall(mcid_pattern, file_data)) != 0:
                    comp_name = re.findall(company_name, file_data)[0][1]
                    worksheet.write(row, col+3, comp_name.strip())
                elif len(re.findall(mcid_pattern, file_data)) == 0:
                    os.system(python_path + 'pdf2txt.py -o ' + os.path.join(current_dir + temp_folder) + 'temp_html.txt -t html ' + os.path.join(current_dir + input_folder + folder_list[folder_c],sub_folder_list[sub_folder_c], files ))
                    input_file_html = open(os.path.join(current_dir + temp_folder + 'temp_html.txt'), 'r',encoding="utf8")
                    file_data_html = input_file_html.read()
                    if len(re.findall(company_name, file_data_html)) !=0:
                        comp_name1 = re.findall(company_name, file_data_html)[0][1]
                        worksheet.write(row, col+3, comp_name1.strip())
                else:
                    worksheet.write(row, col+3,'Company Name Not Found')
                    #print (re.findall(Letter_date, file_data))
                if len(re.findall(Letter_date, file_data)) !=0 and len(re.findall(mcid_pattern, file_data)) != 0:
                    Letter_dt = re.findall(Letter_date, file_data)[0][1]
                    worksheet.write(row, col+4, Letter_dt.strip())
                elif len(re.findall(mcid_pattern, file_data)) == 0 and len(re.findall(Letter_date_html, file_data_html)) !=0:
                    Letter_dt_html = re.findall(Letter_date_html, file_data_html)[0][1]
                    worksheet.write(row, col+4, Letter_dt_html.strip())
                else:
                    worksheet.write(row, col+4,'No Letter Date') 
                if len(re.findall(filed_date, file_data)) != 0 and len(re.findall(mcid_pattern, file_data)) != 0:
                    filed_dt = re.findall(filed_date, file_data)[0][1]
                    worksheet.write(row, col+5, filed_dt.strip())
                elif len(re.findall(mcid_pattern, file_data)) == 0 and len(re.findall(filed_date_html, file_data_html)) !=0:
                    filed_dt_html = re.findall(filed_date_html, file_data_html)[0][1]
                    worksheet.write(row, col+5, filed_dt_html.strip())                    
                else:
                    worksheet.write(row, col+5,'No Filed Date')
                if len(re.findall(issue_pattern, file_data)) !=0 and len(re.findall(mcid_pattern, file_data)) != 0:
                    worksheet.write(row, col+6, str(len(re.findall(issue_pattern, file_data))))
                elif len(re.findall(issue_pattern, file_data)) ==0 and len(re.findall(mcid_pattern, file_data)) != 0:
                    worksheet.write(row, col+6, str(len(re.findall(issue_pattern2, file_data))))
                elif len(re.findall(mcid_pattern, file_data)) == 0:
                    worksheet.write(row, col+6, str(len(re.findall(issue_pattern_html, file_data_html))))
                else:
                    worksheet.write(row, col+6,0)
##                    print (re.findall(headers_pattern, file_data))
                worksheet.write(row, col+7, numword)
                worksheet.write(row, col+8, negative_count)
                if len(re.findall(limited_to_pattern, file_data_no_tag)) != 0:
                    worksheet.write(row, col+9,1)
                else:
                    worksheet.write(row, col+9,0)
                if len(re.findall(closing_letter_pattern, file_data_no_tag)) != 0:
                    worksheet.write(row, col+10,1)
                else:
                    worksheet.write(row, col+10,0)                    
                row += 1
        sub_folder_c +=1
        if sub_folder_c > len(sub_folder_list) - 1:
            break
    print ("Folder "+ folder_list[folder_c] + " completed")
    folder_c += 1
    sub_folder_c = 0
    if folder_c > len(folder_list)-1:
        break

workbook.close()
input_file.close()

#Time log
print(datetime.datetime.now())
