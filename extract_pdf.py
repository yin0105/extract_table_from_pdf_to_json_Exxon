import tabula
import pdfplumber
import pandas as pd
import sys
import os
from datetime import datetime

def write_into_file(content):
    # my_file.write("{\n")
    my_file.write(content)

# Change multiline of a cell to json
def multiline_to_json(s):
    rr = ""
    for ss in s.splitlines():
        if ss.find(":") > -1:
            sss = '"' + remove_special_characters(ss.split(":", 1)[0]) + '": "' + remove_special_characters(ss.split(":", 1)[1]) + '"'
        else:
            sss = '"' + remove_special_characters(ss) + '": ""'
        if rr != "" and sss != "": rr += ", "
        rr += sss
    return rr

def remove_special_characters(content):
    # return ''.join(e for e in content if e.isalnum() or e == ' ')
    return content.replace('"', '\\"')

# datetime object containing current date and time
now = datetime.now()
dt_string = now.strftime("%Y-%m-%d_%H-%M-%S")

# print environment variables
# s = os.environ["PATH"].split(';')
# for item in s:
    # print(item)

# Get pdf file name from command line
if len(sys.argv) == 1:
    print("INPUT TYPE: python extract_pdf.py pdf_file_name [output_path]")
    exit()
elif len(sys.argv) >= 2:
    pdf_file = sys.argv[1]
    output_path = "."
    if len(sys.argv) > 2: output_path = sys.argv[2]

filename = os.path.basename(pdf_file)
filename = filename[:filename.rfind(".")]

# Get row data of the tables in a pdf file using tabula and Save them as array
df = tabula.read_pdf(pdf_file, pages="all", guess = False, multiple_tables = True) 
t_word = []
first_row = []
for a in df:
    first_row_tmp = []
    t_word_page = []
    for i in range(len(a.index)):
        count = 0
        t_word_tmp = []
        for c in a:
            ss = str(a.loc[i, c])
            if ss != "nan" and ss.strip() != "":
                t_word_tmp.append(" ".join(ss.splitlines()))
        t_word_page.append(t_word_tmp)
        if len(first_row_tmp) <2 and len(t_word_tmp) == 1 :
            first_row_tmp.append(t_word_tmp[0])
            # print(t_word_tmp[0])
        
        # print(str(i) + "  " + str(c) + "  " + "::".join(t_word_tmp))
    first_row.append(first_row_tmp[1])
    print("first_row :: " + first_row_tmp[1])
    t_word.append(t_word_page)


# Parse a pdf file with pdfplumber
pdf = pdfplumber.open(pdf_file) 

# Open output text file
# filename += "_" + f"{dt_string}.txt"
filename = "output.txt"
my_file = open(f"{output_path}\{filename}", "w", encoding="utf8")
print(f"{output_path}\{filename}")
write_into_file("{\n")


# Main Working Flow
page_num = 0 # current page number (1 ~ )
write_started = False # Whether writing started




print(remove_special_characters('233"C'))


# Iterate through pages
for p0 in pdf.pages:
    page_num += 1 # Increase current page
    cell = [] # An array of index of cell data
    word = [] # An array of cell data
    # Parse the tables of PDF
    table = p0.extract_table(table_settings={"vertical_strategy": "lines"})    
    # Change table to DataFrame of Pandas
    df = pd.DataFrame(table[1:], columns=table[0])

    # Set an array of index of cell data and an array of cell data
    # Iterate through rows
    # print(df)
    # my_file.write(df)
    bit_data_flag = False
    
    # first row
    cell_temp = []
    word.append(first_row[page_num - 1])
    for j in  range(len(df.columns)):
        cell_temp.append(len(word) - 1)
    cell.append(cell_temp)

    for i in range(len(df.index)):
        # my_file.write("line " + str(i) + " : ")
        count = 1
        cell_temp = []
        

        
        # Iterate through columns
        for j in  range(len(df.columns)):
            
            ss = str(df.iloc[i, j])          
            # my_file.write(str(j) + ":" + ss + "  ")
            if ss != "None" :
                word.append(ss)
                count += 1
            else:
                if j == 0 or (str(df.iloc[i, 0]) == "Output" and (j == 1 or j == 15)) or (str(df.iloc[i, 0]) == "Output" and (j == 1 or j == 15)):
                    word.append("")
                    count += 1

            cell_temp.append( len(word) - 1)
        cell.append(cell_temp)
        # my_file.write("\n")

    # Main Working Flow
    # Iterate through rows
    for i in range(len(df.index)):  
        # Iterate through columns
        for j in  range(len(df.columns)):
            ss = ""
            if cell[i][j] == -1: continue
            # Get cell number of data
            for k in range(j + 1, len(df.columns)): 
                if cell[i][k] != cell[i][j]: break
            else:
                if k == len(df.columns) - 1: k = len(df.columns)
            k -= 1
            # data (or key of group)
            ww_2 = " ".join(word[cell[i][j]].splitlines())
            if ww_2 == "Management Summary" : print("ii = " + str(ii) + "  i = " + str(i) + "  j= " + str(j) + " k = " + str(k)) 
            # Whether data is group
            if i < len(df.index)-1 and k > j: 
                # print( str(i) + "  " + str(j) + "  " + str(k) + "  ::  " + ww_2)
                if (j == 0 or (j > 0 and cell[i+1][j-1] != cell[i+1][j])) and (k == len(df.columns) - 1 or (k < len(df.columns) - 1 and cell[i+1][k] != cell[i+1][k+1])) and (cell[i+1][j] != cell[i+1][k] or ww_2 in ["Well Information", "Management Summary", "Mud Checks", "Mud Volume Summary", "String No. 3", "String No. 3  - Bit Information", "Wellbore Information", "Deviation Survey (Tie point) -  TABU-B19 Actual Survey",       "COSTS", "STATUS", "MUD CHECK", "BIT DATA", "GAS READINGS", "MUD VOLUME", "PERSONNEL", "PUMP/HYDRAULICS"]) :
                    # print( str(i) + "  " + str(j) + "  " + str(k) + "  ::  " + ww_2)
                    # Get the number of rows in a group
                    for ii in range(i+2, len(df.index)): 
                        if ww_2 == "Management Summary" : print("ii = " + str(ii))
                        if ww_2 == "Well Information" and word[cell[ii][j]] == "Daily Operations Information": break
                        if ww_2 == "Daily Operations Information" and word[cell[ii][j]] == "Time Log": break
                        if ww_2 == "Management Summary" and word[cell[ii][j]] == "Mud Checks": break 
                        if ww_2 == "Mud Checks" and word[cell[ii][j]] == "Mud Volume Summary": break
                        if ww_2 == "Mud Volume Summary" and word[cell[ii][j]] == "Daily Job Supply Summary": break
                        if ww_2 == "String No. 3" and word[cell[ii][j]] == "String No. 3  - Bit Information": break
                        if ww_2 == "String No. 3  - Bit Information" and word[cell[ii][j]] == "String Components": break
                        if ww_2 == "Deviation Survey (Tie point) -  TABU-B19 Actual Survey" and word[cell[ii][j]] == "Survey Data - Shows all surveys entered during the report period": break


                        if not ((j == 0 or (j > 0 and cell[ii][j-1] != cell[ii][j])) and (k == len(df.columns) - 1 or (k < len(df.columns) - 1 and cell[ii][k] != cell[ii][k+1])) and (cell[ii][j] != cell[ii][k] or ww_2 in ["Well Information", "Daily Operations Information", "Management Summary", "Mud Checks", "Mud Volume Summary", "String No. 3", "String No. 3", "String No. 3  - Bit Information", "Deviation Survey (Tie point) -  TABU-B19 Actual Survey",      "COSTS", "STATUS", "MUD CHECK", "BIT DATA", "GAS READINGS", "MUD VOLUME", "PERSONNEL", "PUMP/HYDRAULICS"])) : break
                        if not(ww_2 in ["Well Information", "Daily Operations Information", "Management Summary", "Mud Checks", "Mud Volume Summary", "String No. 3", "String No. 3", "String No. 3  - Bit Information", "Deviation Survey (Tie point) -  TABU-B19 Actual Survey",       "COSTS", "STATUS", "MUD CHECK", "BIT DATA", "GAS READINGS", "MUD VOLUME", "PERSONNEL", "PUMP/HYDRAULICS", "Wellbore Information"]):
                            for jj in range(j+1, k+1): 
                                if cell[ii][jj] - cell[ii-1][jj] != cell[ii][j] - cell[ii-1][j]:break
                            else:
                                continue
                            break
                    else:
                        if ii < i + 2: ii = i + 2
                        if ww_2 == "Management Summary" : print("ii = " + str(ii))
                        if ii == len(df.index) - 1 : ii += 1
                    

                    if ww_2 == "PUMP/HYDRAULICS": print ("ii = " + str(ii))
                    
                    if ww_2 == "Time Log": # if name of group is "OPERATION SUMMARY"
                        pre_cell = -2
                        header = []
                        # Get the column headers in a group
                        for jj in range(j, k + 1):
                            if cell[i+1][jj] != pre_cell:
                                header.append(" ".join(word[cell[i+1][jj]].splitlines()))
                                pre_cell = cell[i+1][jj]
                            cell[i+1][jj] = -1
                        # Get the columns data in group
                        for iii in range(i + 2, ii): 
                            group_column_num = 0
                            tmp_list = []
                            pre_cell = -2
                            # Get cell data
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    tmp_list.append( word[cell[iii][jj]])
                                    pre_cell = cell[iii][jj]
                                cell[iii][jj] = -1

                            # Split cell data to multiple rows
                            while len(tmp_list) > 0 and len(tmp_list[0].splitlines()) > 0:

                                # Get data of the first line of the first column
                                comp_str = tmp_list[0].splitlines()[-1].strip()  
                                
                                # Compare data of word array with comp_str and Set data
                                for t_a in t_word[page_num - 1]:
                                    for t_b in t_a:
                                        # Whether Is there comp_str in data of word array
                                        if t_b.find(comp_str.strip()) > -1 :
                                            f = False
                                            for c in tmp_list:  
                                                if len(c.splitlines()) == 0: continue                                           
                                                for t_c in t_a:
                                                    if t_c.find(c.splitlines()[-1].strip()) > -1:  
                                                        f = True
                                                        break
                                                if f == False: break
                                            # If there is comp_str in data of word array, check other data
                                            if f == True:                                                
                                                for t_c in t_a:
                                                    if len(tmp_list) == 0: break
                                                    if len(tmp_list[0].splitlines()) ==0: break
                                                    if t_c.find(tmp_list[0].splitlines()[-1].strip()) == -1: continue
                                                    ssss = ""
                                                    tt = tmp_list[:]
                                                    for c in range(len(tmp_list)): 
                                                        sss = ""                                               
                                                        if len(tmp_list[c].splitlines()) == 0 :
                                                            sss = header[c] + ":   "
                                                            continue    

                                                        tt[c] =  tmp_list[c]                                                 
                                            
                                                        for c_len in range(len(tt[c].splitlines())):
                                                            
                                                            for t_d in t_a:
                                                                if t_d.find(tt[c].splitlines()[len(tt[c].splitlines())- c_len -1].strip()) > -1:
                                                                    cc = "  ".join(tt[c].splitlines()[len(tt[c].splitlines())- c_len -1:])
                                                                    sss = '"' + header[c] + '": "' + remove_special_characters(cc) + '" '
                                                                    tt[c] = "\n".join(tt[c].splitlines()[: len(tt[c].splitlines())- c_len -1])
                                                                    break
                                                            else:
                                                                continue
                                                            break
                                                        else:
                                                            break
                                                        if ssss != "": ssss += ", "
                                                        ssss += sss
                                                    else:
                                                        if ss == "":
                                                            ss = '{' + ssss + '}'
                                                        else:    
                                                            ss = '{' + ssss + '}, ' + ss
                                                        tmp_list = tt[:]

                                                if ss != "": ss += "\n"
                                                if len(tmp_list) == 0: 
                                                    # my_file.write("\n")
                                                    # my_file.write(ss)
                                                    write_into_file("\n")
                                                    write_into_file(ss)
                                                    break
                    elif ww_2 in ["Daily Offline Time Log Summary", "Drilling Parameters (Fast Drill)", "Daily Job Supply Summary", "String Components", "Daily Drilling Parameters and Hydraulics - If hydraulics are blank, check \"String Detail by Job - Hydr & AV\" report for errors", "Wellbore (Hole) Sections", "Survey Data - Shows all surveys entered during the report period", "Casing Strings - Only shows information for the longest casing component in a string - for other components see casing detail report"]: # if name of group is ...
                        print("ww_2" + ww_2)
                        # col header
                        pre_cell = -2
                        header = []
                        # Get the columns data in group
                        for jj in range(j, k + 1):
                            if cell[i+1][jj] != pre_cell:
                                header.append(" ".join(word[cell[i+1][jj]].splitlines()))
                                pre_cell = cell[i+1][jj]
                            cell[i+1][jj] = -1
                        # Get the columns data in group
                        for iii in range(i + 2, ii): 
                            header_cc = 0
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if header[header_cc] != "":
                                        if sss != "": sss += ", "
                                        sss += '"' + header[header_cc] + '": "' + remove_special_characters(" ".join(word[cell[iii][jj]].splitlines())) + '" '
                                    pre_cell = cell[iii][jj]
                                    header_cc += 1
                                cell[iii][jj] = -1
                            sss = '{' + sss + '}'
                            if ss != "": ss += ", "
                            ss += sss
                    elif ww_2 in ["PERSONNEL"]: 
                        # Double Column Header
                        pre_cell = -2
                        sss = ""
                        for jj in range(j, k + 1):
                            if cell[i+1][jj] != pre_cell:
                                if sss != "": sss += ", "
                                sss += '"' + '": "'.join(" ".join(word[cell[i+1][jj]].splitlines()).split(":")) + '"'
                                pre_cell = cell[i+1][jj]
                            cell[i+1][jj] = -1
                        # sss = '"' + '": "'.join(sss.split(":")) + '"'
                        sss = '{' + sss + '}'
                        if ss != "": ss += ", "
                        ss += sss
                        # Get the columns data in group
                        pre_cell = -2
                        header = []
                        for jj in range(j, k + 1):
                            if cell[i+2][jj] != pre_cell:
                                header.append(" ".join(word[cell[i+2][jj]].splitlines()))
                                pre_cell = cell[i+2][jj]
                            cell[i+2][jj] = -1
                        # Get the columns data in group
                        for iii in range(i + 3, ii): 
                            header_cc = 0
                            sss = ""
                            hh_count = 0
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if header[header_cc] != "":
                                        hh_count += 1
                                        if sss != "": 
                                            if hh_count == 3:
                                                sss += "}, {"
                                            else:
                                                sss += ", "
                                        sss += '"' + header[header_cc] + '": "' + remove_special_characters(word[cell[iii][jj]]) + '" '
                                    pre_cell = cell[iii][jj]
                                    header_cc += 1
                                cell[iii][jj] = -1
                            sss = '{' + sss + '}'
                            if ss != "": ss += ", "
                            ss += sss
                    elif ww_2 == "SUPPORT CRAFT": # if name of group is ...
                        pre_cell = -2
                        header = []
                        # Get the columns data in group
                        for jj in range(j, k + 1):
                            if cell[i+1][jj] != pre_cell:
                                header.append(" ".join(word[cell[i+1][jj]].splitlines()))
                                pre_cell = cell[i+1][jj]
                            cell[i+1][jj] = -1
                        # Get the columns data in group
                        for iii in range(len(word[cell[i+2][j]].splitlines())): 
                            header_cc = 0
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[i+2][jj] != pre_cell:
                                    if header[header_cc] != "":
                                        if sss != "": sss += ", "
                                        if len(word[cell[i+2][jj]].splitlines()) == 0 :
                                            sss += '"' + remove_special_characters(header[header_cc]) + '": "" '
                                        else:
                                            sss += '"' + remove_special_characters(header[header_cc]) + '": "' + remove_special_characters(word[cell[i+2][jj]].splitlines()[iii]) + '" '
                                    pre_cell = cell[i+2][jj]
                                    header_cc += 1
                                # cell[i+2][jj] = -1
                            sss = '{' + sss + '}'
                            if ss != "": ss += ", "
                            ss += sss
                    elif ww_2 in [         "BULKS", "PUMP/HYDRAULICS", "BIT DATA"]: 
                        # Row Header
                        # if ww_2 == "BIT DATA":
                        #     print("j = " + str(j) + "  k = " + str(k))
                        #     my_file.write("BBBBBBBBBBBBBBBBBBBBBBBBBBBB")
                        for jj in range(j + 1, k + 1):
                            # if ww_2 == "BIT DATA":
                            #     my_file.write(str(cell[i+1][jj]) + "==" + str(cell[i+1][jj-1]) + "\n")
                            if cell[i+1][jj] == cell[i+1][jj-1]: continue
                            sss = ""
                            for iii in range(i + 1, ii):                                 
                                if ' '.join(word[cell[iii][j]].splitlines()) == '' : continue
                                if sss != "" : sss += ", "
                                sss += '"' + ' '.join(word[cell[iii][j]].splitlines()) + '": "' + remove_special_characters(' '.join(word[cell[iii][jj]].splitlines())) + '"'
                                # print(str(iii) + "  " + str(jj) + "  " + " ".join(word[cell[iii][j]].splitlines()) + ": " +" ".join(word[cell[iii][jj]].splitlines()))
                            sss = '{' + sss + '}'
                            
                            if ss != "": ss += ", "
                            ss += sss
                        for iii in range(i + 1, ii):
                            for jj in range(j, k+1):
                                cell[iii][jj] = -1
                        # if ww_2 == "BIT DATA":
                        #     print("j = " + str(j) + "  k = " + str(k))
                        #     my_file.write(ss + "EEEEEEEEEEEEEEEEEEEEEEEEEEEEE")
                    elif ww_2 in ["PIPE DATA"]: 
                        # Multiple Row Header
                        jj_count = 0
                        base_j = j
                        for jj in range(j + 1, k + 1):
                            if cell[i+1][jj] == cell[i+1][jj-1]: continue
                            jj_count += 1
                            if jj_count % 2 == 0:
                                base_j = jj
                                continue

                            sss = ""
                            for iii in range(i + 1, ii): 
                                if sss != "" : sss += ", "
                                sss += '"' + remove_special_characters(' '.join(word[cell[iii][base_j]].splitlines())) + '": "' + remove_special_characters(' '.join(word[cell[iii][jj]].splitlines())) + '"'
                                # print(str(iii) + "  " + str(jj) + "  " + " ".join(word[cell[iii][j]].splitlines()) + ": " +" ".join(word[cell[iii][jj]].splitlines()))
                            sss = '{' + sss + '}'
                            
                            if ss != "": ss += ", "
                            ss += sss
                        for iii in range(i +1, ii):
                            for jj in range(j, k+1):
                                cell[iii][jj] = -1

                    elif ww_2 in ["Well Information", "Daily Operations Information", "Management Summary", "Mud Checks", "Mud Volume Summary", "String No. 3", "String No. 3", "String No. 3  - Bit Information", "Wellbore Information", "Deviation Survey (Tie point) -  TABU-B19 Actual Survey"]:
                        print("www = " + ww_2) 
                        pre_cell = -2
                        ss += "{\n"
                        for iii in range(i + 1, ii): 
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if sss != "" and word[cell[iii][jj]] != "": sss += ", "
                                    ww = " ".join(word[cell[iii][jj]].splitlines()) 
                                    if len(word[cell[iii][jj]].splitlines()) > 0 :
                                        sss += '"' + remove_special_characters(word[cell[iii][jj]].splitlines()[0]) + '": "'
                                        if len(word[cell[iii][jj]].splitlines()) > 1 : 
                                            sss += remove_special_characters(word[cell[iii][jj]].splitlines()[1])
                                        sss += '"'
                                    pre_cell = cell[iii][jj]
                                cell[iii][jj] = -1
                            if sss != "":
                                if ss != "{\n": ss += ", "
                                ss += sss
                        ss += "}"

                        
                    else: # etc group
                        print("ww_s = " + ww_2)
                        pre_cell = -2
                        for iii in range(i + 1, ii): 
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if sss != "" and word[cell[iii][jj]] != "": sss += ", "
                                    ww = " ".join(word[cell[iii][jj]].splitlines()) 
                                    sss += word[cell[iii][jj]]
                                    pre_cell = cell[iii][jj]
                                cell[iii][jj] = -1
                            if sss != "":
                                sss = '{' + sss + '}'
                                
                                if ss != "": ss += ", "
                                ss += sss
            # Output to output file
            if ww_2 != "": 
                ww_2 = remove_special_characters(ww_2)
                if write_started:
                    # my_file.write(", \n")
                    write_into_file(", \n")
                # if ww_2.__contains__("ype Time"):
                #     # Remove special characters
                #     ww_2 = ww_2.replace('"', '')
                if ss != "" and ss[0] == "{": 
                    #ww_2 = '"' + ww_2 + '": {\n'
                    ww_2 = '"' + ww_2 + '": [\n'
                elif ww_2.find(":") > -1:
                    # ww_2 = '"' + ww_2.split(":", 1)[0] + '": "' + ww_2.split(":", 1)[1] + '"'
                    ww_2 = '"' + remove_special_characters(ww_2.split(":", 1)[0]) + '": "' + ww_2.split(":", 1)[1] + '"'
                elif len(ww_2.splitlines()) > 1:
                    # ww_2 = '"' + ww_2.splitlines()[0] + '": "' + "  ".join(ww_2.splitlines()[1:]) + '"'
                    ww_2 = '"' + remove_special_characters(ww_2.splitlines()[0]) + '": "' + "  ".join(ww_2.splitlines()[1:]) + '"'
                else:
                    # Let's remove special characters if it's header
                    ww_2 = remove_special_characters(ww_2)
                    ww_2 = '"' + ww_2 + '": ""'

                # my_file.write(ww_2)
                if ww_2.__contains__("Currentstatus"):
                    ww_2 = ww_2.replace('6"', '6\\"').replace('5.5"', '5.5\\"').replace('10"', '10\\"').replace('2"', '2\\"').strip()

                write_into_file(ww_2)
                write_started = True

            # if ss != "": my_file.write(ss)
            if ss != "" and ww_2.__contains__("OPERATION SUMMARY"):
                ss = ss.replace('6"', '6\\"').replace('5.5"', '5.5\\"').replace('10"', '10\\"').replace('2"', '2\\"').strip()

            if ss != "": write_into_file(ss)

            # if ww_2 != "" and ss != "" and ss[0] == "{": my_file.write('\n}')
            if ww_2 != "" and ss != "" and ss[0] == "{": write_into_file('\n]')
            for jj in range(j, k+1):
                cell[i][jj] = -1

    # my_file.write("\n")
# my_file.write("}")
write_into_file("\n")
write_into_file("}")
my_file.close()
