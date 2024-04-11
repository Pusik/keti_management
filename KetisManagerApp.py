import re
import openpyxl
from os import listdir
from os.path import isfile, join
from datetime import date, timedelta
#from ProjectBudget import ProjectBudget

TargetYear = 2023
TargetYearBegin = date(TargetYear, 1, 1)
TargetYearEnd = date(TargetYear, 12, 31)


# Define the working folder
WorkingFolder = "C:/Users/ParkPusik/OneDrive - OPENLAB/MOBILITY/____센터장App/20231231"

FileBudget = "연구_예산관리_예실대비표(4P4C3369)_Button(엑셀).xlsx"
FileIrregularEmployee = "인사_비상근전문계약직_과제참여현황_Button(엑셀다운).xlsx"
WholeProjectInfo = "연구_과제현황_과제종합정보_Button(엑셀데이터다운).xlsx"

class ProjectBudget:
    project_file_id = ""    # 계정번호(파일명)
    project_cur_id = ""     # 계정번호
    project_base_id = ""    # 과제번호
    expenses_internal = 0   # 내부인건비
    expenses_external = 0   # 계약직내부인건비
    expenses_overhead = 0   # 간접비
    expenses_internal_target = 0   # 내부인건비
    expenses_external_target = 0   # 계약직내부인건비
    expenses_overhead_target = 0   # 간접비
    date_begin = date(2000, 1, 1)     # 과제시작일
    date_end = date(2000, 1, 1)       # 과제종료일
    months = 0
    monthly_target = 0

    def __init__(self, path_name):
        self.path_name = path_name
        #print(path_name)

        # Define variable to load the dataframe
        wb1 = openpyxl.load_workbook(path_name)

        # Define variable to read sheet
        ws1 = wb1.active

        # Iterate the loop to read the cell values
        for row_idx in range(0, ws1.max_row):
            if ws1.cell(row=row_idx+1, column=2).value == "121":
                self.expenses_internal = int(ws1.cell(row=row_idx+1, column=7).value)
                #print("내부인건비", ws1.cell(row=row_idx+1, column=7).value)
            elif ws1.cell(row=row_idx+1, column=2).value == "201":
                self.expenses_external = int(ws1.cell(row=row_idx+1, column=7).value)
                #print("계약직내부인건비", ws1.cell(row=row_idx+1, column=7).value)
            elif ws1.cell(row=row_idx+1, column=2).value == "123":
                self.expenses_overhead = int(ws1.cell(row=row_idx+1, column=7).value)
                #print("간접비", ws1.cell(row=row_idx+1, column=7).value)

        wb1.close()

    def show(self):
        print(self.path_name)
        print("과제번호(파일): ", self.project_file_id)
        print("과제번호: ", self.project_base_id)
        print("계정번호: ", self.project_cur_id)
        print("과제시작일: ", self.date_begin)
        print("과제종료일: ", self.date_end)
        print("월수/월수({TargetYear}): ", self.months, self.months_target)
        print("내부인건비: ", "{:,}".format(self.expenses_internal))
        print("계약직내부인건비: ", "{:,}".format(self.expenses_external))
        print("간접비: ", "{:,}".format(self.expenses_overhead))
        print("내부인건비({TargetYear}): ", "{:,}".format(self.expenses_internal_target))
        print("계약직내부인건비({TargetYear}): ", "{:,}".format(self.expenses_external_target))
        print("간접비({TargetYear}): ", "{:,}".format(self.expenses_overhead_target))

# List all files in the folder
Files = [f for f in listdir(WorkingFolder) if isfile(join(WorkingFolder, f))]
Projects = []
for f in Files:
    if f.startswith("연구_예산관리_예실대비표("):
        ##print(f)
        
        proj = ProjectBudget(WorkingFolder + "/" + f)
        # Find a string after startswith
        proj.project_file_id = re.findall(r'\((.*?)\)', f)[0]

        Projects.append(proj)

# Define the beginning/ending date of the project
wb2 = openpyxl.load_workbook(WorkingFolder + "/" + WholeProjectInfo)
ws2 = wb2.active

for proj in Projects:
    for row_idx in range(0, ws2.max_row):
        #print(ws2.cell(row=row_idx+1, column=9).value, proj.project_cur_id)

        if ws2.cell(row=row_idx+1, column=9).value == proj.project_file_id:
            proj.project_cur_id = proj.project_file_id
            #print("과제시작일", ws2.cell(row=row_idx+1, column=11).value)
            #print("과제종료일", ws2.cell(row=row_idx+1, column=12).value)
            str_date = ws2.cell(row=row_idx+1, column=11).value
            str_date = str_date.replace(".", "")
            proj.date_begin = date(int(str_date[0:4]), int(str_date[4:6]), int(str_date[6:8]))

            str_date = ws2.cell(row=row_idx+1, column=12).value
            str_date = str_date.replace(".", "")
            proj.date_end = date(int(str_date[0:4]), int(str_date[4:6]), int(str_date[6:8]))

            # 과제번호
            proj.project_base_id = ws2.cell(row=row_idx+1, column=1).value
            break
    # No match
    #print("No match for ", proj.project_cur_id)

    for row_idx in range(0, ws2.max_row):
        #print(ws2.cell(row=row_idx+1, column=1).value, proj.project_cur_id)

        if ws2.cell(row=row_idx+1, column=1).value == proj.project_file_id:
            proj.project_cur_id = ""
            #print("과제시작일", ws2.cell(row=row_idx+1, column=11).value)
            #print("과제종료일", ws2.cell(row=row_idx+1, column=12).value)
            str_date = ws2.cell(row=row_idx+1, column=11).value
            str_date = str_date.replace(".", "")
            proj.date_begin = date(int(str_date[0:4]), int(str_date[4:6]), int(str_date[6:8]))

            str_date = ws2.cell(row=row_idx+1, column=12).value
            str_date = str_date.replace(".", "")
            proj.date_end = date(int(str_date[0:4]), int(str_date[4:6]), int(str_date[6:8]))

            # 과제번호
            proj.project_base_id = ws2.cell(row=row_idx+1, column=1).value
            break

for proj in Projects:
    if proj.date_end < TargetYearBegin or proj.date_begin > TargetYearEnd:
        #print("Out of range", proj.project_base_id, proj.date_begin, proj.date_end)
        Projects.remove(proj)

for proj in Projects:
    # Calculate the month from the date_begin to date_end
    diff = proj.date_end - proj.date_begin
    proj.months = int(diff.days / 30)
    #print("Total months: ", proj.months)
    if proj.months > 12:
        print("Long project")

    monthly_expenses_internal = int(proj.expenses_internal / proj.months)
    monthly_expenses_external = int(proj.expenses_external / proj.months)
    monthly_expenses_overhead = int(proj.expenses_overhead / proj.months)

    if proj.date_begin <= TargetYearBegin and proj.date_end <= TargetYearEnd:
        diff = proj.date_end - TargetYearBegin
        proj.months_target = int(diff.days / 30)
    elif proj.date_begin <= TargetYearBegin and proj.date_end >= TargetYearEnd:  
        proj.months_target = 12
    elif proj.date_begin >= TargetYearBegin and proj.date_end >= TargetYearEnd:
        diff = TargetYearEnd - proj.date_begin
        proj.months_target = int(diff.days / 30)   
    elif proj.date_begin >= TargetYearBegin and proj.date_end <= TargetYearEnd:
        diff = proj.date_end - proj.date_begin
        proj.months_target = int(diff.days / 30)
    else:
        print("Error", proj.project_cur_id, proj.date_begin, proj.date_end)

    proj.expenses_internal_target = monthly_expenses_internal * proj.months_target
    proj.expenses_external_target = monthly_expenses_external * proj.months_target
    proj.expenses_overhead_target = monthly_expenses_overhead * proj.months_target


#for proj in Projects:
#    proj.show()

# Duplication check
#dup = 0
#for proj1 in Projects:
#    for proj2 in Projects:
#        if proj1.project_cur_id == proj2.project_file_id:
#            dup += 1
#            break
#print("Total", len(Projects), "Duplication: ", dup)

# Summary of projects
total_expenses_internal_target = 0
total_expenses_external_target = 0
total_expenses_overhead_target = 0

for proj in Projects:
    # Acculmulate the expenses and overhead
    total_expenses_internal_target += proj.expenses_internal_target
    total_expenses_external_target += proj.expenses_external_target
    total_expenses_overhead_target += proj.expenses_overhead_target

# Format comma separated number
print("Total 내부인건비: ", "{:,}".format(total_expenses_internal_target))
print("Total 계약직내부인건비: ", "{:,}".format(total_expenses_external_target))
print("Total 간접비: ", "{:,}".format(total_expenses_overhead_target))
print("Total OH", "{:,}".format(total_expenses_internal_target + total_expenses_overhead_target))


# Write Projects into a excel file
wb3 = openpyxl.Workbook()
ws3 = wb3.active

ws3.append(["과제번호(파일)", "과제번호", "계정번호", "과제시작일", "과제종료일", "월수", "월수({TargetYear})", "내부인건비", "계약직내부인건비", "간접비", "내부인건비({TargetYear})", "계약직내부인건비({TargetYear})", "간접비({TargetYear})"])
for proj in Projects:
    ws3.append([proj.project_file_id, proj.project_base_id, proj.project_cur_id, proj.date_begin, proj.date_end, proj.months, proj.months_target, proj.expenses_internal, proj.expenses_external, proj.expenses_overhead, proj.expenses_internal_target, proj.expenses_external_target, proj.expenses_overhead_target])   

wb3.save(WorkingFolder + "/Summary_Projects" + str(TargetYear) + ".xlsx")
print("Summary_Projects" + str(TargetYear) + ".xlsx is saved")

wb2.close()
wb3.close()
