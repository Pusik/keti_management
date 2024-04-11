import re
import openpyxl
from os import listdir
from os.path import isfile, join
#from ProjectBudget import ProjectBudget


class ProjectBudget:
    project_file_id = ""    # 계정번호(파일명)
    project_cur_id = ""     # 계정번호
    project_base_id = ""    # 과제번호
    expenses_internal = 0   # 내부인건비
    expenses_external = 0   # 계약직내부인건비
    expenses_overhead = 0   # 간접비
    date_begin = ""         # 과제시작일
    date_end = ""           # 과제종료일

    def __init__(self, path_name):
        self.path_name = path_name
        #print(path_name)

        # Define variable to load the dataframe
        wb = openpyxl.load_workbook(path_name)

        # Define variable to read sheet
        ws = wb.active

        # Iterate the loop to read the cell values
        for row_idx in range(0, ws.max_row):
            if ws.cell(row=row_idx+1, column=2).value == "121":
                self.expenses_internal = int(ws.cell(row=row_idx+1, column=7).value)
                #print("내부인건비", ws.cell(row=row_idx+1, column=7).value)
            elif ws.cell(row=row_idx+1, column=2).value == "201":
                self.expenses_external = int(ws.cell(row=row_idx+1, column=7).value)
                #print("계약직내부인건비", ws.cell(row=row_idx+1, column=7).value)
            elif ws.cell(row=row_idx+1, column=2).value == "123":
                self.expenses_overhead = int(ws.cell(row=row_idx+1, column=7).value)
                #print("간접비", ws.cell(row=row_idx+1, column=7).value)

    def show_path(self):
        print(self.path_name)
        print("과제번호(파일): ", self.project_file_id)
        print("과제번호: ", self.project_base_id)
        print("계정번호: ", self.project_cur_id)
        print("과제시작일: ", self.date_begin)
        print("과제종료일: ", self.date_end)
        print("내부인건비: ", self.expenses_internal)
        print("계약직내부인건비: ", self.expenses_external)
        print("간접비: ", self.expenses_overhead)

# Define the working folder
WorkingFolder = "C:/Users/ParkPusik/OneDrive - OPENLAB/MOBILITY/____센터장App/20231231"

FileBudget = "연구_예산관리_예실대비표(4P4C3369)_Button(엑셀).xlsx"
FileIrregularEmployee = "인사_비상근전문계약직_과제참여현황_Button(엑셀다운).xlsx"
WholeProjectInfo = "연구_과제현황_과제종합정보_Button(엑셀데이터다운).xlsx"

# List all files in the folder
Files = [f for f in listdir(WorkingFolder) if isfile(join(WorkingFolder, f))]
Projects = []
for f in Files:
    if f.startswith("연구_예산관리_예실대비표("):
        ##print(f)
        
        proj = ProjectBudget(WorkingFolder + "/" + f)
        # Find a string after startswith
        proj.project_file_id = re.findall(r'\((.*?)\)', f)[0]

        #proj.show_path()
        Projects.append(proj)

        #WorkingFiles.append(f)

#print(WorkingFiles)

# Define the beginning/ending date of the project
wb2 = openpyxl.load_workbook(WorkingFolder + "/" + WholeProjectInfo)
ws2 = wb2.active

for proj in Projects:
    for row_idx in range(0, ws2.max_row):
        #print(ws2.cell(row=row_idx+1, column=9).value, proj.project_cur_id)

        if ws2.cell(row=row_idx+1, column=9).value == proj.project_file_id:
            proj.project_cur_id = proj.project_file_id
            #print("과제시작일", ws2.cell(row=row_idx+1, column=5).value)
            #print("과제종료일", ws2.cell(row=row_idx+1, column=6).value)
            proj.date_begin = ws2.cell(row=row_idx+1, column=5).value
            proj.date_end = ws2.cell(row=row_idx+1, column=6).value
            # 과제번호
            proj.project_base_id = ws2.cell(row=row_idx+1, column=1).value
            break
    # No match
    #print("No match for ", proj.project_cur_id)

    for row_idx in range(0, ws2.max_row):
        #print(ws2.cell(row=row_idx+1, column=1).value, proj.project_cur_id)

        if ws2.cell(row=row_idx+1, column=1).value == proj.project_file_id:
            proj.project_cur_id = ""
            #print("과제시작일", ws2.cell(row=row_idx+1, column=5).value)
            #print("과제종료일", ws2.cell(row=row_idx+1, column=6).value)
            proj.date_begin = ws2.cell(row=row_idx+1, column=5).value
            proj.date_end = ws2.cell(row=row_idx+1, column=6).value
            # 과제번호
            proj.project_base_id = ws2.cell(row=row_idx+1, column=1).value
            break

# Duplication check
dup = 0
for proj1 in Projects:
    for proj2 in Projects:
        if proj1.project_base_id == proj2.project_file_id:
            dup += 1
            break

# Summary of projects
for proj in Projects:
    proj.show_path()
print("Total", len(Projects), "Duplication: ", dup)



