import re
import openpyxl
from typing import List
from os import listdir
from os.path import isfile, join
from datetime import date, timedelta

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

        # Find a string after startswith
        self.id_ = self.project_file_id = re.findall(r'\((.*?)\)', path_name)[0]

        # Define variable to load the dataframe
        wb1 = openpyxl.load_workbook(path_name)

        # Define variable to read sheet
        ws1 = wb1.active

        offset = 1
        # Iterate the loop to read the cell values
        for row_idx in range(offset, ws1.max_row):
            if ws1.cell(row=row_idx, column=2).value == "121":
                self.expenses_internal = int(ws1.cell(row=row_idx, column=7).value)
                #print("내부인건비", ws1.cell(row=row_idx, column=7).value)
            elif ws1.cell(row=row_idx, column=2).value == "201":
                self.expenses_external = int(ws1.cell(row=row_idx, column=7).value)
                #print("계약직내부인건비", ws1.cell(row=row_idx, column=7).value)
            elif ws1.cell(row=row_idx, column=2).value == "123":
                self.expenses_overhead = int(ws1.cell(row=row_idx, column=7).value)
                #print("간접비", ws1.cell(row=row_idx, column=7).value)

        wb1.close()

    def show(self):
        print(self.path_name)
        print("과제번호(파일): ", self.project_file_id)
        print("과제번호: ", self.project_base_id)
        print("계정번호: ", self.project_cur_id)
        print("과제시작일: ", self.date_begin)
        print("과제종료일: ", self.date_end)
        print("월수/월수({0})".format(TargetYear), self.months, self.months_target)
        print("내부인건비: ", "{:,}".format(self.expenses_internal))
        print("계약직내부인건비: ", "{:,}".format(self.expenses_external))
        print("간접비: ", "{:,}".format(self.expenses_overhead))
        print("내부인건비({0})".format(TargetYear), "{:,}".format(self.expenses_internal_target))
        print("계약직내부인건비({0})".format(TargetYear), "{:,}".format(self.expenses_external_target))
        print("간접비({0})".format(TargetYear), "{:,}".format(self.expenses_overhead_target))

class Contract:
    employee_id = ""    # 사번
    name = ""           # 성명
    project_id = ""    # 계정번호
    date_start = date(2000, 1, 1)    # 발령 시작일
    date_end = date(2000, 1, 1)      # 발령 종료일
    working_time = 0  # 시간
    expenses_per_hour = 0  # 시간당 단가
    expenses_holiday = 0   # 주유수당
    expenses_monthly = 0   # 월정액
    expenses_extra = 0   # 추가부담

    def __init__(self, id, name, project_id, date_start, date_end, working_time, expenses_per_hour, expenses_holiday, expenses_monthly, expenses_extra):
        self.id_ = id + name + project_id + str(date_start) + str(date_end)
        self.employee_id = id
        self.name = name
        self.project_id = project_id
        self.date_start = date_start
        self.date_end = date_end
        self.working_time = int(working_time)
        self.expenses_per_hour = int(expenses_per_hour)
        if expenses_holiday == None:
            self.expenses_holiday = 0
        else:
            self.expenses_holiday = int(expenses_holiday)
        self.expenses_monthly = int(expenses_monthly)
        if expenses_extra == None:
            self.expenses_extra = 0
        else:
            self.expenses_extra = int(expenses_extra)

    def show(self):
        print("Name: ", self.name, "ID: ", self.employee_id, "Project ID: ", self.project_id)
        print("Date Start: ", self.date_start, "Date End: ", self.date_end)
        print("Working Time: ", self.working_time, "Expenses per hour: ", "{:,}".format(self.expenses_per_hour), "Expenses Holiday: ", "{:,}".format(self.expenses_holiday), "Expenses Monthly: ", "{:,}".format(self.expenses_monthly), "Expenses Extra: ", "{:,}".format(self.expenses_extra))
        print("-------------------------------------")

class Employee:
    name = ""  # 성명
    contracts: List[Contract] = []  # 계약

    total_salary = 0  # 연봉
    total_expenses = 0  # 지출

    def __init__(self, id, name):
        self.id_ = id
        self.name = name

    def add(self, contract):
        self.contracts.append(contract)

        self.total_salary = self.total_salary + contract.expenses_monthly
        self.total_expenses = self.total_expenses + contract.expenses_monthly + contract.expenses_extra

    def show(self):
        print("Name: ", self.name, "ID: ", self.id_, "Total Salary: ", "{:,}".format(self.total_salary), "Total Expenses: ", "{:,}".format(self.total_expenses))
        print("-------------------------------------")
        #for contract in self.contracts:
        #    contract.show()
        print("=====================================")

