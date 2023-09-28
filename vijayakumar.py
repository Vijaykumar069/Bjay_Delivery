import openpyxl
from datetime import datetime, timedelta


def readFile(fileName):
    #Reads an file and returns a list of lists, where each sublist represents a row in the file.
    workbook = openpyxl.load_workbook(fileName)
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(list(row))
    return data


def dateTime(dateTime):
    try:
        return datetime.strptime(dateTime, '%Y-%m-%d %H:%M:%S')
    except ValueError:
        return None


def employees(employees):
    #Prints the name and position of the given employees.
    for employee in employees:
        print(employee[7], employee[0])

def employee_whoworked_7consecutivedays(data):
    # Returns a list of employees who have worked for 7 consecutive days.
    employees = []
    for i in range(len(data) - 6):
        if all(data[i + j][5] == data[i][5] for j in range(7)):
            employees.append(data[i])
    return employees


def employee_withlessthan_10hours_and_greaterthan_1hour(data):
    """Returns a list of employees who have less than 10 hours of time between shifts but greater than 1 hour."""
    employees = []
    for i in range(len(data) - 1):
        shift_end = dateTime(data[i][3])
        shift_start = dateTime(data[i + 1][2])
        if shift_end is not None and shift_start is not None:
            time_diff_hours = (shift_start - shift_end).total_seconds() / 3600
            if data[i + 1][5] == data[i][5] and 1 < time_diff_hours < 10:
                employees.append(data[i])
    return employees


def employee_whoworked_morethan_14hours_in_a_singleshift(data):
    """Returns a list of employees who have worked for more than 14 hours in a single shift."""
    employees = []
    for row in data:
        shift_start = dateTime(row[2])
        shift_end = dateTime(row[3])
        if shift_start is not None and shift_end is not None:
            shift_duration_hours = (shift_end - shift_start).total_seconds() / 3600
            if shift_duration_hours > 14:
                employees.append(row)
    return employees


def mainFunction():
    # Reads the file and prints the name and position of employees who satisfy the given conditions.
    fileName = "v.xlsx"
    data = readFile(fileName)
    first_task = employee_whoworked_7consecutivedays(data)
    print("Employees who worked for 7 consecutive days : ")
    employees(first_task)
    second_task= employee_withlessthan_10hours_and_greaterthan_1hour(data)
    print("Employees who have less than 10 hours of time between shifts but greater than 1 hour :\n")
    employees( second_task)
    third_task = employee_whoworked_morethan_14hours_in_a_singleshift(data)
    print("\nEmployees who have worked for more than 14 hours in a single shift : ")
    employees(third_task)


mainFunction()
