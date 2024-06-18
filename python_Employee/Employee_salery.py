import pandas as pd


#  Excel file and sheet name
employee_df = pd.read_excel('employee_data.xlsx', sheet_name='Employee Data')
timesheet_df = pd.read_excel('employee_data.xlsx', sheet_name='Timesheet')

def get_employee_times(employee_id):
    return timesheet_df[timesheet_df['Employee ID'] == employee_id]

def calculate_lop(timesheet):
    late_or_early_count = 0
    for index, row in timesheet.iterrows():
        hours_worked = row['Hours Worked']
        if hours_worked < 9:
            late_or_early_count += 1
    lop = (late_or_early_count // 3) * 0.5
    return lop

def calculate_payroll(employee_df, timesheet_df):
    payroll = []

    for index, employee in employee_df.iterrows():
        employee_id = employee['Employee ID']
        name = employee['Employee Name']
        salary = employee['Salary']
        emp_type = employee['Employee Type']
        
        if emp_type != 'Management':
            timesheet = get_employee_times(employee_id)
            lop = calculate_lop(timesheet)
            total_salary = salary - (salary / 30) * lop

            payroll.append({
                'Employee ID': employee_id,
                'Employee Name': name,
                'Employee Type': emp_type,
                'Salary': salary,
                'LOP': lop,
                'Total Salary': total_salary
            })

    return payroll

# Calculate payroll

payroll = calculate_payroll(employee_df, timesheet_df)

# Convert to DataFrame

payroll_df = pd.DataFrame(payroll)
print(payroll_df)

#  new Excel file
payroll_df.to_excel('abc.xlsx',index=False)