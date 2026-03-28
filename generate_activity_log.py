import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
import os

print("Reading employee data...")
employees = pd.read_csv('data/employee_data.csv')
real_emp_ids = employees['EmpID'].tolist()

# Ghost users — NOT in employee list
fake_ids = [99001, 99002, 99003, 99004, 99005,
            99006, 99007, 99008, 99009, 99010]

all_user_ids = real_emp_ids + fake_ids

actions = ['login', 'logout', 'file_access', 'email_sent',
           'report_generated', 'system_update', 'data_export', 'meeting_joined']

departments = ['Sales', 'Production', 'Finance & Accounting',
               'HR', 'IT', 'Operations', 'Marketing', 'Legal']

n = 500000
print(f"Generating {n:,} rows of activity log...")
np.random.seed(42)

start_date = datetime(2024, 1, 1)

activity_data = {
    'user_id': np.random.choice(all_user_ids, n),
    'login_time': [start_date + timedelta(
                    days=int(d), hours=int(h), minutes=int(m))
                   for d, h, m in zip(
                       np.random.randint(0, 365, n),
                       np.random.randint(0, 24, n),
                       np.random.randint(0, 60, n))],
    'action': np.random.choice(actions, n),
    'department': np.random.choice(departments, n),
    'session_duration_mins': np.random.randint(1, 480, n),
    'ip_address': [f"192.168.{random.randint(1,255)}.{random.randint(1,255)}"
                   for _ in range(n)]
}

activity_df = pd.DataFrame(activity_data)
activity_df['login_time'] = pd.to_datetime(activity_df['login_time'])

# Active employee list — only real employees (no fake IDs)
employee_df = employees[['EmpID', 'FirstName', 'LastName',
                          'DepartmentType', 'EmployeeStatus']].copy()
employee_df.columns = ['emp_id', 'first_name', 'last_name',
                        'department', 'status']

os.makedirs('data', exist_ok=True)

print("Saving both sheets into ONE Excel file...")
with pd.ExcelWriter('data/employee_activity.xlsx') as writer:
    activity_df.to_excel(writer, sheet_name='Activity_Log', index=False)
    employee_df.to_excel(writer, sheet_name='Active_Employees', index=False)

size_mb = os.path.getsize('data/employee_activity.xlsx') / (1024 * 1024)
print(f"\nDone!")
print(f"File: data/employee_activity.xlsx")
print(f"Size: {size_mb:.1f} MB")
print(f"Sheet 1 - Activity_Log    : {len(activity_df):,} rows")
print(f"Sheet 2 - Active_Employees: {len(employee_df):,} rows")
print(f"Ghost users hidden inside : {len(fake_ids)}")