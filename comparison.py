import pandas as pd # type: ignore
import math
from datetime import date
from openpyxl import load_workbook # type: ignore
from openpyxl.styles import Border, PatternFill, Side # type: ignore
from openpyxl.styles import Font # type: ignore

# yes = input("Enter Yesterday Date: ")
yes = '09'
yesLocation = "//192.168.1.231/Planning Internal/Capacity planning/Capacity Report/2024/12. Dec/" + str(yes) + "/"

# today = input("Enter Today's Date: ")
today = 10
todLocation = "//192.168.1.231/Planning Internal/Capacity planning/Capacity Report/2024/12. Dec/" + str(today) + "/"

yes_comparison = pd.read_csv(yesLocation + 'Unit wise Buyer wise Plan Qty.csv')
tod_comparison = pd.read_csv(todLocation + 'Unit wise Buyer wise Plan Qty.csv')

comparison = pd.DataFrame()
comparison['Units'] = None
comparison['Yesterday Qty.'] = None
comparison['Today Qty.'] = None
cnt = 0
for index, row in yes_comparison.iterrows():
    if row['Factory+Buyer'] == '-' and cnt < 6:
        comparison.loc[cnt, 'Units'] = row['Pl. Board']
        comparison.loc[cnt, 'Yesterday Qty.'] = row[2]
        cnt += 1
cnt = 0
for index, row in tod_comparison.iterrows():
    if row['Factory+Buyer'] == '-' and cnt < 6:
        comparison.loc[cnt, 'Today Qty.'] = row[2]
        cnt += 1

comparison['Change'] =comparison['Today Qty.'] - comparison['Yesterday Qty.']
print(comparison)

print("\nSuccessfully done :)")
from datetime import datetime
import calendar
current_month = datetime.now().month
month_names = list(calendar.month_name)
current_month = month_names[current_month]
