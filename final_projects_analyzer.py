"""
import xlrd
from xlrd.sheet import ctype_text


workbook = xlrd.open_workbook("results_testing.xlsx")
sheet = workbook.sheet_by_index(0)


dates = sheet.col(2)
id_projects = sheet.col(1)
id_projects_list = []

for idx, cell_obj in enumerate(id_projects):
    if idx != 0:
      id_projects_list.append(cell_obj.value)


id_projects_list.sort(reverse=True)

"""

import pandas as pd

xl = pd.ExcelFile("Results/evolution_time_results.xlsx")
df = xl.parse("Sheet1")

df = df.set_index(['Project id', 
                   'Project date', 
                   'Project hour'], drop=False)

last_projects = df.groupby(['Project id', 'Project date'], as_index=False)['Project hour'].max()

last_project = last_projects.groupby(['Project id'], as_index=False)['Project date'].max()

index_last_project = last_project.set_index(['Project id', 'Project date'])

index_last_projects = last_projects.set_index(['Project id', 
                                               'Project date'], drop=False)

final_last_project = index_last_projects.loc[index_last_project.index.tolist()]

test = final_last_project.set_index(['Project id', 
                   'Project date', 
                   'Project hour'])

result = df.loc[test.index.tolist()]


writer = pd.ExcelWriter('Results/final_projects_results.xlsx')
result.to_excel(writer, sheet_name='Sheet1', index=False, engine='xlsxwriter')
writer.save()

