
import json
import zipfile
import xlsxwriter
from os import walk

#Create the Excel
workbook = xlsxwriter.Workbook('results.xlsx')			
worksheet = workbook.add_worksheet()

#Set the columns widths
worksheet.set_column('A:F', 20)

#Header
bold = workbook.add_format({'bold': 1})
worksheet.write('A1', 'Project', bold)
worksheet.write('B1', 'Project id', bold)
worksheet.write('C1', 'Project date', bold)
worksheet.write('D1', 'Project hour', bold)
worksheet.write('E1', 'SpriteNames', bold)
worksheet.write('F1', 'TotalSprites', bold)


#Analysis of the number of sprites

folder_path = "path/to/the/projects"
row = 1

for (path, ficheros, archivos) in walk(folder_path):
			for filename in archivos:
					if filename != "last.json":
							print filename
							project = path + "/" + filename
							sprites = []
							try:
									zip_file = zipfile.ZipFile(project, "r")
									json_project = json.loads(zip_file.open("project.json").read())
  
									for sprite_info in json_project['children']:
											try:
													sprite_name = sprite_info['objName']
													sprites.append(sprite_name)
											except:
													pass	

							except:
									pass


							#________Print results in Excel_________#

							#Project name, complete
							project_name = project.split("/")[-1]
							worksheet.write(row, 0, project_name) 

							#Project_id
							project_id = project_name.split("_")[0]
							worksheet.write(row, 1, project_id)

							#Date
							p = project_name.split("_")[1:]
							date = p[0].split("-")[0] + "-" + p[0].split("-")[1] + "-" + p[0].split("-")[2]
							worksheet.write(row, 2, date)

							#Hour
							hour = p[0].split("-")[3] + ":" + p[0].split("-")[4]
							worksheet.write(row, 3, hour)

							#Sprite Names
							worksheet.write(row, 4, str(sprites))
							worksheet.write(row, 5, len(sprites))

							#New row
							row += 1


workbook.close()
							
							








