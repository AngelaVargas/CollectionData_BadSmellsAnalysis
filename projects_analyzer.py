import json
import sys
import os
import zipfile
import xlsxwriter
import ast


class Mastery: 

  """Analyzer of projects sb2, and create an Excel with solutions"""
  def __init__(self):

    self.total_projects = []


  """Start the analysis."""
  def process(self,filename):

   zip_file = zipfile.ZipFile(filename, "r")

   for i in zip_file.namelist():
      json_project = zip_file.extract(i)
      self.total_projects.append(json_project)

   
 
  def analyze(self):

    workbook = xlsxwriter.Workbook('results.xlsx')			#Creo el Excel donde guardar los resultados
    worksheet = workbook.add_worksheet()

    # Set the columns widths.
    worksheet.set_column('A:K', 15)
    worksheet.set_column('L:Q', 20)

    #Header
    bold = workbook.add_format({'bold': 1})
    worksheet.write('A1', 'Project', bold)
    worksheet.write('B1', 'Abstraction', bold)
    worksheet.write('C1', 'Parallelization', bold)
    worksheet.write('D1', 'Logic', bold)
    worksheet.write('E1', 'Synchronization', bold)
    worksheet.write('F1', 'FlowControl', bold)
    worksheet.write('G1', 'UserInteractivity', bold)
    worksheet.write('H1', 'DataRepresent.', bold)
    worksheet.write('I1', 'TotalMastery', bold)
    worksheet.write('J1', 'AverageMastery', bold)
    worksheet.write('K1', 'Level', bold)
    worksheet.write('L1', 'DefaultSpriteNames', bold)
    worksheet.write('M1', 'TotalDefaultNames', bold)
    worksheet.write('N1', 'DuplicateScripts', bold)
    worksheet.write('O1', 'TotalDupScripts', bold)
    worksheet.write('P1', 'DeadCode', bold)
    worksheet.write('Q1', 'AttributeInitialization', bold)

    row = 1

    for file_name in self.total_projects:

      #Request to hairball
      metricMastery = "hairball -p mastery.Mastery " + file_name
      metricSpriteNaming = "hairball -p convention.SpriteNaming " + file_name
      metricDuplicateScript = "hairball -p duplicate.DuplicateScripts " + file_name
      metricDeadCode = "hairball -p blocks.DeadCode " + file_name
      metricInitialization = "hairball -p initialization.AttributeInitialization " + file_name

      #Response from hairball
      resultMastery = os.popen(metricMastery).read()
      resultSpriteNaming = os.popen(metricSpriteNaming).read()
      resultDuplicateScript = os.popen(metricDuplicateScript).read()
      resultDeadCode = os.popen(metricDeadCode).read()
      resultInitialization = os.popen(metricInitialization).read()

      
      #Print results in Excel
      project_name = file_name.split("/")[-1]
      worksheet.write(row, 0, project_name) 

      lines = resultMastery.split("\n") 
      dic = ast.literal_eval(lines[1])
      worksheet.write(row, 1, dic["Abstraction"])
      worksheet.write(row, 2, dic["Parallelization"])
      worksheet.write(row, 3, dic["Logic"])
      worksheet.write(row, 4, dic["Synchronization"])
      worksheet.write(row, 5, dic["FlowControl"])
      worksheet.write(row, 6, dic["UserInteractivity"])
      worksheet.write(row, 7, dic["DataRepresentation"])

      total_mastery = lines[2].split(":")[1]
      worksheet.write(row, 8, total_mastery)

      average_mastery = lines[3].split(":")[1]
      worksheet.write(row, 9, average_mastery)

      level = lines[4].split(":")[1]
      worksheet.write(row, 10, level)

      lines = resultSpriteNaming.split("\n")
      sprite_names = lines[2]
      worksheet.write(row, 11, sprite_names)

      total_default_names = lines[1].split(" ")[0]
      worksheet.write(row, 12, total_default_names)

      lines = resultDuplicateScript.split("\n")

      try:
       duplicate_scripts = lines[2]
       worksheet.write(row, 13, duplicate_scripts)
      except:
       worksheet.write(row, 13, "")

      total_dup_scripts = lines[1].split(" ")[0]
      worksheet.write(row, 14, total_dup_scripts)

      lines = resultDeadCode.split("\n")

      try:
       dead_code = lines[2]
       worksheet.write(row, 15, dead_code)
      except:
       worksheet.write(row, 15, "")

      lines = resultInitialization.split("\n")

      try:
       initialization = lines[2]
       worksheet.write(row, 16, initialization)
      except:
       worksheet.write(row, 16, "")
      

      row = row + 1

   
    workbook.close()



def main():
    """The entrypoint for the analysis"""

    mastery = Mastery()
    mastery.process(sys.argv[1])		#Paso el zip por terminal	
    mastery.analyze()



