import json
import sys
import os
import zipfile
import xlsxwriter
import ast
from os import walk


class Mastery: 

  """Analyzer of projects sb2, and create an Excel with solutions"""
  def __init__(self):

    self.total_projects = []
    self.total_projects_id = 0


  """Start the analysis."""
  def process(self,filename):

   #_____Analysis with zip____#

   """
   zip_file = zipfile.ZipFile(filename, "r")
   print zip_file

   for i in zip_file.namelist():
      json_project = zip_file.extract(i)
      files = json_project.split("/")
      if len(files) == 9 and files[-1] != "last.json":
        self.total_projects.append(json_project)
   """
   
   #____Analysis with folder____#

   for (path, ficheros, archivos) in walk("/home/angela/Escritorio/drScratch/ProyectosGregorio/projects_2"):
     if self.total_projects_id == 0:
       self.total_projects_id = len(ficheros)
     for json_project in archivos:
       if json_project != "last.json":
          self.total_projects.append(path + "/" + json_project)



  def analyze(self):

    #Create the Excel
    workbook = xlsxwriter.Workbook('results_3.xlsx')			
    worksheet = workbook.add_worksheet()
    worksheet_2 = workbook.add_worksheet()

    #Set the columns widths.
    worksheet.set_column('A:N', 15)
    worksheet.set_column('M:X', 20)
    worksheet_2.set_column('A:D', 20)

    #Header
    bold = workbook.add_format({'bold': 1})
    worksheet.write('A1', 'Project', bold)
    worksheet.write('B1', 'Project id', bold)
    worksheet.write('C1', 'Project date', bold)
    worksheet.write('D1', 'Project hour', bold)
    worksheet.write('E1', 'Abstraction', bold)
    worksheet.write('F1', 'Parallelization', bold)
    worksheet.write('G1', 'Logic', bold)
    worksheet.write('H1', 'Synchronization', bold)
    worksheet.write('I1', 'FlowControl', bold)
    worksheet.write('J1', 'UserInteractivity', bold)
    worksheet.write('K1', 'DataRepresent.', bold)
    worksheet.write('L1', 'TotalMastery', bold)
    worksheet.write('M1', 'AverageMastery', bold)
    worksheet.write('N1', 'Level', bold)
    worksheet.write('O1', 'DefaultSpriteNames', bold)
    worksheet.write('P1', 'TotalDefaultNames', bold)
    worksheet.write('Q1', 'DuplicateScripts', bold)
    worksheet.write('R1', 'TotalDupScripts', bold)
    worksheet.write('S1', 'DeadCode', bold)
    worksheet.write('T1', 'Total blocks DeadCode', bold)
    worksheet.write('U1', 'AttributeInitialization', bold)
    worksheet.write('V1', 'Total AttributeInitialization', bold)
    worksheet.write('W1', 'Block Count', bold)
    worksheet.write('X1', 'Total blocks', bold)

    worksheet_2.write('A1', 'Total projects', bold)
    worksheet_2.write('B1', 'Total id projects', bold)
    worksheet_2.write('C1', 'Total wrong projects', bold)


    row = 1
    wrong_count = 0
    total_projects = 0

    for path_file in self.total_projects:

      total_projects = total_projects + 1

      print path_file
      file_name = path_file.split("/")[-1]
      project_id = file_name.split("_")[0]
      p = file_name.split("_")[1:]
      date = p[0].split("-")[0] + "-" + p[0].split("-")[1] + "-" + p[0].split("-")[2]
      hour = p[0].split("-")[3] + ":" + p[0].split("-")[4] 


      #Request to hairball
      metricMastery = "hairball -p mastery.Mastery " + path_file
      metricSpriteNaming = "hairball -p convention.SpriteNaming " + path_file
      metricDuplicateScript = "hairball -p duplicate.DuplicateScripts " + path_file
      metricDeadCode = "hairball -p blocks.DeadCode " + path_file
      metricInitialization = "hairball -p initialization.AttributeInitialization " + path_file
      metricCountBlocks = "hairball -p blocks.BlockCounts " + path_file


      try:
        #Response from hairball
        resultMastery = os.popen(metricMastery).read()
        resultSpriteNaming = os.popen(metricSpriteNaming).read()
        resultDuplicateScript = os.popen(metricDuplicateScript).read()
        resultDeadCode = os.popen(metricDeadCode).read()
        resultInitialization = os.popen(metricInitialization).read()
        resultCountBlocks = os.popen(metricCountBlocks).read()

      
        #________Print results in Excel_________#

        #Project name, complete
        project_name = file_name
        worksheet.write(row, 0, project_name) 


        #Project_id
        worksheet.write(row, 1, project_id)


        #Date
        worksheet.write(row, 2, date)

        
        #Hour
        worksheet.write(row, 3, hour)


        #Mastery categories
        lines = resultMastery.split("\n") 
        dic = ast.literal_eval(lines[1])
        worksheet.write(row, 4, dic["Abstraction"])
        worksheet.write(row, 5, dic["Parallelization"])
        worksheet.write(row, 6, dic["Logic"])
        worksheet.write(row, 7, dic["Synchronization"])
        worksheet.write(row, 8, dic["FlowControl"])
        worksheet.write(row, 9, dic["UserInteractivity"])
        worksheet.write(row, 10, dic["DataRepresentation"])


        #Total_mastery
        total_mastery = lines[2].split(":")[1]
        worksheet.write(row, 11, total_mastery)


        #Average_mastery
        average_mastery = lines[3].split(":")[1]
        worksheet.write(row, 12, average_mastery)


        #Overall programming competence
        level = lines[4].split(":")[1]
        worksheet.write(row, 13, level)


        #Default sprite names
        sprite_names = resultSpriteNaming.split("\n")[2:-1]
        worksheet.write(row, 14, str(sprite_names))


        #Total of default sprite names
        lines = resultSpriteNaming.split("\n")[1]
        total_default_names = lines.split(" ")[0]
        worksheet.write(row, 15, total_default_names)


        #Duplicate scripts
        lines = resultDuplicateScript.split("\n")
        try:
         duplicate_scripts = lines[2:-1]
         result = ""
         for dup in duplicate_scripts:
           result += dup
           result += "\n"
         worksheet.write(row, 16, result)
        except:
         worksheet.write(row, 16, "")


        #Total number of duplicate scripts
        total_dup_scripts = lines[1].split(" ")[0]
        if total_dup_scripts != '':
          worksheet.write(row, 17, str(total_dup_scripts))
        else:
          worksheet.write(row, 17, "0")


        #Dead code
        lines = resultDeadCode.split(".sb2")
        dead_code = ast.literal_eval(lines[1])
        try:
         worksheet.write(row, 18, str(dead_code))
        except:
         worksheet.write(row, 18, "")


        #Total blocks of dead code
        if len(dead_code) == 0:
            worksheet.write(row, 19, "0")            
        else:
            worksheet.write(row, 19, len(resultDeadCode.split(",")))


        #Attribute initialization
        lines = resultInitialization.split(".sb2")
        dicc = ast.literal_eval(lines[1])
        keys = dicc.keys()
        values = dicc.values()
        items = dicc.items()
        
        for keys, values in items:
          list = []
          attribute = ""
          internalkeys = values.keys()
          internalvalues = values.values()
          internalitems = values.items()
          flag = False
          counterFlag = False

          for internalkeys, internalvalues in internalitems:
            if internalvalues == 1:
                counterFlag = True
                for value in list:
                   if internalvalues == value:
                        flag = True
                if not flag:
                   list.append(internalkeys)
                   if len(list) < 2:
                      attribute = str(internalkeys)
                   else:
                      attribute = attribute + ", " + str(internalkeys)

          dicc[keys] = attribute
    
        removekeys = []
        for key, value in dicc.iteritems():
           if dicc[key] == "":
              removekeys.append(key)

        for removekey in removekeys:
           del dicc[removekey]

        worksheet.write(row, 20, str(dicc))


        #Total AttributeInitialization
        totalAtt = len(dicc.keys())
        worksheet.write(row, 21, str(totalAtt))


        #Count of total blocks
        lines = resultCountBlocks.split("\n")
        countBlocks = lines[1:-2]
        worksheet.write(row, 22, str(countBlocks))


        #Total blocks
        totalBlocks = lines[-2]
        worksheet.write(row, 23, str(totalBlocks))
         
      

      except:
        worksheet.write(row, 4, "Wrong project")
        wrong_count = wrong_count + 1
        

      row = row + 1


    #Summary of stats   
    worksheet_2.write(1, 0, total_projects)
    worksheet_2.write(1, 1, self.total_projects_id)
    worksheet_2.write(1, 2, wrong_count)
    
    workbook.close()



def main():
    """The entrypoint for the analysis"""

    mastery = Mastery()
    #With zip
    mastery.process(sys.argv[1])
    mastery.analyze()



