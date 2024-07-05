import sys
import os

#setting up directories
laapsangFolderDir = os.path.dirname(os.path.abspath(__file__))
readExcelDir = os.path.join(laapsangFolderDir, "readExcel.py")
reportsFolderDir = os.path.join(laapsangFolderDir, "Reports")

#Executes readExcel.py
with open(readExcelDir, 'r') as f:
    script_code = f.read()
exec(script_code)


#Processes the excel lab reports
for labReport in os.listdir(reportsFolderDir):
    if labReport.endswith(".xls"):
        labReportDir = os.path.join(reportsFolderDir, labReport)
        content = readFile(labReportDir)
        labInterp = labReport[:-5]
        #writeFile(reportsFolderDir, content, labInterp)

        workingContent = content["PT_INIT"] + "\n"  + "\n" + sys.argv[1]
        writeFile(reportsFolderDir, workingContent, labInterp)


