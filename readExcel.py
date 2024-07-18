import xlrd
import os
import docx
from docx import Document
from xlrd import open_workbook
import re 




#If number is float or int, return it without touching it.
#If number has a space, "<" or ">", extract only the float/int and return it
def sanitize(number):
    if isinstance(number, (int, float)):
        return number
    
    elif number is None or number == "":
        return -1

    else:
        trueValue = re.match('^([0-9.]+)',number)
        if trueValue:
            return float(trueValue.group(1))
        
        #trueValue = re.match('.*?([0-9.]+)',number)
        #if trueValue:
            #return float(trueValue.group(1))
        
        trueValue = re.match('^<([0-9.]+)', number)
        if trueValue:
            return float(trueValue.group(1))
        
        trueValue = re.match('^>([0-9.]+)', number)
        if trueValue:
            return float(trueValue.group(1))


#Reads xls file and appends to dictionary
def readFile(filename):
    book = open_workbook(filename, on_demand=True)
    for sheetname in book.sheet_names():
        if sheetname.endswith('Lab'):
            sheet = book.sheet_by_name(sheetname)

    a = {}
    try: 
        #Demographics
        a["INIT"] = sheet.cell_value(3,5)
        a["MRN"] = sheet.cell_value(4,5)
        a["DATE"] = sheet.cell_value(5,5)
        a["SPEC_ID"] = sheet.cell_value(6,5)
        
        #aPTT results
        a["LAPTT_R"] = sanitize(sheet.cell_value(11,1))
        a["PTTMX_R"] = sanitize(sheet.cell_value(14,1))
        a["LTT_R"] = sanitize(sheet.cell_value(12,1))
        a["LTTHEP_R"] = sanitize(sheet.cell_value(13,1))
        
        a["LAPTT_U"] = sanitize(sheet.cell_value(11,3))
        a["LTT_U"] = sanitize(sheet.cell_value(12,3))
        a["LTTHEP_U"] = sanitize(sheet.cell_value(13,3))
        a["PTTMX_U"] = sanitize(sheet.cell_value(14,3))
        
        a["LTT_L"] = sanitize(sheet.cell_value(12,2))

        a["LAPTT_P"] = sanitize(sheet.cell_value(11,4))
        a["LTT_P"] = sanitize(sheet.cell_value(12,4))
        a["LTTHEP_P"] = sanitize(sheet.cell_value(13,4))
        a["PTTMX_P"] = sanitize(sheet.cell_value(14,4))
        
        a["SCLA1_R"] = sanitize(sheet.cell_value(11,6))
        a["SCLA2_R"] = sanitize(sheet.cell_value(12,6))
        
        #This returns 7 on a blank sheet for some reason
        a["SCCOR_R"] = sanitize(sheet.cell_value(13,6))
        a["SCCOR_U"] = sanitize(sheet.cell_value(13,7))
        
        #DRVVT results
        a["DRVVS_R"] = sanitize(sheet.cell_value(25,1))
        a["DRVVMX_R"] = sanitize(sheet.cell_value(26,1))
        a["DRVVC_R"] = sanitize(sheet.cell_value(27,1))
        a["PCTCO_R"] = sanitize(sheet.cell_value(28,1))
        
        a["DRVVS_U"] = sanitize(sheet.cell_value(25,3))
        a["DRVVMX_U"] = sanitize(sheet.cell_value(26,3))
        a["PCTCO_U"] = sanitize(sheet.cell_value(28,3))
        
        a["DRVVS_P"] = sanitize(sheet.cell_value(25,4))
        a["DRVVMX_P"] = sanitize(sheet.cell_value(26,4))
        a["PCTCO_P"] = sanitize(sheet.cell_value(28,4))
        
        #DPT results
        a["DPTS_R"] = sanitize(sheet.cell_value(33,1))
        a["DPTMX_R"] = sanitize(sheet.cell_value(34,1))
        a["DPTC_R"] = sanitize(sheet.cell_value(35,1))
        a["DPTCOR_R"] = sanitize(sheet.cell_value(36,1))
        
        a["DPTS_U"] = sanitize(sheet.cell_value(33,3))
        a["DPTMX_U"] = sanitize(sheet.cell_value(34,3))
        a["DPTCOR_U"] = sanitize(sheet.cell_value(36,3))
        
        a["DPTS_P"] = sanitize(sheet.cell_value(33,4))
        a["DPTMX_P"] = sanitize(sheet.cell_value(34,4))
        a["DPTCOR_P"] = sanitize(sheet.cell_value(36,4))

        #Antigenics - Battery 1
        a["IGG_ACA_R"] = sanitize(sheet.cell_value(41,1))
        a["IGM_ACA_R"] = sanitize(sheet.cell_value(42,1))
        a["IGG_B2_R"] = sanitize(sheet.cell_value(43,1))
        a["IGM_B2_R"] = sanitize(sheet.cell_value(44,1))
        a["IGA_ACA_R"] = sanitize(sheet.cell_value(48,1))
        a["IGA_B2_R"] = sanitize(sheet.cell_value(49,1))
        a["IGG_PSPT_R"] = sanitize(sheet.cell_value(50,1))
        a["IGM_PSPT_R"] = sanitize(sheet.cell_value(51,1))
        
        #Antigenics - Battery 2
        a["IGG_ACA_U"] = sanitize(sheet.cell_value(41,3))
        a["IGM_ACA_U"] = sanitize(sheet.cell_value(42,3))
        a["IGG_B2_U"] = sanitize(sheet.cell_value(43,3))
        a["IGM_B2_U"] = sanitize(sheet.cell_value(44,3))
        a["IGA_ACA_U"] = sanitize(sheet.cell_value(48,3))
        a["IGA_B2_U"] = sanitize(sheet.cell_value(49,3))
        a["IGG_PSPT_U"] = sanitize(sheet.cell_value(50,3))
        a["IGM_PSPT_U"] = sanitize(sheet.cell_value(51,3))
        
    except:
        a["INIT"] = -2
        a["MRN"] = -2
        a["DATE"] = -2
        a["SPEC_ID"] = -2
        
        a["LAPTT_R"] = -2
        a["PTTMX_R"] = -2
        a["LTT_R"] = -2
        a["LTTHEP_R"] = -2
        
        a["LAPTT_U"] = -2
        a["LTT_U"] = -2
        a["LTTHEP_U"] = -2
        a["PTTMX_U"] = -2
        
        a["LAPTT_P"] = -2
        a["LTT_P"] = -2
        a["LTTHEP_P"] = -2
        a["PTTMX_P"] = -2
        
        a["SCLA1_R"] = -2
        a["SCLA2_R"] = -2
        
        a["SCCOR_R"] = -2
        a["SCCOR_U"] = -2
        
        a["DRVVS_R"] = -2
        a["DRVVMX_R"] = -2
        a["DRVVC_R"] = -2
        a["PCTCO_R"] = -2
        
        a["DRVVS_U"] = -2
        a["DRVVMX_U"] = -2
        a["PCTCO_U"] = -2
        
        a["DRVVS_P"] = -2
        a["DRVVMX_P"] = -2
        a["PCTCO_P"] = -2
        
        a["DPTS_R"] = -2
        a["DPTMX_R"] = -2
        a["DPTC_R"] = -2
        a["DPTCOR_R"] = -2
        
        a["DPTS_U"] = -2
        a["DPTMX_U"] = -2
        a["DPTCOR_U"] = -2
        
        a["DPTS_P"] = -2
        a["DPTMX_P"] = -2
        a["DPTCOR_P"] = -2
        
        a["IGG_ACA_R"] = -2
        a["IGM_ACA_R"] = -2
        a["IGG_B2_R"] = -2
        a["IGM_B2_R"] = -2
        a["IGA_ACA_R"] = -2
        a["IGA_B2_R"] = -2
        a["IGG_PSPT_R"] = -2
        a["IGM_PSPT_R"] = -2
        
        a["IGG_ACA_U"] = -2
        a["IGM_ACA_U"] = -2
        a["IGG_B2_U"] = -2
        a["IGM_B2_U"] = -2
        a["IGA_ACA_U"] = -2
        a["IGA_B2_U"] = -2
        a["IGG_PSPT_U"] = -2
        a["IGM_PSPT_U"] = -2
        
    return a 
           

#Writes a doc
def writeFile(path, content, filename):

    interpretedDoc = Document()
    interpretedDoc.add_paragraph(content)
    outputPath = os.path.join(path, filename + ".doc")
    interpretedDoc.save(outputPath)
    