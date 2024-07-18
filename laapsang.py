import os

def testLevelCheck():
    level = 0

    if all(report[key] != -1 for key in ["LAPTT_R", "DRVVS_R", "DPTS_R", "IGG_ACA_R", "IGM_ACA_R", "IGG_B2_R", "IGM_B2_R"]) and all(report[key] == -1 for key in ["IGA_ACA_R", "IGA_B2_R", "IGG_PSPT_R", "IGM_PSPT_R"]):
        level = 1
    
    elif all(report[key] != -1 for key in ["IGA_ACA_R", "IGA_B2_R", "IGG_PSPT_R", "IGM_PSPT_R"]) and all(report[key] == -1 for key in ["LAPTT_R", "DRVVS_R", "DPTS_R", "IGG_ACA_R", "IGM_ACA_R", "IGG_B2_R", "IGM_B2_R"]):
        level = 2
    
    elif all(report[key] != -1 for key in ["LAPTT_R", "DRVVS_R", "DPTS_R"]) and all(report[key] == -1 for key in ["IGG_ACA_R", "IGM_ACA_R", "IGG_B2_R", "IGM_B2_R", "IGA_ACA_R", "IGA_B2_R", "IGG_PSPT_R", "IGM_PSPT_R"]):
        level = 3
    
    elif all(report[key] != -1 for key in ["IGG_ACA_R", "IGM_ACA_R", "IGG_B2_R", "IGM_B2_R"]) and all(report[key] == -1 for key in ["LAPTT_R", "DRVVS_R", "DPTS_R", "IGA_ACA_R", "IGA_B2_R", "IGG_PSPT_R", "IGM_PSPT_R"]):
        level = 4
    
    return level


def aPTT_Section():

    aPTT_Report = ""

    if report["LAPTT_R"] > report["LAPTT_U"]:
        aPTT_Report += (f"The initial clotting time in the aPTT-based testing phase is {report["LAPTT_R"]:.1f} seconds, which is {(report["LAPTT_R"] - report["LAPTT_U"]):.1f} seconds above the upper limit.")

        #Mixed Section
        #should I add a line if we see a prolognation (which shouldn't happen?) 
        # Add prolongation line      
        if (report["PTTMX_R"] - report["PTTMX_U"]) >= 0:
            aPTT_Report += (f" The clotting time in the mixing phase of the aPTT-based system shortens but still remains above the upper limit by {(report["PTTMX_R"] - report["PTTMX_U"]):.1f} seconds.")
        elif report["PTTMX_R"] - report["PTTMX_U"] <= 0 :
            aPTT_Report += (f" The clotting time in the mixing phase of the aPTT-based system is now within the normal reference interval.")

        #STACLOT_LA Section
        # If tet is positive, whole section is positive
        if report["SCCOR_R"] > report["SCCOR_U"]:
            aPTT_Report += (f" In the confirmatory phase of the aPTT-based system (Staclot-LA), the clotting time abnormally corrects by {(report["SCCOR_R"] * 100):.1f}% in the presence of high concentration phospholipid (99th percentile upper limit of normal {(report["SCCOR_U"] * 100):.1f}%).")
            testStatus["aPTT"] = "Positive"
        else:
            aPTT_Report += (f" In the confirmatory phase of the aPTT-based system (Staclot-LA), the clotting time has no abnormal shortening (patient: {(report["SCCOR_R"] * 100):.1f}, 99th percentile upper limit of normal {(report["SCCOR_U"] * 100):.1f}%).")
            testStatus["aPTT"] = "Negative"

        #LTT Section
        if report["LTT_R"] != -1:           
            if report["LTT_R"] < report["LTT_L"]:
                aPTT_Report += (f" The thrombin time is {(report["LTT_L"] - report["LTT_R"]):.1f} seconds shorter than the lower limit normal of {report["LTT_L"]:.1f} seconds.")
            elif report["LTT_R"] > report["LTT_U"]:
                aPTT_Report += (f" The thrombin time is {(report["LTT_R"] - report["LTT_U"]):.1f} seconds prolonged above the upper limit normal of {report["LTT_U"]:.1f} seconds")
            else:
                aPTT_Report += (f" The thrombin time is normal.")

        #LTTHEP Section
        if report["LTT_R"] < report["LTTHEP_R"]:
            aPTT_Report += (f" Following incubation of the plasma with heparinase, the thrombin time does not shorten, but remains prolonged by {(report["LTTHEP_P"]):.1f} seconds.")
        elif report["LTTHEP_R"] < report["LTTHEP_U"]:
            aPTT_Report += (f" Following incubation of the plasma with heparinase, the thrombin time shortens by {(report["LTT_R"] - report["LTTHEP_R"]):.1f} seconds, returning to within the normal range.")
        else:
            aPTT_Report += (f" Following incubation of the plasma with heparinase, the thrombin time shortens by {(report["LTT_R"] - report["LTTHEP_R"]):.1f} seconds.")

    #Negative
    else:
        aPTT_Report += (f"The intial clotting time is normal in the aPTT-based testing phase.")
        testStatus["aPTT"] = "Negative"
    
    return aPTT_Report

def DRVVT_DPT_Sections(testname, initR, initU, mixR, mixU, corrR, percentR, percentU):
    
    battery_Report = ""

    #Initial clotting time section
    if initR > initU:
        battery_Report += (f"The initial clotting time in the {testname}-based testing phase is {initR:.1f} seconds, which is {(initR - initU):.1f} seconds above the upper limit normal.")

        #Mixed Section
        #should I add a line if we see a prolognation (which shouldn't happen?)
        if mixR > mixU:
            battery_Report += (f" The clotting time in the mixing phase of the {testname}-based system shortens but still remains above the upper limit by {(mixR - mixU):.1f} seconds.")
        else:
            battery_Report += (f" The clotting time in the mixing phase of the {testname}-based system is now within the normal reference interval.")

        #High concentration phospholipid
        if corrR != -1:
            prefix = ("In the confirmatory phase, with high-concentration phospholipid,")
            suffix = (f"(patient {(percentR * 100):.1f}%, upper limit of normal {(percentU * 100):.1f}%).")
            
            #Im assuming this should Never happen?
            if corrR > initR:
                battery_Report +=  (f" {prefix} there is no shortening of the clotting time.")
            #-----------------------------------

            elif percentR > percentU:
                battery_Report += (f" {prefix} the clotting time shortens abnormally {suffix}")
                testStatus[f"{testname}"] = "Positive"

            elif percentR <= percentU:
                battery_Report += (f" {prefix} the clotting time has no abnormal shortening {suffix}")
                testStatus[f"{testname}"] = "Negative"
    else:
        battery_Report += (f"The inital clotting time is normal in the {testname}-based testing phase.")
        testStatus[f"{testname}"] = "Negative"

    return battery_Report

#If both tests of the same type are negative, should I make a new function to generate a seperate combined sentence?
def Antigenic_Section(testR, testU, testName, testType, unit):
    
    antigenicReport = ""

    if testR > testU:
        antigenicReport += (f"Solid phase testing for {testName} {testType} is elevated at {testR} {unit} (99th percentile upper limit of normal {testU} {unit}).")
        testStatus["Antigenic_Positive"] = "Positive"
    else:
        antigenicReport += (f"Solid phase testing is negative for {testName} {testType} antibodies.")
        testStatus[f"{testName}{testType}"] = "Negative"
    
    return antigenicReport

#Conclusion for Level 1 testing
def Conclusion_L1():

    conclusion = ""
    testingSystems = ""
    Positive_Tests = []

    if "aPTT" in testStatus:
        if testStatus["aPTT"] == "Positive":
            Positive_Tests.append("aPTT-based")
    
    if "DRVVT" in testStatus:
        if testStatus["DRVVT"] == "Positive":
            Positive_Tests.append("DRVVT-based")

    if "DPT" in testStatus:
        if testStatus["DPT"] == "Positive":
            Positive_Tests.append("DPT-based")
    
    if "Antigenic_Positive" in testStatus:
        if testStatus["Antigenic_Positive"] == "Positive":
            Positive_Tests.append("Antigenic-based")
    
    if len(Positive_Tests) == 4:
        testingSystems += ", notably in all four testing systems, does provide some"

    elif len(Positive_Tests) == 3:
        testingSystems += (f", notably in the {Positive_Tests[0]}, {Positive_Tests[1]} and {Positive_Tests[2]} testing systems, does provide some")
    
    elif len(Positive_Tests) == 2:
        testingSystems += (f", notably in the {Positive_Tests[0]} and {Positive_Tests[1]} testing systems, does provide some")

    elif len(Positive_Tests) == 1:
        testingSystems += (f", notably in the {Positive_Tests[0]} testing system, does provide some")
    
    elif len(Positive_Tests) == 0:
        testingSystems += (f" does not provide")
    
    conclusion = (f"The current study{testingSystems} laboratory evidence for antiphospholipid syndrome.")

    if len(Positive_Tests) <= 3:
        conclusion += (" No additional evidence is contributed by the other testing systems employed in this study.")
    
    if len(Positive_Tests) > 0:
        conclusion += (" Confirmatory repeat LA/APS testing after at least 12 weeks is recommended.")

    return conclusion

#Conclusion for Level 2 testing
def Conclusion_L2_L4():

    conclusion = ""
    testingSystems = ""
    Positive_Tests = []

    if "Antigenic_Positive" in testStatus:
        if testStatus["Antigenic_Positive"] == "Positive":
            Positive_Tests.append("Antigenic-based")
    
    if len(Positive_Tests) == 1:
        testingSystems += (f", notably in the {Positive_Tests[0]} testing system, does provide some")
    
    elif len(Positive_Tests) == 0:
        testingSystems += (f" does not provide")

    conclusion = (f"The current study{testingSystems} laboratory evidence for antiphospholipid syndrome.")

    return conclusion

def Conclusion_L3():

    conclusion = ""
    testingSystems = ""
    Positive_Tests = []

    if "aPTT" in testStatus:
        if testStatus["aPTT"] == "Positive":
            Positive_Tests.append("aPTT-based")
    
    if "DRVVT" in testStatus:
        if testStatus["DRVVT"] == "Positive":
            Positive_Tests.append("DRVVT-based")

    if "DPT" in testStatus:
        if testStatus["DPT"] == "Positive":
            Positive_Tests.append("DPT-based")
    
    if len(Positive_Tests) == 3:
        testingSystems += (f", notably in the {Positive_Tests[0]}, {Positive_Tests[1]} and {Positive_Tests[2]} testing systems, does provide some")
    
    elif len(Positive_Tests) == 2:
        testingSystems += (f", notably in the {Positive_Tests[0]} and {Positive_Tests[1]} testing systems, does provide some")

    elif len(Positive_Tests) == 1:
        testingSystems += (f", notably in the {Positive_Tests[0]} testing system, does provide some")
    
    elif len(Positive_Tests) == 0:
        testingSystems += (f" does not provide")
    
    conclusion = (f"The current study{testingSystems} laboratory evidence for antiphospholipid syndrome.")

    return conclusion



def FinalReport(level):

    FinalReport = ""

    if level == 1:

        #aPTT
        FinalReport += aPTT_Section() + "\n\n"

        #DRVVT
        FinalReport += DRVVT_DPT_Sections("DRVVT", report["DRVVS_R"], report["DRVVS_U"], report["DRVVMX_R"], report["DRVVMX_U"], report["DRVVC_R"], report["PCTCO_R"], report["PCTCO_U"]) + "\n\n"

        #DPT
        FinalReport += DRVVT_DPT_Sections("DPT", report["DPTS_R"], report["DPTS_U"], report["DPTMX_R"], report["DPTMX_U"], report["DPTC_R"], report["DPTCOR_R"], report["DPTCOR_U"]) + "\n\n"

        #Battery 1
        FinalReport += Antigenic_Section(report["IGG_ACA_R"], report["IGG_ACA_U"], "IgG", "anti-cardiolipin", "CU")
        
        FinalReport += " " + Antigenic_Section(report["IGM_ACA_R"], report["IGM_ACA_U"], "IgM", "anti-cardiolipin", "CU")

        FinalReport += " " + Antigenic_Section(report["IGG_B2_R"], report["IGG_B2_U"], "IgG", "anti-beta-2 glycoprotein-1", "CU")

        FinalReport += " " + Antigenic_Section(report["IGM_B2_R"], report["IGM_B2_U"], "IgG", "anti-beta-2 glycoprotein-1", "CU")

        FinalReport += "\n\n" + Conclusion_L1()

    if level == 2:

        #Battery 2
        FinalReport += Antigenic_Section(report["IGA_ACA_R"], report["IGA_ACA_U"], "IgA", "anti-cardiolipin", "APL")
        
        FinalReport += " " + Antigenic_Section(report["IGA_B2_R"], report["IGA_B2_U"], "IgA", "anti-beta-2 glycoprotein-1", "U/mL")

        FinalReport += " " + Antigenic_Section(report["IGG_PSPT_R"], report["IGG_PSPT_U"], "IgG", "phosphatidylserine/prothrombin complex", "U/mL")

        FinalReport += " " + Antigenic_Section(report["IGM_PSPT_R"], report["IGM_PSPT_U"], "IgM", "phosphatidylserine/prothrombin complex", "U/mL")

        FinalReport += "\n\n" + Conclusion_L2_L4()

    if level == 3:

        #aPTT
        FinalReport += aPTT_Section() + "\n\n"

        #DRVVT
        FinalReport += DRVVT_DPT_Sections("DRVVT", report["DRVVS_R"], report["DRVVS_U"], report["DRVVMX_R"], report["DRVVMX_U"], report["DRVVC_R"], report["PCTCO_R"], report["PCTCO_U"]) + "\n\n"

        #DPT
        FinalReport += DRVVT_DPT_Sections("DPT", report["DPTS_R"], report["DPTS_U"], report["DPTMX_R"], report["DPTMX_U"], report["DPTC_R"], report["DPTCOR_R"], report["DPTCOR_U"]) + "\n\n"

        FinalReport += "\n\n" + Conclusion_L3()

    if level == 4:

        #Battery 1
        FinalReport += Antigenic_Section(report["IGG_ACA_R"], report["IGG_ACA_U"], "IgG", "anti-cardiolipin", "CU")
        
        FinalReport += " " + Antigenic_Section(report["IGM_ACA_R"], report["IGM_ACA_U"], "IgM", "anti-cardiolipin", "CU")

        FinalReport += " " + Antigenic_Section(report["IGG_B2_R"], report["IGG_B2_U"], "IgG", "anti-beta-2 glycoprotein-1", "CU")

        FinalReport += " " + Antigenic_Section(report["IGM_B2_R"], report["IGM_B2_U"], "IgG", "anti-beta-2 glycoprotein-1", "CU")

        FinalReport += "\n\n" + Conclusion_L2_L4()

    return FinalReport




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
        report = readFile(labReportDir)
        testStatus = {}
        level = testLevelCheck()

        content = FinalReport(level)
        
        labInterpFileName = labReport[:-4]
        writeFile(reportsFolderDir, content, labInterpFileName)

