#MAC
windows = True 
testing = False

if (not windows):
	if (not testing):
		WATCHDIR = "/Volumes/Coag/Lupus/Lupus Reports/2. Need Supervisor Review/"   
		REPORTDIR = "/Volumes/Coag/Lupus/Lupus Reports/5. Final Report Archive/"
		RUNDIR = "/Volumes/Coag/Lupus/Lupus Reports/2. Need Supervisor Review/"
		OUTPUTDIR = "/Volumes/Coag/Lupus/Lupus Reports/3. Need Pathologist Review/"
		OUTPUTDBFILE = "./PastReportsDB.csv"	
	else:
		WATCHDIR  = "/Users/ernestchan/Desktop/Temp"
		REPORTDIR = "/Users/ernestchan/Desktop/reports/"
		RUNDIR    = "/Users/ernestchan/Desktop/Temp"
		OUTPUTDIR = "/Users/ernestchan/Desktop/Temp"
		OUTPUTDBFILE = "./PastReportsDB.csv"	

#WINDOWS
if (windows):
	if (not testing): 
		WATCHDIR = "N:\\Coag\\Lupus\\Lupus Reports\\2. Need Supervisor Review"
		REPORTDIR = "N:\\Coag\\Lupus\\Lupus Reports\\5. Final Report Archive\\"
		RUNDIR    = "N:\\Coag\\Lupus\\Lupus Reports\\3. Need Pathologist Review"	
		OUTPUTDIR = RUNDIR
		OUTPUTDBFILE = "N:\\Coag\\Lupus\\Lupus Reports\\Tools\\PastReportsDB.csv"	
	
	else:
		WATCHDIR = "N:\\Coag\\Lupus\\Lupus Reports\\2. Need Supervisor Review"
		REPORTDIR = "N:\\Coag\\Lupus\\Lupus Reports\\5. Final Report Archive\\"
		RUNDIR    = "N:\\Coag\\Lupus\\Lupus Reports\\3. Need Pathologist Review"	
		OUTPUTDIR = "N:\\Coag\\Lupus\\Lupus Reports\\6. Archived Auto Drafts\\test"	
		OUTPUTDBFILE = "N:\\Coag\\Lupus\\Lupus Reports\\Tools\\PastReportsDB.csv"	



