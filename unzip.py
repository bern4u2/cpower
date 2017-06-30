import zipfile
import csv
import os

#rootdir = https://cpower.sharepoint.com/CNE Trinity Air Permitting Data/Curtailment Plan Project/S drive files/
exclude = ["C:/Users/Bernadette.Werntges/OneDrive - CPower/ELMO/S drive file info/zipfiles", 
"S:/KSQ File Share/siteapps/Documents/L/Lifebridge Health/LI30SD22 Northwest Recommision and OxiCat/NW OxiCat Finance document.zip",
"S:/KSQ File Share/siteapps/Documents/M/Metropolitan Water Reclamation District of Greater Chicago/SummaryUsageData.zip"]


def zipextractor(rootdir,zipdir): #goes through the S drive files and looks for .zip extensions, then extracts the file to a new folder
	for subdir, dirs, files in os.walk(rootdir):
		dirs[:] = [] #don't go into subfolders (because subfolder is where the zip extracted data is being saved), just look at the files
		for file in files:
			if file == "Salesforce_logfile.csv":
				path = os.path.join("r'",subdir, file) #defining path to open the file
				print ('reading ', path)
				extractdir = os.path.join(zipdir, os.path.splitext(file)[0] + '_zips') #new dir within zipdir for name of this file
				if not os.path.exists(extractdir):#create this dir if it doesn't exist
					os.mkdir(extractdir) 
				if os.path.splitext(file)[1] == '.csv':
					with open(path, 'rU', newline ='') as logfile: #open the file
						reader = csv.DictReader(logfile, dialect = 'excel') #in read mode
						for row in reader:#look through all rows
							if row['File Type'].lower() == '.zip': #if the file type is a zip, then do other things
								subzipfile = os.path.splitext(row['Filename'])[0] #name of zip file
								subextractdir = os.path.join(extractdir + os.sep, subzipfile) #create a new path zipdir > name of file > name of zip file
								print (subextractdir)
								if not os.path.exists(subextractdir): #if it doesn't already exist, create it and extract
									os.mkdir(subextractdir)
									with zipfile.ZipFile(row['Path'], 'r') as zip_ref:
										zip_ref.extractall(subextractdir)
						
				
def main():
	#rootdir = "https://cpower.sharepoint.com/CNE Trinity Air Permitting Data/Curtailment Plan Project/S drive files/"
	#zipdir = "https://cpower.sharepoint.com/CNE Trinity Air Permitting Data/Curtailment Plan Project/S drive files/zipfiles"
	rootdir = "C:/Users/Bernadette.Werntges/OneDrive - CPower/ELMO/S drive file info"
	zipdir = "C:/Users/Bernadette.Werntges/OneDrive - CPower/ELMO/S drive file info/zipfiles"
	if not os.path.exists(zipdir):
		os.mkdir(zipdir)

	zipextractor(rootdir, zipdir)

if __name__ == '__main__':
	main()