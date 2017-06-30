import os
import fnmatch
import csv
import shutil
import pprint

rootdir = "S:/" #location to look through files
newdir = "C:/Users/Bernadette.Werntges/OneDrive - CPower/s_drive_search/" #location to copy found files to
logfile = os.path.join(newdir,'logfile.csv') #creates a log file under the new directory
fieldnames = ['Filename', 'Path', 'Created Time', 'Modified Time', 'Size', 'File Type', 'Owner', 'Owner Domain'] # headers for logfile csv and organizing data
results = []
keywords = ['*curtailment*', '* cp *'] #list of keywords to look for in file names
exclude = set(["S:/KSQ File Share/siteapps/Documents",
"S:/Load Response/Siebel Documents/Site documentation",
"S:/Load Response/Air Permits and Certifications",
"S:/Load Response/COE/Jim Majsak Files/_Customers",
"S:/Load Response/COE/Mark Ramsay Files/Demand Response 2014",
"S:/Load Response/COE/Mark Ramsay Files/Demand Response 2013",
"S:/Load Response/COE/Mark Ramsay Files/Demand Response 2012",
"S:/Load Response/COE/Mark Ramsay Files/Demand Response 2011",
"S:/Load Response/ERCOT/Customers",
"S:/Load Response/ERCOT/Aaron's Files - Customer Folders",
"S:/Demand Response/Salesforce",
"S:/NY File Share/Curtailment Plans",
"S:/NY File Share/Curtailment Plans/Curtailment Plans",
"S:/Load Response/Regions/TX/Engineering"])


def parse_dirs(rootdir, newdir): # walk through folders, subfolders that are not in the exclude list, and collects metadata of files that meet the criteria
	for subdir, dirs, files in os.walk(rootdir, topdown=True):
		dirs[:] = [d for d in dirs if d not in exclude]
		for keyword in keywords:
			for file in fnmatch.filter(files, keyword): #filters out only files that contain the keywords
				if keyword in file.lower(): # if filename contains any of the keywords, copy the file to the new directory
					while True:
						try:
							path = os.path.join(subdir, file) #full file path (just adds the root file to the file name)
							[mode, ino, dev, nlink, uid, gid, size, atime, mtime, ctime] = os.stat(path) #collect file metadata
							rootfile_size+=size #add size of file to total rootfile size
							sd = win32security.GetFileSecurity (path, win32security.OWNER_SECURITY_INFORMATION) #to get owner, it is within security information
							owner_sid = sd.GetSecurityDescriptorOwner()
							ownername, domain, typ = win32security.LookupAccountSid (None, owner_sid)	
							filetype = os.path.splitext(file)[1]
							ctimedate = datetime.fromtimestamp(ctime).strftime('%Y/%m/%d %H:%M:%S') #convert seconds to date-time
							mtimedate = datetime.fromtimestamp(mtime).strftime('%Y/%m/%d %H:%M:%S') # convert seconds to date-time
							#shutil.copy2(path, newdir) #copies the file (path) to the new directory (newdir)
							results.append(dict(zip(fieldnames, [file,path,ctimedate,mtimedate,size,filetype, ownername, domain]))) # create a dictionary of {fieldnames: metadata} for the file and add to results list
						except error as e:
							print (e,file)
							time.sleep(15)
							continue

						#catch the winerror which is also usually some version of 'network drive not available' wait 15 seconds and try again
						except OSError as err:
							try:
								subdir = win32api.GetShortPathName(subdir) #winerror 3 when an old 8.3 shortname is used, the full path is not found. this passes the short path as subdir
								print (err, file)
								continue
							except OSError as stillerr:
								print (stillerr, file)
								time.sleep(15)
								continue
							except error as pywinerror:
								print (pywinerror, file)
								time.sleep(15)
								continue
						break
						
					"""	except shutil.Error as serror:
							print (serror)
							continue
					"""
	pprint.pprint(results) #this prints all of the results at the end
	#create logfile, add results
	with open(logfile, 'w', newline='') as csvfile:
		writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
		writer.writeheader()
		writer.writerows(results)
	

def create_new_dir(newdir): #creates the new directory if it doesn't exist already
	if not os.path.exists(newdir):
		os.mkdir(newdir)


def main():
	create_new_dir(newdir)
	parse_dirs(rootdir, newdir)

if __name__ == '__main__':
	main()