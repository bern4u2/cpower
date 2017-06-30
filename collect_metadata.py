import os
import csv
from datetime import datetime
import win32api
import win32con
import win32security
from pywintypes import error
import time

"""NOTES: put in logfile already open error
add a cache
"""

fieldnames = ['Filename', 'Path', 'Created Time', 'Modified Time', 'Accessed Time', 'Size', 'File Type', 'Owner', 'Owner Domain'] # headers for logfile csv and organizing data
results = []
exclude = ["S:\\Load Response\\Siebel Documents\\Site documentation\\Middlesex County Regional School District\\MRESC - Old Bridge\\MRESC - Old Bridge High School\\James A McDivitt 4252",
"Work in process - Only for Working - please do not send anything from this folder"]
#a file in MRESC was causing problems, I manually looked through it and none were curtailment plans, exclude 'work in process' folder per Arusyak

def collect_metadata(rootdir, newdir, logfilename): # walk through all folders, subfolders in the root directory, and collect metadata
	start_time = datetime.now()
	rootfile_size = 0
	file_count = 0
	dirs_count = 0
	logfile = os.path.join(newdir,logfilename +'.csv') #creates a log file under the new directory
	for subdir, dirs, files in os.walk(rootdir, followlinks=True):
		dirs[:] = [d for d in dirs if d not in exclude] #exclude the files in 'exclude'
		dirs_count +=1
		for file in files:
			file_count +=1
			while True: #keep looping until all files are finished or an error occurs that is not included in the exceptions
				try:
					path = os.path.join("r'",subdir, file) #full file path (just adds the root file to the file name)
					[mode, ino, dev, nlink, uid, gid, size, atime, mtime, ctime] = os.stat(path) #collect file metadata
					rootfile_size+=size #add size of file to total rootfile size
					sd = win32security.GetFileSecurity (path, win32security.OWNER_SECURITY_INFORMATION) #to get owner, it is within security information
					owner_sid = sd.GetSecurityDescriptorOwner()
					ownername, domain, typ = win32security.LookupAccountSid (None, owner_sid)	
					filetype = os.path.splitext(file)[1]
					ctimedate = datetime.fromtimestamp(ctime).strftime('%Y/%m/%d %H:%M:%S') #convert seconds to date-time
					mtimedate = datetime.fromtimestamp(mtime).strftime('%Y/%m/%d %H:%M:%S') # convert seconds to date-time
					curtime = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
					print (file_count) # tracking purposes for large directories
					results.append(dict(zip(fieldnames, [file,path,ctimedate,mtimedate,curtime,size,filetype, ownername, domain]))) # create a dictionary of {fieldnames: metadata} for the file and add to results list

					#catch the pesky pywintypes errors which are usually some version of 'network drive not available' print the error and what file it's on, wait 15 seconds and try again
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
		
	#create logfile, add results
	with open(logfile, 'w', newline='', encoding='utf-8') as csvfile:
		writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
		writer.writeheader() # headers are the fieldnames
		writer.writerows(results)
	
	end_time = datetime.now()
	total_time = end_time - start_time # elapsed time
	print (rootdir, "took ", total_time.days, 'days', total_time.seconds, 'seconds', total_time.microseconds, 'microseconds to parse through', \
		dirs_count, ' directories,', file_count, 'files, a total size of', rootfile_size*(2**(-20)), 'MB')

	return results
def create_new_dir(newdir): #creates the new directory if it doesn't exist already
	if not os.path.exists(newdir):
		os.mkdir(newdir)


def main():
	rootdir = input('Enter root directory to parse through: ') #location to look through files
	'''
	followlinks = None
	while followlinks is None:
		links_question= input('Do you want to follow links under this root directory? (Yes/No): ')
		if links_question.lower() == 'yes':
			followlinks = True
		elif links_question.lower() == 'no':
			followlinks = False
		else:
			print ("Please type 'Yes' or 'No'")
			'''
	newdir = input('Enter directory to copy files to or where to store logfile: ')
	logfilename = input('Name of log file to be stored in above directory (i.e. ISONE_logfile): ')
	create_new_dir(newdir)
	collect_metadata(rootdir, newdir, logfilename)

if __name__ == '__main__':
	main()