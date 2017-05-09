import os
import csv
from datetime import datetime
import win32api
import win32con
import win32security
from pywintypes import error
import time

"""NOTES: put in S drive access error, logfile already open error
put in timestamp start and end
add a cache
"""

rootdir = "S:\KSQ File Share\siteapps\Documents" #location to look through files
newdir = "C:/Users/Bernadette.Werntges/OneDrive - CPower/test/" #location to copy found files to
logfile = os.path.join(newdir,'logfile.csv') #creates a log file under the new directory
fieldnames = ['Filename', 'Path', 'Created Time', 'Modified Time', 'Size', 'File Type', 'Owner', 'Owner Domain'] # headers for logfile csv and organizing data
results = []
file_count = 0


def parse_dirs(rootdir, newdir): # walk through all folders, subfolders in the root directory, and collect metadata
	start_time = now()
	rootfile_size = os.path.getsize(rootdir)
	for subdir, dirs, files in os.walk(rootdir):
		for file in files:
			file_count +=1
			while True:
				try:
					path = os.path.join(subdir, file) #full file path (just adds the root file to the file name)
					[mode, ino, dev, nlink, uid, gid, size, atime, mtime, ctime] = os.stat(path)
					sd = win32security.GetFileSecurity (path, win32security.OWNER_SECURITY_INFORMATION)
					owner_sid = sd.GetSecurityDescriptorOwner()
					ownername, domain, typ = win32security.LookupAccountSid (None, owner_sid)	
					filetype = os.path.splitext(file)[1]
					ctimedate = datetime.fromtimestamp(ctime).strftime('%Y/%m/%d %H:%M:%S') #convert seconds to date-time
					mtimedate = datetime.fromtimestamp(mtime).strftime('%Y/%m/%d %H:%M:%S') # convert seconds to date-time
					# create a dictionary of {fieldnames: metadata} for the file and add to results list
					results.append(dict(zip(fieldnames, [file,path,ctimedate,mtimedate,size,filetype, ownername, domain])))
				except error as e:
					print (e,file)
					time.sleep(15)
					continue

				except OSError as err:
					print (err, file)
					time.sleep(15)
					continue
				
				break
	
	#create logfile, add results
	with open(logfile, 'w', newline='') as csvfile:
		writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
		writer.writeheader()
		writer.writerows(results)
	end_time = now()
	total_time = end_time - start-time
	print (rootdir, "took ", total_time.days, 'days', total_time.seconds, 'seconds', total_time.microseconds, 'microseconds to parse through', //
		file_count, 'files, a total size of', rootfile_size/1000000, 'MB')

def create_new_dir(newdir): #creates the new directory if it doesn't exist already
	if not os.path.exists(newdir):
		os.mkdir(newdir)


def main():
	create_new_dir(newdir)
	parse_dirs(rootdir, newdir)

if __name__ == '__main__':
	main()