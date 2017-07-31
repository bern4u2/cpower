import os
import csv
from datetime import datetime
import win32con
from pywintypes import error
import time
import win32com.client
from extract_msg import Message, Attachment

"""NOTES: put in logfile already open error
add a cache
auto-add date to logfile name
"""

exclude = ["S:\\Load Response\\Siebel Documents\\Site documentation\\Middlesex County Regional School District\\MRESC - Old Bridge\\MRESC - Old Bridge High School\\James A McDivitt 4252",
"Work in process - Only for Working - please do not send anything from this folder"]
#a file in MRESC was causing problems, I manually looked through it and none were curtailment plans, exclude 'work in process' folder per Arusyak

def parse_msg(rootdir, logfilename): # walk through all folders, subfolders in the root directory, and collect metadata
	start_time = datetime.now()
	rootfile_size = 0
	file_count = 0
	dirs_count = 0
	logfile = os.path.join(newdir,logfilename +'.csv') #creates a log file under the new directory
	for subdir, dirs, files in os.walk(rootdir, followlinks=True):
		dirs[:] = [d for d in dirs if d not in exclude] #exclude the files in 'exclude'
		dirs_count +=1
		for file in files:
			while True: #keep looping until all files are finished or an error occurs that is not included in the exceptions
				try:
					curtime = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
					path = os.path.join("r'",subdir, file) #full file path (just adds the root file to the file name)
					filetype = os.path.splitext(file)[1]
					if filetype == ".msg":
						print (file_count) # tracking purposes for large directories

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
		
def save_msg_attachments(path):
#https://stackoverflow.com/questions/26322255/parsing-outlook-msg-files-with-python
	#define the COM outlook Object
	outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
	#Call the object method 'open' and define as 'msg'
	msg = outlook.OpenSharedItem(path)
	attachments = msg.Attachments
	for att in attachments:
		att.SaveAsFile(os.path.join(path, att.FileName))



def main():
	rootdir = input('Enter root directory to parse through: ') #location to look through files

	logfilename = input('Name of log file to be stored in above directory (i.e. ISONE_logfile): ')


if __name__ == '__main__':
	main()