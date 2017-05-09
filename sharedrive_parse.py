import os
import fnmatch
import csv
import shutil
import pprint

rootdir = "C:/Users/Bernadette.Werntges/OneDrive - CPower/" #location to look through files
newdir = "C:/Users/Bernadette.Werntges/OneDrive - CPower/test/" #location to copy found files to
logfile = os.path.join(newdir,'logfile.csv') #creates a log file under the new directory
fieldnames = ['filename', 'path'] # headers for logfile csv and organizing data
results = []
keywords = ['*curtailment*', '* cp *'] #list of keywords to look for in file names


def parse_dirs(rootdir, newdir): # walk through all folders, subfolders in the root directory, and copy certain files to new directory
	for subdir, dirs, files in os.walk(rootdir):
		for keyword in keywords:
			for file in fnmatch.filter(files, keyword):
				#if keyword in file.lower(): # if filename contains any of the keywords, copy the file to the new directory
				path = os.path.join(subdir, file) #full file path (just adds the root file to the file name)
				
				try:
					#shutil.copy2(path, newdir) #copies the file (path) to the new directory (newdir)
					results.append({'filename': file, 'path': path}) # add the filename and the full path to a results list
				except IOError as e:
					print (e)
					continue
				except shutil.Error as serror:
					print (serror)
					continue

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