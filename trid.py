import os
import subprocess
import csv

trid_run = 'C:/Users\Bernadette.Werntges/Downloads/trid_w32/trid.exe'
fieldnames = ['filepath', 'trID_result']


def trid(trid_run, mystery_files, fieldnames): #goes through subdirs, files under the path: 'mystery files', and runs the trID program http://mark0.net/soft-trid-e.html
	result_array = [] #define empty array for results
	file_count =0 #add file counter
	for subdir,dirs,files in os.walk(mystery_files): #walk through all subdirs, files
		for file in files:
			if os.path.splitext(file)[1] == "" and not os.path.splitext(file)[0].startswith('.'): #if there is no file extension, do stuff
				file_count +=1
				filepath = os.path.join(subdir, file)
				extenstion_guess = subprocess.check_output(trid_run + " " + os.path.relpath(filepath, start=os.curdir), shell=True) #run trid with subprocess module
				print (extenstion_guess)
				result_array.append(dict(zip(fieldnames, [filepath, extenstion_guess]))) #add filepath and the trid results to results array
	return result_array

def csv_write(result_array,fieldnames, results_file): #write results to .csv
	with open(results_file, 'w', newline='', encoding='utf-8') as csvfile:
		writer = csv.DictWriter(csvfile, fieldnames)
		writer.writeheader() # headers are the fieldnames
		writer.writerows(result_array) #write the results array to a csv file

def main():
	os.chdir("OneDrive - CPower\ELMO\S drive file info\zipfiles") # trid throws error when a dash is in the file name/path, changing directory first allows us to just call the subdir and avoids the error
	#mystery_files = "C:\\Users\\Bernadette.Werntges\\OneDrive - CPower\\ELMO\\S drive file info\\zipfiles\\Salesforce_logfile\\WE_00D40000000IXNbEAO_1\\Attachments"
	mystery_files = input('Path name to parse through where file extensions are missing: ')
	results_file = input('Name for results file (it will be located at OneDrive - CPower\ELMO\S drive file info\zipfiles): ')
	csv_write(trid(trid_run, mystery_files=mystery_files, fieldnames=fieldnames),fieldnames = fieldnames, results_file = results_file)
	
	

if __name__ == '__main__':
	main()