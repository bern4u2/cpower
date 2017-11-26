import csv
import shutil
import os

def copy_files_from_csv(read_dir, copy_to_dir)
	for subdir, dirs, files in os.walk(read_dir): #go through each file in the user-defined directory
		for file in files:
			if os.path.splitext(file)[1] = '.csv': #look for .csv files only
				base_path = os.path.splitext(file)[0]
				newdir = os.path.join(copy_to_dir, base_path) #new directory name will be the filename within the user-definied main directory
				with open(file, 'r', newline='', encoding='utf-8') as csvfile:
						reader = csv.DictReader(csvfile, fieldnames=fieldnames)
						for row in reader:
							shutil.copy2(row['Path'], newdir) #copies the file to the new directory (newdir)

def main():
	read_dir = input('Enter the directory path where read-files are located: ') #location to look through files
	copy_to_dir = input ('Enter directory path to copy files to: ') #main directory to copy the files to
	create_new_dir(copy_to_dir)

	copy_files_from_csv(read_dir, copy_to_dir)

if __name__ == '__main__':
	main()