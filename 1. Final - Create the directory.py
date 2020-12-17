# This section of code automatically creates the necessary folders in the computer's directory (desktop) for the data to be stored

# Get the user's computer username for the directory
import getpass
username = getpass.getuser()
import os 

# Create Micasa folder
main_dir = ("C:/Users/%s/Desktop/Micasaestucasa" %username)
os.mkdir(main_dir) 
print("Directory '% s' is built!" % main_dir) 

# Create working file folder
main_dir = ("C:/Users/%s/Desktop/Micasaestucasa/working file" %username)
os.mkdir(main_dir) 
print("Directory '% s' is built!" % main_dir)

# Create export folder
main_dir = ("C:/Users/%s/Desktop/Micasaestucasa/export" %username)
os.mkdir(main_dir) 
print("Directory '% s' is built!" % main_dir)

