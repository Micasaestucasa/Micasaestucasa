# Micasaestucasa

PROJECT DESCRIPTION
Group Project ID: 2269 - This project uses Python in conjunction with Microsoft Excel to run a webscraper, download a CSV file from a scraped online database (Inside AirBnB),  and run some high level data analysis on the downloaded CSV file. The final analysis is displayed in an "input" excel file that is automatically created and used as a source of input for the webscraper. All downloaded CSV files and input files are also automatically created and included in a specific folder in the desktop directory of the user for ease of use)

WHO IS BEHIND THIS PROJECT?
This is a mandatory group project part of the course "Programming with Advanced Computer Languages" supervised by Dr. Mario Silic.

The group project was done in a group of two people: 
- Antoine Rabilloud (Student-ID:***********) with the pseudonym "Rabibi"
- Giovanni Grosso (Student-ID: ***********) with the pseudonym "BigJohn"

SHOPPING LIST TO ENSURE THE PROGRAM RUNS
Setting up to run the code:
In order to run the project the user will need: 
1- Anaconda3
  a. Sypder (running through Anaconda3) which will used as the main tool to run our code
2- Microsoft Excel (used as a way for the users to input the type of data they want downloaded from the online database)
3- A decent internet connection

MAIN LIBRAIRIES USED
The coding was done using Spyder from which we imported several python librairies/packages listed bellow: 
Openpyxl - used to read our input excel file for the webscraper and used to write on the final created excel file with our analysis
Beautiful Soup - used for parsing HTML and XML documents (webpages). was used to extract data from HTML (our csv data file)
Requests - Used in conjunction to Beautiful Soup to get the specifc website we want to analyse
re -  Used to specify what strings of the HTML page to look out for (based on our excel input file) to extract the right CSV on the online database
Pandas - Was used in conjunction with Numpy for our data manipulation and analysis (specificaly when creating numerical tables and time series)
Numpy - helped us analyse the arrays of data that we downloaded as a csv and used to perform high-level mathematical functions that we operated on these arrays

IDEA OF THE PROJECT / SITUATION
The goal was to give simple analysis tools to a user who does not specifically know how to code (the idea being to recreate a real life work environment where you the "coder" has to adjust to the technical limits of the people he/she needs to deliver a working solution to).

The file was then modified to only cover the last accessible years (2007 - 16) and eliminated pack of countries (European Union etc.) to only have individual countries.

The user can first choose one of the 162 countries and get several information (such as year with max. NNI and its value, year with min. NNI and its value, average of its NNI over 2007-16, median of its NNI over 2007-16, as well as the CAGR of its NNI's development over 2007-16).

The user can then choose an individual year (2007-16) to display other information (such as the country with the max. NNI and its value in that year, the country with the min. NNI and its value in that year, the mean NNI value of all countries in that year, as well as the median NNI value of all countries in that year).

NNI stands for Net National Income and is an economics term used in the accounting of national income. It is the difference between what is earned by nationals living inside and outside the country put together and non-nationals living in the country. It can be defined as NNI = C + I + G + (NX) + net foreign factor income - indirect taxes - manufactured capital depreciation where C = Consumption, I = Investments, G = Government spending and NX = Net exports (exports minus imports).

GUIDE TO OUR GITHUB REPOSITORY
Four files are distributed within this repository.
This file (README.md) is used for a project description as well as insutrctions on how to use and test the code done by the student.

Two files contain code ________________________________________________ is the file that is binding and shall be tested and used by the professor, _____________________________________ only provides the professor with an example of how the student functioned in the development of his code writing and testing.

A last file ____________________________ is a modified file taken directly from the INSIDE AIRBNB website (shared above) which is read in the ____________ file.

The file containing the code ____________________ also contains extensive comments for all or most of the coding lines. The comments should be read to gain deeper insight into the functioning of the code and logic used.

TESTING THE PROJECT
To test the project: The excel file (NNI_Index.csv) and (AF_proj.py) have to be downloaded and put in the same folder. The user then has to make sure she/he has functioning and latest python3 as well as pandas and NumPy libraries. By using the terminal, the user then open the file by accessing the folder (writing "cd ") and by sliding the folder which contains the files in the terminal and by then writing (AF_proj.py) if that is still the code file's name. Displayed will be an exemplary of the first 10 countries (alphabetically) in the list with corresponding values in some years. The user will then be prompted to enter a country's name (refer to line 19 - 22 of this file for explanations on what the code asks for inputs and expected outputs).

We hope you will enjoy this new project, 

Best Regards,

Rabibi and BigJohn
