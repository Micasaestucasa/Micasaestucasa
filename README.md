# Micasaestucasa

## GUIDE TO OUR GITHUB REPOSITORY
Four files are distributed within this repository.
This file (README.md) is used for a project description as well as insutrctions on how to use and test the code done by the student.

Two files contain code ________________________________________________ is the file that is binding and shall be tested and used by the professor, _____________________________________ only provides the professor with an example of how the student functioned in the development of his code writing and testing.

A last file ____________________________ is a modified file taken directly from the INSIDE AIRBNB website (shared above) which is read in the ____________ file.

The file containing the code ____________________ also contains extensive comments for all or most of the coding lines. The comments should be read to gain deeper insight into the functioning of the code and logic used.


## PROJECT DESCRIPTION

**Group Project ID: 2269**
- **This project uses Python in conjunction with Microsoft Excel** to run a webscraper, download a CSV file from a scraped online database [Inside AirBnB](http://insideairbnb.com/get-the-data.html),  and run some high level data analysis on the downloaded CSV file. The final analysis is displayed in an "input" excel file that is automatically created and used as the MAIN source of input for the webscraper. All downloaded CSV files and input files are also automatically created and included in a specific newly created folder in the desktop directory of the user for ease of use.


###### **WHO IS BEHIND THIS PROJECT?**

**This is a mandatory group project part of the course "Programming with Advanced Computer Languages" supervised by Dr. Mario Silic.**

The group project was done in a group of two people: 
- Antoine Rabilloud (Student-ID:***********) with the pseudonym "Rabibi"
- Giovanni Grosso (Student-ID: ***********) with the pseudonym "BigJohn"


###### **SHOPPING LIST TO RUN THE PROGRAM**

**In order to run the project the user will need:**
 1. A computer (yes, yes...)
 2. Anaconda3
     - Sypder (running through Anaconda3) which will be used as the main tool to run our code
 3. Microsoft Excel (used as a way for the users to input the type of data they want downloaded from the online database)
 4. A decent internet connection


###### **MAIN PYTHON LIBRAIRIES USED**

**The coding was done using Spyder from which we imported several python librairies/packages listed bellow:**
- Openpyxl - used to read our input excel file for the webscraper and used to write on the final created excel file with our analysis
- Beautiful Soup - used for parsing HTML and XML documents (webpages). It was used to extract data from HTML (our csv data file)
- Requests - Used in conjunction to Beautiful Soup to get the specifc website we want to analyse
- re -  Used to specify what strings of the HTML page to look out for (based on our excel input file) to extract the right CSV on the online database
- Pandas - Was used in conjunction with Numpy for our data manipulation and analysis (specificaly when creating numerical tables and time series)
- Numpy - helped us analyse the arrays of data that we downloaded as a csv and used to perform high-level mathematical functions that we operated on these arrays

###### **IDEA OF THE PROJECT / SITUATION**
**The goal was to give simple analysis tools to a user who does not specifically know how to code** (the idea being to recreate a real life work environment where - you, the "coder" - has to adjust to the technical limits of the people he/she needs to deliver a working solution to).
once the data is downloaded the person should be able to run further analysis on excel through the files that were created. ( - you, the coder - generates an analysis that helps 80% of the people that will use the program but give the possibility to the users to make further analysis (on excel!) if desired). 


###### **THE DATABASE/WEBSITE**
**To run this program we go on the [INSIDE AIRBNB](http://insideairbnb.com/get-the-data.html) website.**
The webiste gathers monthly data from AIRBNB listings all over the world. When on the website, several Excel and CSV files with different data points are available. We decide to use the "listing.csv" files for our analysis as they were identified as the best for data visualizations by the website itself and seemed to contain the most interesting datapoints. 

### **TESTING THE PROJECT**
###### **SET-UP**
**To run the project please:**
- First ensure that you have all the "shopping list items" needed to run the code. (This is highlighted in the **SHOPPING LIST TO RUN THE PROGRAM** section of this ReadMe document. 
- Second, look for our code on our Github page  (available in the __________________ file of our directory)
- Third Copy and paste the code in the SPYDER notebook and run the program

**We detail below the different segments of our project and provide you with a detailed decscription of what should happen what other things the user can do to further test our code and "play around".**

###### **INPUT FILE**
**We use an input excel file as our main source of input to select the specific city/country, year and month that interests us for our analysis (for more clarity only a select number of cities are available for selection)**
For a smooth test we have also decided **to already prefill the input excel file so that the user can download some data to test immediately without having to input anything in the file itself.** Should the user want to analyse data for another city or month, he/she can input the variables for the webscraper directly through the created excel.

###### **CODE EXPLANATION**
**The code we created should:**
1. Set up a directory on the user's desktop where a "micasaestucasa" folder will be automatically created (with subfolders for the different excel and CSV files used). 
2. An Excel file used for input will be automatically created and prefilled with needed variables to run the webscraper. The excel also contains a detailed how-to guide so that any person wanting to change the scope of the analysis (city/country, year, month) can do so whilst ensuring that the code continues running smoothly. 
3. Once the excel file is created the code takes the prefilled input OR the user's new input (in the input file) and starts scrapping the web to get the right CSV file for our analysis. (NOTE that this might take up to a minute depending on the connection)
4. Once the right file is found by the web scraper, it is downloaded and put in an export subfolder located in the main "micasaestucasa" folder on the user's desktop
5. The code renames the file to detail what kind of data is included in the file downloaded (eg: London 03.12)
6. The program then saves the CSV as an Excel for ease of use and to avoid any errors
7. The program then uses pythons capabilities (pandas, numpy librairies specifically) to analyse the big data base (WAY faster than excel)
8. the program returns the results of the analysis on the SPYDER dashboard AND directly on the file used for input

![Step 1.](https://github.com/Micasaestucasa/Micasaestucasa/blob/main/step1.jpg)
![Step 2.](https://github.com/Micasaestucasa/Micasaestucasa/blob/main/step2.jpg)


We hope you will enjoy this new project, 

Best Regards,

Rabibi and BigJohn 
