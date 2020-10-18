Master BITSian Data Store
======================

This is a framework which aims to build an entire BITSian Directory.  Currently, we have different sources of data - member registrations,  campus student data, Chapter based data, and so on. The data, which  can be varying a great amount in detail, in Excel sheet format(.xlsx)  is loaded into the Master BITSian database to build such a BITSian directory.



Pre-requisites
======================
  * User/Server must have python 2.7 installed with the following 
  python packages :  
    * openpyxl
    * sqlite3
    * Levenshtein 




Instructions to Use
======================
To make sure that all the commands go as detailed below, 
open the terminal/command prompt inside the directory 
where source file (BITSAA.py) is present. Make sure you 
load your data file in this directory, or provide the 
exact path of the files while executing the instructions.

1. Setup the columns of Master BITSian Database. This has 
to be run once only, until and unless new columns need to 
be added to the master database. Admin needs to provide 
a csv file with column names of BITSAA Master database. 

  Command : `python BITSAA.py -s [filename]`  OR   `python BITSAA.py --setup [filename]`
  
  * [filename] : must be in csv format.
  * The csv file must have the format with first column as 
  index number, second with names of the actual columns in 
  master database, third for priority(0)/date(1) categorization, 
  fourth and fifth for threshold. It's recommended that you 
  use the columns.csv provided with the project for the time being.

  Example : `python BITSAA.py -s columns.csv`

2. Creating the labels for the data. Data has to be provided
 in the .xlsx format. This basically creates a csv file with
  those columns of the data, which can be mapped to BITSian 
  master database.

  Command : `python BITSAA.py -l [filename] [row number at which column names are written in data file] [output csv file name(optional)]`

  Example : `python BITSAA.py -l Data1.xlsx 1 columns.csv`

  * If the name of output file is not provided, it generates
   the output file with the name of datafile, in csv format.
  * Now, the user can check this csv file to see which columns
   got mapped correctly, and if satisfied go on to the next step. 
   Else, if some column didn't get mapped, user can open the data
    file and change that column's name to something more natural.

    **  Example : A column with name 'NA' or column left blank 
    contains phone numbers may not appear in csv file. In this 
    case, the user can open the data file, and change the column 
    name to "Phone" or "Phone numbers" or for that case, anything 
    that would naturally represent the data, and run step 2 again. 
    The program predicts the mapping, and if written in natural 
    language, and a corresponding column exists in the Master Database, 
    then that column will most probably appear in the csv file created 
    through this step.

3. Adding(inserting) the data with priority and date.
  
  Command : `python BITSAA.py -i [filename] [priority] [date] [row start number] [labels csv filename]`
    * [priority]            : Priority of the data being added. 
    Priority can vary form 0 to 9, where 0 represents the highest priority.
    * [date]                : should be given in 'yyyy-mm-dd' format (without the quotes).
    * [row start number]    : the row number in the excel sheet 
    from which the data starts.
    * [labels csv filename] : (Optional) name of the csv file 
    generated through step 2. This can be avoided if a custom 
    filename wasn't used in step 2 for the output.

  Example : `python BITSAA.py -i Data1.xlsx 5 2013-12-05 2 columns.csv`

4. Getting back the csv file with Best data based on priority 
and date of data inside Master Database. Say if for a record, 
i.e. for a BITSian, a list of phone numbers is there(Added 
through different data file at different times) in the master 
database, then while outputting distribution sheet, it'll 
give the highest priority/latest phone number for that BITSian, 
based on if 'phone number' column is priority based or time based.

  Command : `python BITSAA.py -xd [output filename]`
  
  Example : `python BITSAA.py -xd distribution.csv`

5. Getting back the data files. This kind of behaves like a 
backup, and outputs all excel files that were inserted into 
master database at one time or another. 

  Command : `python BITSAA.py -o [Directory name]`

    * [Directory Name] : is optional. If not provided, a 
    directory with 'Backup' name is created in the program directory.

  Example : `python BITSAA.py -o Backup` 
    will create a directory called Backup with all the .xlsx 
    files that were once used for inserting data inside master 
    database.

6. Wiping the database. This automatically also runs step 5. 
It also mails the list of admins that a wiping action on the 
database has been performed. And, then deletes the master database.

  Command : `python BITSAA.py -p`




Modules and their functions(in Alphabetical Order)
====================================================
A detailed documentation of the modules made is provided 
in the official documentation, a link to which can be found 
below. However, an abstract summary of what each module does 
is being discussed here :
  
  * `addToDatabase`     : adds the data from .xlsx file to the database
  * `compare`           : compares two data in the same column, 
  if they are almost identical.
  * `createRecords`     : creates records in the database
  * `createLabels`      : Based on the excel sheet provided, it 
  picks up the column names and predicts the corresponding 
  matching column name of database, to create a csv file with 
  the matched column names for the excel sheet data provided.
  * `dictToList`        : converts a dictionary to a list.
  * `getDataset`        : extracts data from the excel sheet in 
  the form of list of dictionaries.
  * `getDist`           : gives the excel distribution file, 
  containing the best possible BITSian data based on priority 
  and date of the data existing inside the master database.
  * `getMaster`         : gives the master database as a csv file.
  * `getOriginalFiles`  : gives back the .xlsx files that were 
  used to put in data in the database.
  * `maintainance`      : regularly checks for redundancy in the 
  data by using compare module, and removes if any.
  * `sendemail`         : sends an email to the list of admins 
  giving them the updates regarding major changes to the database.
  * `setup`             : creates the master bitsian table with 
  columns given in the columns.csv file passed to it.

  * `main`              : runs the program based on the option used.




Documentation
======================
The documentation is done using Sphinx tool. Those who want 
to develop this further can look upon this doc as a precursor. 
It is available in the '/proj_docs/_build/html/' directory as 
the 'index.html' file.




Known Bugs/Restrictions
============================
  * The data is currently accepted in Excel Sheet format(.xlsx) only. 
  * Documentation yet to be updated.
  * Distribution sheet, for now is being returned in the csv 
  format instead of .xlsx format. This will be fixed in the next update.
  * The option '-m' (maintainance module) doesn't work properly 
  for the time being. A patch for the same is almost ready. 
  Integration will be done soon. Because of this, right now, '-xd' 
  and '-xm' will basically return the same database, that is, the 
  master database.
  * The email to admins happen only when user tries to purge the 
  master database, for now. It is not yet associated with any other action. 
  * The email, for now is using gmail smtp server and a dummy 
  gmail account. This can be replaced by any other smtp server.




Authors
=============
All of us were students at BITS Pilani University (Goa Campus). 

Soumya Dipta Biswas
Bhavul Gauri
Rohan Saxena

