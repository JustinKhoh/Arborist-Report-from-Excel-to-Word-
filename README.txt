Welcome to my repository. I hope my code will be useful in creating your Arborist report

(I have little to no coding experience so please bear with me)
###############################################################################################################################
I specifically designed this code to help Arborists convert their Excel table into Word format. Our requirements may be slightly different
but I beleive the code itself can be tweaked to fit your needs (if you know what's going on)

My code is spllit into 2 parts
a) Arborist Excel to Word_Part 1.py (henceforth refered to as Part 1)
b) Arborist Excel to Word_Part 2.py (henceforth refered to as Part 2)

Ensure that both parts are in the same workign directory

Part 1: Convert all rows into individual Word docx
Part 2: Append all Word docx into 1 Word file

NOTE: There is a sample "Tree.xlsx" file and "Photos" folder for a trial run. Missing photos are intentional to test the code
#################################################################################################
The code is done purely on python an requires the installation of several libraries for it to work.
(if on windows machine)

Step 1: Open cmd from your START menu
Step 2: Ensure you have python installed (go to https://www.python.org/)
Step 3: Type in "pip install openpyxl"
Step 4: Type in "pip install pandas"
Step 5: Type in "pip install spire.doc spire.xls"
Step 6: Type in "pip install python-docx"
(If i am missing out any libraries just google what you're missing)
Step 7: Close cmd
Step 8: Install sublime.text (https://www.sublimetext.com/) (any code reader will do but i like its colour scheme and i was taught using sublime)

########################################################################################
PART 1: CONVERT ALL ROWS INTO INDIVIDUAL WORD DOCX

The code is split into several parts, but can be roughly broken down as shown below

Segment 1: Determine Variables
Segment 2: Splitting each row in main table into indivudal sheets, convert each sheet into individual .xlsx document
(The first tree would be named Tree1.xlsx etc.)
Segment 3: Truncate horizontal row in each .xlsx file
Segment 4: Convert individual cell words to Aptos and Font Size 12
Segment 5: Left align cell values
Segemnt 6: Convert each Excel file into word and insert up to 4 images into word table
Segment 7: Replaces watermark with T0XXX [Common Name]
Note that if there is no common name, the code will substitute it as nan, this must be edited manually

#########################################################################################
PART 2: APPEND ALL WORD DOCX INTO 1 WORD FILE

WARNING: RUN PART 1 FIRST BEFORE RUNNING PART 2

The code is split into several parts, but can be roughly broken down as shown below

Segment 1: Determine variables
Segment 2: Create seperate folder to store .docx files
Segment 3: Move .docx files into folder created in segment 2
Segment 4: Appends all docx files into 1 file and save it under the given name

















