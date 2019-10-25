# LID Validation Logic (Base Version)
# Technologies Used -> Python 3.6.2, OpenPyXL Libraries, PyCharm IDE
# Developed by B B Susheel Kumar
# bb.susheelkumar@yahoo.com

from openpyxl import load_workbook
import xml.etree.ElementTree as ET
import fnmatch
import os
import subprocess

# Enable (1) / Disable (0) Debugging Logs
dLogs = 0


# ----------------------------------> START of LID EXCEL SHEET INITIALIZATION Logic <--------------------------------

# Name of LID Excel Sheet
lidSheetFile = "LIDSheet.xlsx"

# Opening Excel Workbook
lidBook = load_workbook(filename=lidSheetFile, read_only=False)
# Go to Strings WorkSheet in the Excel File Opened
lidSheet = lidBook['Strings']
# Read only Column "B" Cells (Cell Address)-> LID's, So lidCol contains list of all LID's in the Excel Sheet
lidCol = lidSheet['B']
# Read only Row "2" Cells (Cell Address)-> Languages, So langRow contains list of all Languages in the Excel Sheet
langRow = lidSheet[2]

# Extract the data in langRow cell addresses to langList (Stores the list of Languages Supported in LID Excel Sheet)
langList = []

# ----------------------------------> END of LID EXCEL SHEET INITIALIZATION Logic <----------------------------------


# Row index of LID Excel Sheet
rowCounter = 0
# Counter which gives the number of LID's in XML which are same as in LID Excel Sheet
matched = 0

# Counter which gives the number of LID's in XML which are found in LID Excel Sheet
foundCounter = 0


# -------------------------------------> START of LID EXCEL SHEET SCANNING Logic <------------------------------------
# Extracting the data in langRow cell addresses and adding it to langList
# to make a list of all the available languages in the given LID Excel sheet
for cell in langRow:
    if cell.value is not None:
        langList.append(cell.value.strip())

# Path of the Project/Folder/File to be parsed (Default: Current Working Directory)
my_path = os.getcwd()

# Stores the paths of all strings.xml files in a given Project/Folder/File
strFileList = []
# Stores the language of the respective strings.xml file
# Index of strFileList => Index in strFileLang:: Language of strings.xml file in the path as in the index of strFileList
strFileLang = []

# Contains the Column's related to various languages. (For Korea => Column 'K' which is nothing but 11th Column)
# valueList = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
#              'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ']
valueList = [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26,
             27, 28, 29, 30, 31, 32, 33, 34, 35, 36]
# -------------------------------------> END of LID EXCEL SHEET SCANNING Logic <--------------------------------------


# -----------------------------------------> START of FILE SCANNING Logic <-------------------------------------------
# Logic to walk recursively through various folders, files in the given path (my_path)
for root, dirs, files in os.walk(my_path):
    if dLogs: print("Root: %s, Dirs: %s, File: %s" % (root, dirs, files))
    valueDir = fnmatch.filter(dirs, "values*")
    for filteredDir in valueDir:
        if dLogs: print(dirs)
        if filteredDir in langList:
            dirFiles = os.listdir(os.path.join(root, filteredDir))
            xmlfiles = fnmatch.filter(dirFiles, "*strings.xml")
            if dLogs: print(xmlfiles)
            count = 0
            while count < len(xmlfiles):
                strFileLang.append(langList.index(filteredDir))
                count += 1
            strFileList.extend(os.path.join(root, filteredDir, f) for f in xmlfiles)
# -----------------------------------------> END of FILE SCANNING Logic <-------------------------------------------


# Prints the list of column number in the Excel Sheet to be checked for the respective strings.xml
if dLogs: print("\n Columns to be checked: "+strFileLang+"\n")

# Counter to iterate through the list of path's having strings.xml
fileCounter = 0


# ---------------------------------------------> START of MAIN Logic <------------------------------------------------
for strFile in strFileList:
    if dLogs: print("File Number: "+fileCounter+" Related Column in LID Sheet: "+valueList[strFileLang[fileCounter]])
    if dLogs: print("File Path: "+strFile+" "+"\n")
    # sheetColIdx gives the column in LID Excel Sheet to be checked for the respective language
    sheetColIdx = valueList[strFileLang[fileCounter]]
    # Parsing the strings.xml file as an Element Tree (DOM Tree)
    eleTree = ET.parse(strFile)
    # Get the root element of the XML file (root is generally <resources> tag)
    treeRoot = eleTree.getroot()
    # Listing all the <string> elements
    strEle = eleTree.findall('string')

    # Logic to Validate the XML LID with LID in Excel Sheet
    # Get the LID's in XML and check them in the Excel Sheet for respective Values
    # Compare values of respective LID's between XML and Excel File
    # If not similar then replace, else skip to next LID in strings.xml
    for cell in lidCol:
        # LID's start after second row in the LID Excel Sheet
        if rowCounter >= 2:
            for lidTag in strEle:
                # Get the LID Attribute from <string> tag, it gives LID Number
                lidName = lidTag.get('LID')
                # NULL Check for Cell Value in Excel and LID Number in XML
                if cell.value is not None and lidName is not None and lidName.startswith('LID'):
                    if cell.value.strip() == lidName:
                        if dLogs: print("FOUND : ", lidName)
                        # Get the respective LID Value for a LID Number from XML
                        lidValue = lidTag.text
                        # While comparing the values in the XML, we need to remove double quotes at the Start and End
                        if lidValue.startswith('"') and lidValue.endswith('"'):
                            lidValue = lidValue[1:-1]
                        if dLogs: print("XML : ", lidValue)
                        # ----------> START: Get the respective LID Value for a LID Number from EXCEL <---------------
                        sheetValue = lidSheet.cell(row=rowCounter + 1, column=sheetColIdx).value
                        # In case of LID value which has Multiple Lines in the Excel Sheet, split them as lines
                        lidLines = sheetValue.splitlines()
                        # Join the lines split as a single line (Multiple Line LID's should have '\n' embedded in it
                        sheetValue = "\\n".join(lidLines)
                        # ----------> END: Get the respective LID Value for a LID Number from EXCEL <---------------
                        if dLogs: print("EXCEL : ", sheetValue)
                        if sheetValue == lidValue:
                            matched += 1
                            if dLogs: print("MATCHED")
                        else:
                            if dLogs: print("NOT MATCHED")
                            # If not matched, replace the value in strings.xml with Excel Sheet Value.
                            lidTag.text = '"'+sheetValue+'"'
                            if dLogs: print("CHANGED in XML")
                        if dLogs: print("\n")
                        foundCounter += 1
        rowCounter += 1
    fileCounter += 1
    rowCounter = 0
    # Write and Save the strings.xml file
    eleTree.write(strFile)
# ---------------------------------------------> END of MAIN Logic <------------------------------------------------


# ---------------------------------------------> START of RESULTS <------------------------------------------------
print("\nLID's FOUND : %d" % foundCounter)
print("LID's MATCHED : %d\nLID's NOT MATCHED : %d" % (matched, (foundCounter - matched)))
# ----------------------------------------------> END of RESULTS <-------------------------------------------------

print('Enter "push" command to upload changes to GIT')

