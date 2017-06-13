Excel-SelfDestruct
==================
A simple VBA script injection which causes an Excel Spreadsheet to close after a fixed period of time.

This script is designed for large companies with a central repository of Spreasheets which need to be closed so the next person can work on them.


Features:

* Activity detection - timer gets reset each time the user moves around the spreadsheet or interacts with any shett
* Closure Message - Closes the spreadsheet _and then_ displays a message
* One file - All code is contained within one 'BAS' file which can be inserted into 'ThisWorkbook' without any other dependencies


Installation
------------

1. Change your Excel Sheet (probably `.xlsx` extension) into an 'Excel with Macros' file using "Save As" (`.xlsm` extension).

2. Enable the VBA / Developer mode if you havn't already. [Instructions from Microsoft](https://msdn.microsoft.com/en-us/library/office/ee814737(v=office.14).aspx#Anchor_2)

3. Open the "Visual Basic" editor (Under the 'Developer Ribbon', 'Visual Basic')

4. Double click on `ThisWorkbook` in the project editor

5. Paste the [code](self-destruct.bas) into the empty code window that opens

6. Save and close the document
