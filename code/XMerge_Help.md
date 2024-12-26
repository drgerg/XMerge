## XMerge: The Source File Merger

**XMerge** helps you merge the contents of multiple .txt, .csv, .xls, or .xlsx files into one .xlsx spreadsheet.

You specify or create the Output folder.  This is where your output file will be when XMerge finishes.  All of XMerge's activities happen in this folder.

XMerge makes copies of source files and creates your output.xlsx file.  Because it works from copies, any mistakes you make in selecting options will do no permanent damage.  When things don't go like you wanted, change settings and try again.

**XMerge does not manipulate the original source files in any way.**

#### UNATTENDED MODE

There are four command-line arguments (switches, options) you can use:  
`xmerge -u "c:\path\to\LastXMerge.ini\file\"` runs XMerge in unattended mode. This is useful in batch files.  
`xmerge -c` creates a .csv output file in addition to the normal .xlsx file.  
`xmerge -t` creates a .txt tab-delimited file in addition to the normal .xlsx file.  
`xmerge -x` does NOT create the normal .xlsx file.  

For example: If you want to run unattended, and end up with only a .csv file, you would run:  
`xmerge -u "c:\path\to\LastXMerge.ini\file\" -c -x`

Unattended mode does not modify the LastXMerge.ini file. It references it for some values, but leaves it intact.

#### CHOOSE YOUR COLUMN NAMES

![The controls section.](.\\img\\controls.png)

There is a radio button section in the Controls block that allows you to select between "All Columns" or "Configured Cols".  If you want to specify the columns XMerge will pull into the new output .xlsx file, you do that by listing the column names in a spreadsheet called 'ColumnNames.xlsx'.

If you select the 'Configured Cols' option, ColumnNames.xlsx is added to this job's Output Folder.  The file is only used if "Configured Cols" is checked.

The  file must have at least one sheet (tab) named "MERGE" (no quotes).  There is some explanatory text on that sheet as well.

Replace whatever you find there with the column names you want to be in your new merged spreadsheet.  You can backup name sets by making a new tab and saving previous column names in it.

**NOTE**: The column names you specify must already be in each of the source files you are merging.  If not, XMerge will create the columns in your merge file, but there will be nothing in them.

#### SPECIFY THE HEADER ROW

People use the top of their spreadsheets in many different ways.  Some are plain, just a row of column labels (aka, the 'header'). Others have ornate formatting on multiple rows above the header, which they intend to be seen on reports.

As a result, XMerge may need your help understanding which row is actually the header row in your source files.  

'auto' understands 1 or 2 rows.  If your source files have more than two rows in the header (above the first row of data), you will need to specify which row is the header row (the row with column labels or names).

To do that, replace the word 'auto' with a single number, for example: '4'.

#### SOURCE FILES MISSING A COLUMN

If a source file doesn't have an expected column, XMerge just fills cells for that column with empty strings.  No fuss, no muss.

#### OUTPUT FILENAME

You may specify the output filename in a box on the screen, or leave the default "XMerge_Export" there.
No file extension is expected or needed.

You can specify new output folders anytime you wish, and you can change the output filename anytime you wish.

#### DUAL CONFIGURATION FILES (xmerge.ini and LastXMerge.ini)

**XMerge stores your current configuration in two .ini files.**

- The file **xmerge.ini** lives in the installation folder.  This is the **system configuration file**. The values contained in this system .ini are used for repeatedly working in the same output folder. The last folder you worked in will be where you start when you return to XMerge after exiting.
- The file **LastXMerge.ini** lives in each output folder you specify.  This file contains a record of the source file paths and filenames.  This allows you to pick up where you left off last time you worked in any specific output folder.  Each time you select that output folder, the LastXMerge.ini file is read.  All you have to do is click 'Go' to replicate whatever happened last using that output folder.

#### What happens next

XMerge creates a "temp" folder in the Output folder, and it is this folder into which the source files are copied.  Then XMerge converts them to .xlsx and copies the contents of each file concurrently into the new output file.  Once the conversion and copying is finished, XMerge removes the "temp" folder and the files that were copied into it.

The original source files are never touched.

### CHECKBOXES AND BUTTONS

**CHANGE OUTPUT FOLDER**:  Select this checkbox when you are ready to start a new merge project.  This box is checked by default the first time you run XMerge.

**CLEAR OUTPUT FOLDER**:  Select this checkbox if you want to delete the contents of the current outputs folder before continuing.  Be aware: all files will be deleted from the current output folder (the one listed on line 3 "Currently selected Output folder:".)

**APPEND SRC FILENAMES**: If you check this box, the source file filename (minus the extension) will be added to the last column in your spreadsheet.  This tells you where that row of data came from.

**EXPORT TO .CSV FILE** Create a comma-delimited output file.

**EXPORT TO .TXT FILE** Create a tab-delimited output file.

**EXPORT TO .XLSX fILE** Create an .xlsx output file. This is the default, so it's checked.

**'GET NEW DATA' BUTTON**:  Use this when you are first starting XMerge, or anytime you want to change data sets.

**'GO' BUTTON**: Use this to copy, convert, and compile the same files you worked with last time.  This is helpful when you have made edits and re-generated the source data files.

### COMMAND-LINE OPTIONS

Typing: xmerge -u "C:\Users\me\Desktop\myspecialfolder\" runs XMerge in unattended mode.  This can only be used to repeat previously run merges.

Typing: xmerge -c results in the creation of a comma-delimited version of your output (Export) file.

Typing: xmerge -t results in the creation of a tab-delimited version of your output (Export) file.

Typing: xmerge -x prevents the creation of an .xlsx output (Export) file.

#### FUZZY FONTS ON HI-DPI DISPLAYS

High-DPI displays introduce the potential for fuzzy fonts due to scaling.  This behavior is determined by your Display Settings in Windows.

If you see anything other than 100% here:

![Display Settings](.\\img\\hi-dpi-3.png)

then you are scaling the desktop on that display.  This may cause what looks like slightly out-of-focus fonts in XMerge.

These fuzzy fonts are a Tkinter thing.  It bothers some people.  

There is one way to improve things somewhat.  Your mileage may vary.

- Look at the File|About dialog in XMerge to find your installation folder
- Navigate to that folder and find XMerge.exe
- Select and right-click on XMerge.exe.  Select **Properties**.
- Look for the **Compatibility** tab and play with these settings.

![The Properties dialog.](.\\img\\hi-dpi-1.png)
![Scaling override settings.](.\\img\\hi-dpi-2.png)

**XMerge** is written in <span style="color:red;">Python</span> and complied using <span style="color:red;">Pyinstaller</span>.  The installer is built for Windows using the <span style="color:red;">Inno Setup Compiler</span>.

**Many thanks** to <span style="color:green;">bauripalash (Palash Bauri)</span> for [tkhtmlview](https://github.com/bauripalash/tkhtmlview), which makes this help file window look so good!
