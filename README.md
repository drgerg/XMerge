# Convert and combine multiple flat data files into one without Excel

If you are an Excel guru, you don't need this.  But if you aren't, and you have several flat data files that you need to quickly and simply compile into one .xlsx spreadsheet file, XMerge can help.

**Flat data files** means comma-delimited (.csv), tab-delimited (.txt), .xls files, and .xlsx files.  You can argue the last two aren't 'flat files', but XMerge works on them anyway.

Choose your settings, select your source files (must live in same folder), the hit the 'Go' button.

XMerge copies the source files to a temp directory and works on them there.  The originals are not touched.  The temp directory is deleted when XMerge finishes.

XMerge puts the output file you specified into the folder you specified.

The output file will have a header row that was created one of two ways:

 - from the source files: copying all the column headings,
 - from a configuration sheet you defined.

It's a pretty easy tool to use.  You can use it unattended from a terminal.  **xmerge.exe -u "path-to-existing-LastXMerge.ini-file"** gets that done.

 - The -c commandline argument generates a .csv output file.  
 - The -x commandline argument prevents the default .xlsx output file.  

XMerge was written in Python.  The interface is Tkinter, which is an adequate GUI platform but won't win any beauty contests.  This was complied into a Windows executable using Pyinstaller, and then Inno Setup Compiler 6.2.0 cooked it down into a single .exe installer.

As with all software, you are responsible to make sure it works properly on your files before using it in any production environment.  XMerge has been tested extensively, but that doesn't mean every possible scenario has been tried.  You might be trying to use it in a way no one else thought of, so test first.
