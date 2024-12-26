# Combine multiple .CSV, .TXT, .XLS, .XLSX files into one without Excel

In my regular job, we have **tons** of tasks to track, and to help, we use **lots** of spreadsheets. The sheet files are often scattered all over the place across a huge folder structure that makes the task of getting today's data more time consuming than it should be.  This is what spurred me to create **XMerge**.

If you are a seasoned Excel guru, you may not need this.  Even if you are, if you find yourself combining the same spreadsheets into one over and over, or if you need to do that as part of a batch process, XMerge is a handy tool.  The first time you set up a job, XMerge creates a **'.ini'** file where it stores the particulars of this merge job.  From that time forward, each time you point XMerge to that output folder, it will load the settings it finds in that file. You can decide to tweak it, but you don't have to.  This has saved me a **lot** of time.

**Flat data files** means comma-delimited (.csv), tab-delimited (.txt), .xls files, and .xlsx files.  You can argue the last two aren't 'flat files', but XMerge works on them anyway.  Output files can be .csv (comma-delimited), .txt (tab-delimited) or .xlsx formats.

### XMerge does not require Excel, a Microsoft login, the cloud, and Ethernet connection or any outside dependancies.

Choose your settings, select your source files from one or many folders, then hit the 'Go' button.

XMerge copies the source files to a temp directory and works on them there.  The originals are not touched.  The temp directory is deleted when XMerge finishes.

XMerge puts the output file you specified into the folder you specified.

The output file will have a header row that was created one of two ways:

 - from the source files: copying all the column headings. 
 - from a configuration sheet you defined.

If Row 1 or Row 2 is your header, you do nothing. Otherwise you can tell it which row to use.

It's a pretty easy tool.  You **can** use it unattended from a terminal (I use Powershell).  **xmerge.exe -u "path-to-existing-LastXMerge.ini-file"** gets that done.

 - The -t commandline argument generates a .txt (tab-delimited) file.
 - The -c commandline argument generates a .csv output file.  
 - The -x commandline argument prevents the default .xlsx output file.
 - The -u commandline argument runs XMerge in unattended mode, useful for batch files.

XMerge was written in Python.  The GUI is written for Tkinter, nothing fancy.  This was complied into a Windows executable using Pyinstaller, and then Inno Setup Compiler cooks it down into a single .exe installer for Windows.

XMerge has been used for over two years on a regular basis. It has been field-tested extensively, but that doesn't mean every possible scenario has been tried.  You might be trying to use it in a way no one else thought of, so do due diligence and test first.
