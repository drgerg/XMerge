This is XMerge v1.3.5

v1.3.4 Added a new checkbox labeled 'Append Src Filenames'.  Selecting this box
appends the source filename (minus extension) to each row of the output
sheet.  This allows you to see where the data on that row came from.

v1.3.5 is the first use of the 'tkhtmlview' module which allows us to move
away from plain text in the Help box and display formatted text, which is
easier to read.  The help file is written in markdown, then translated to
html with the Python 'markdown' package.  The resulting html file is then 
packaged up when the compiler runs.  This also allows us to include images
in the Help file, which is nice.

  
