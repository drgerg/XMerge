This is XMerge v1.3.6

v1.3.6 Found a good way to prevent the new HTML help text from being scaled up
to huge on high-DPI displays.  The ".call('tk', 'scaling', 1.0)" function
resets scaling to '1.0' for only the help window which works great.

v1.3.5 is the first use of the 'tkhtmlview' module which allows us to move
away from plain text in the Help box and display formatted text, which is
easier to read.  The help file is written in markdown, then translated to
html with the Python 'markdown' package.  The resulting html file is then 
packaged up when the compiler runs.  This also allows us to include images
in the Help file, which is nice.