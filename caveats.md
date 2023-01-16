## Function Over Form: Python and tkinter

I wrote this application to solve a problem I had.  I periodically needed a quick and easy way to combine a bunch of tab-delimited files into a spreadsheet.

The app would work great as a command-line app, but some people really prefer a gui, so I just recycled one I had already developed using tcl/tkinter.  

Tkinter is not the most beautiful interface in the world, but it's adequate.  On a high-DPI display, it suffers the artifacts of scaling, which tends to show up as fuzzy fonts.

There are only a few things you can do about that, none of which are really great.

In the end, as far as I'm concerned, function is more important than form, so I'm satisfied with the app as it stands.

Finding the **tkhtmlview** module was a bonus, because now at least the help file could have formatted text.  That is a huge improvement over plain text.  Thanks to [bauripalash](https://github.com/bauripalash) for his work on that!