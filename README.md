# Year 2 Summative 4

My second major summative for CEP.
This time my friend and I had to write a program which allocates students
to CCAs based on certain metrics. (Of course, their names and NRICs have been
obfuscated otherwise my school will probably get a lawsuit).
After writing the allocator, we wrote a GUI application that looks like
some crappy Windows XP virus.

Despite having 9 months of experience writing Python at this point, we
decided that it would be extremely cash money of us to hurl everything
into one massive file instead of splitting them up into smaller files.
That's how we got a 2000-line file of almost unreadable code with up to
13 levels of indentation.

This insanity is compounded by the fact that we haven't learned the concept
of MVC at the time, so we essentially misappropriated OOP in some of the most
horrifying ways possible. For example, the main program is wrapped around
a monolithic class called `Allocator`. Although we knew how classes work,
it never occurred to us that we could use classes to represent students or
spreadsheets as well. Hence, everything was stored as a dictionary with opaque
methods of accessing the data within. I wouldn't be surprised if any snippets
of code from this project finds its way onto
[r/badcode](https://www.reddit.com/r/badcode/).

We have improved since then.

## Notes

All files in this repository except for the README.md and .gitignore in
the root directory were written in 2019. This README was written
retrospectively.
