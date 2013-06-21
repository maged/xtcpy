xtcpy
=====

version 0.1.0

Python program to convert a directory of xls files (and all their sheets) to corresponding csv files. 
Note it does not suport xlsx.
The script relies on the great xlrd library. 

The outputted csv files will be of the format: ::

xlsfilename_sh{i}, sheet_number


Usage 
-----

You can run it so: 
xlsTocsv.py -i xlsfiledir/

-h for help. 

