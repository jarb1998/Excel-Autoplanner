# Excel-Autoplanner
-This script autofills in a teachers timetable based on the class and day/time they have said class, with respect to how the subject is colour coded and how it is merged.
-It is written in python and uses the openpyxl and xlsxwriter packages
-In order to run the script it must be in the same folder as the excel document you wish to run it on, and in addition make sure that the name of the excel document is in the appropriate save/open commands on lines 9, 197, 200, 238

-Note this script is made to run on a specific format of timetable(as seen in the example data) wherein there is one section for teachers names and one for all the classe's being taught. Every teacher's name must have there initals in brackets next to it and every class must have the teachers initals in brackets next to it in order for the script to know which teacher to plot the class to. It is also important that there is a whitespace between the teacers name and there initials in brackets like so
  John Smith (JS)
 and not
  John Smith(JS)
 
 -In addition the script cannot copy over theme colours to new cells, so if you wish to copy colour coding over it must be coloured with standard/custom colours
 -Finally the script has a overwrite log, if you accidentally schedule the teacher for two classes at the exact same timeslot it will be written into the overwrite log text file. 
