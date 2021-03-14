# Excel-Autoplanner
-This program looks though a specified column(e.g. G) on excel and transfers all cells which contain numbers greater than 500 to a new column along with the appropriate colour that 
cell contained. Any time it finds a cell in the specified column it looks across that row at cells on the same row and additional specified columns(e.g. column B on the same row) 
and copy that information onto the same row but one column over as the information from the first specfied column. 

-For example if it hit a number >500 in column G(e.g. 600) then it would look across at cells on the same row and additional specified columns(e.g. A) and it would take the contents
from that cell as well, say for example it was 'central'. Then it copys the 600 to row 2 in another column(e.g. K) and will copy central to the cell on the same row but one column 
along, meanwhile doing this the program has copied background colours from the orginal cells and copied that over.

-It is written in Python, using the openpyxl library

-The idea is to eventually adapt this program to automate the filling in of multiple timetables from one input.
