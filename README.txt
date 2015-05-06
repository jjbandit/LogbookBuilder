#################################################################################
|                               Logbook Builder                                 |
|                                   v 1.0                                       |
|                           (C) Jesse Hughes 2014                               |
|                                                                               |
|                       License: GNU Public License 3.0                         |
|                    http://www.gnu.org/licenses/gpl-3.0.html                   |
|                                                                               |
#################################################################################

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see http://www.gnu.org/licenses/



 ---- File List

	This software should be distributed with three files:

	LogbookBuilder.xlsm
	README.txt
	LICENSE.txt

 ---- Preamble

	The intention of this software is to automate the process of building daily Logbooks.  In the LogbookBuilder.xlsm
	file you should find two sheets: "Template" and "Macro".

	The Template sheet is the master sheet that each new logbook will be built from.  The Macro sheet controls the
	behavior of the program.  

 ---- Usage

	To build logbooks with this program you must open LogbookBuilder.xlsm and input appropriate values to the fields 
	highlighted in orange.

	The Year field indicates what year the logbooks should be built for.  
		This should be a numerical value eg. 1995

	The Starting Month field indicates which month of the year (inclusive) to start building logbooks from.  
		This should be a numerical value between 1 and 12 
	
	The Ending Month field indicates which month of the year (inclusive) to stop building logbooks at.  
		This should be a numerical value between 1 and 12 
	
	The Save To field indicates where to save the logbooks when they are finished building. 
		Valid inputs for this field include names of any folder in the current users 'Home' directory.
		The default is Desktop, however another common input may be Documents.

	Suppose we input:

	Year           |  2012 
	Starting Month |   2
	Ending Month   |   3

	Save To        | Desktop

	When we run the program a prompt will appear indicating we have entered valid inputs: "Wait for the Magic!"
	Logbook Builder will build logbooks for the months of February and March 2012, correctly accounting for the
	leap year.  
	
	It should be noted that in order to build a logbook for only one month both Starting and Ending fields must be 
	set to the same value.

 ---- Uploading to Sharepoint

	Uploading the newly batched logbooks to Sharepoint involves creating a new file entry on the Aquatics Staff Log Book 
	page. It can be found at 
	https://goto.crd.bc.ca/teams/pcs/pr/8020RecreationPrograms/Forms/Aquatic%20Staff%20Log%20Books.aspx
	
	1. Navigate to the bottom of the page and follow the '+ Add document' link.
	2. Find the file using the file browser and upload with the OK button.
	
	3. There are a few changes that must be made for the new file to appear in the proper list.
	
		a) The "Content Type" field must be set to Staff Log Book
		b) The "Year" and "Month" fields must be set to appropriate values
		c) The "Service Area" field must be set to Aquatics


 ---- Known Bugs && TODO
	
	When creating a folder, it only checks if there is currently a folder with the same name as the folder it
	intends to create. This is undesirable behavior as it will create new folders even if no sheets will be over-
	written.

		--TODO 
			Check whether a file will be overwritten in the default save path and rename the file to be saved if this
			is the case.

	There is very little field validation and no error handling.

		--TODO 
			Make some error classes or something..

