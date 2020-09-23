'***************************************************************************
' Class Generator Version 1.6
'   The program will read a database and create classes and collections
'   for the tables within that database.
'
' Author:     Lawrence R. Miller  - lmiller@cpscars.com
' Copyright:  01/08/2003 - All rights reserved.
'
' I hastily wrote the program to help me create Classes and Collections for
' a large database project I am currently working on. It has been a
' tremendous time saver. It will work on any database you select. Whether
' it is a SQL Server and MS Access database. (I am writing my application
' to work with either one.)
'
' I freely distribute it to help other programmers, especially novices, though
' experienced programmers will benefit from the tedium of creating classes.
'
' Please feel free to make any modifications and improvements. If you do please
' update me. Also, feel free to contact me with any questions or suggestions.
'
' It is expressly forbidden to sell or distribute this program in any form for
' financial gain. It is only intended for your use and a Class Generator.
'***************************************************************************
' Two functions FormatPhone and StripPhone are supplied that I use for
' reading and writing phone numbers to the database. We store them as strings
' with no formatting. These routines are used to accomplish this. You can
' remove all references to them in the code if you do not want to do this.
' I am also storing Date Fields as strings and converting them when necessary.
'***************************************************************************

'IMPORTANT:
' I am assuming an .ActiveConnection = DBCars in all created
' Classes and Collections. You will need to modify this portion of the
' ClassGenerator Code to set the .ActiveConnection to your ADODB.Connection
' Variable Name. This is an easy search and replace.
'   Search for ".ActiveConnection = DBCars"
'   Replace with ".ActiveConnection = XXXXX" - XXXXX=Your Connection Variable
'***************************************************************************

'REVISIONS:
'
' I have added option buttons to allow the user to select whether they
' want to generate code with line continuations or to end each line
' and concatenate the previous line to the current one.
'
' EXAMPLE (Line Continuation)
'   Print #iFile, "s = s " & Code & " & _"
'   Print #iFile, "    Code & " & _"
'   etc...
'
' EXAMPLE (Concatination)
'   Print #iFile, "s = s " & Code
'   Print #iFile, "s = s " & Code
'   etc...
'
'***************************************************************************

What you can learn from this program.

1. How to use ADO - Connections, Recordsets, Update, Insert, Delete
2. How to use the ADOX Reference to get various information from a database
	(i.e. TableNames, TableTypes, FieldNames, FieldTypes, etc...)
3. Classes 
4. Collections
5. Writing to a sequential file.

I have placed Read, Update, Delete functions inside of each class as well as in the
collections. I did this because I have a few tables (classes) that do not need a
collection. The table will only have 1 record for a particular type of information I
am trying to receive. As an example I have a DEFAULTS table with 1 record for each 
AutoDealer in my AutoDealer table, and can only have 1 AutoDealer open at a time. 
Therefore, I don't need a collection and only use the single class within my AutoDealer
class.

'***************************************************************************
Possible Improvements:

As I stated above, I hastily wrote this program.

1. Much of the logic can be consolidated from the cmdClass_Click and 
   cmdCollection_Click events as the code is practically identical.
   Move this consolidated code into its own sub or function and call it 
   from the respective command click event.

2. Interface Improvements:
	A. Ability to select more than 1 table at a time.
	B. Ability to generate the Class and Collection with 1 click

3. Ability to add custom Events, Properties, Methods to the Classes and Collections.
	This I didn't do as it is just as easy to do in the environment.


'***************************************************************************
If you make any improvements, please let me know. 
lmiller@cpscars.com

