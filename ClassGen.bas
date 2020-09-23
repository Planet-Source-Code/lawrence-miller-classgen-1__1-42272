Attribute VB_Name = "ClassGen"
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
Option Explicit

'****************************
' Typed Field Data
'****************************
Type FieldData
  Name As String
  Type As Integer
  Desc As String
  Size As Long
End Type

Public QUOTE As String    'Quote Character  (")
Public CC    As String    'Concatenation Character with a Leading and Trailing Space( & )
Public COMMA As String    'Comma with Trailing Space (, )
Public SQ    As String    'Single Quote ("'")

Public ApPath As String         'Application Path
Public ConString As String      'Connection String
Public con As ADODB.Connection  'Connection

Public DBType   As Integer      '1=SQL Server 2=Access
Public Server   As String       'Name of Server
Public DBName   As String       'Name of Database
Public User     As String       'User Login
Public Password As String       'User Password

Public Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'***************************************************************************
' Function SetString
'   Use to truncate Class Properties to maximum length of string,
'   prior to writing to the database. Also translates "'" character to
'   a hyphon "-". SQL doesn't like this in the data.
'
' INPUT:
'       s - The string to be set.
'       L - Maximum Length of the string
'
' OUTPUT:
'       Adjusted string.
'
'  EXAMPLE:
'           Public Property Let Address(ByVal vData As String)
'             mAddress = SetString(vData, 25)
'           End Property
'***************************************************************************
Public Function SetString(ByVal s As String, ByVal L As Integer) As String
  Dim i As Integer
    
    s = Trim$(s)
    Do
      i = InStr(s, "'")
      If i <> 0 Then Mid$(s, i, 1) = "-"
    Loop Until i = 0
    If Len(s) > L Then
      SetString = Trim$(Left$(s, L))
    Else
      SetString = Trim$(s)
    End If
End Function

'***************************************************************************
' Function GetString
'   Wrapper Function to read string data from an .INI file.
'
' INPUT:
'       Section - Section Header of the INI file
'       sKey    - Key to be read
'       Default - Default value if the key has not been set yet.
'
' OUTPUT:
'       String value read from the INI file.
'***************************************************************************
Public Function GetString(Section As String, sKey As String, Default As String) As String
  Dim lResult As Long
  Dim Result As String
  
  Result = Space$(255)
  lResult = GetPrivateProfileString(Section, sKey, Default, Result, 255, ApPath & "ClassGen.INI")
  If lResult > 0 Then
    If Result = "" Then
      GetString = Left$(Default, lResult)
    Else
      GetString = Left$(Result, lResult)
    End If
  Else
    GetString = ""
  End If
End Function

'***************************************************************************
' Function PutString
'   Wrapper Function to write string data to an .INI file.
'
' INPUT:
'       Section - Section Header of the INI file.
'       Text    - Key to write
'       Value   - Value to be written into the key
'
' OUTPUT:
'       Long value returned by WritePrivateProfileString
'         Nonzero on success, zero on failure
'***************************************************************************
Public Function PutString(Section As String, Text As String, Value As String) As Long
  PutString = WritePrivateProfileString(Section, Text, Value, ApPath & "ClassGen.INI")
End Function

Private Sub Main()
    
    ApPath = App.Path
    If Right$(ApPath, 1) <> "\" Then ApPath = ApPath & "\"
    
    DBType = Val(GetString("ClassGen", "DatabaseType", "0"))
    Server = GetString("ClassGen", "ServerName", "")
    DBName = GetString("ClassGen", "DatabaseName", "")
    User = GetString("ClassGen", "User", "sa")
    Password = GetString("ClassGen", "Password", "")
    
    If DBType <> 0 Then
      If Not OpenDataBase Then frmDB.Show vbModal
      frmClassGen.Show
    Else
      frmDB.Show vbModal
    End If
End Sub

'***************************************************************************
' Function OpenDatabase
'   Opens the specified database.
'
' INPUT:
'       None
'
' OUTPUT:
'       True if opened successfully, False if failure
'***************************************************************************
Public Function OpenDataBase() As Boolean
    
    If DBType = 1 Then
      ConString = "Provider=SQLOLEDB.1;" & _
                  "Persist Security Info=False;" & _
                  "Data Source=" & Trim$(Server) & ";" & _
                  "User ID =" & Trim$(User) & ";" & _
                  "Password=" & Trim$(Password) & ";" & _
                  "Initial Catalog=" & Trim$(DBName)
    Else
      ConString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Persist Security Info=False;" & _
                  "Data Source=" & Trim$(DBName) & ";" '& _
                  '"User ID=" & Trim$(User) & ";" & _
                  '"Password=" & Trim$(Password) & ";"
    End If
    On Local Error GoTo Handler
    Set con = New ADODB.Connection
    With con
      .ConnectionString = ConString
      .Open
      If .State = 0 Then
        MsgBox "Unable to open Database"
        OpenDataBase = False
      Else
        OpenDataBase = True
      End If
    End With
    On Local Error GoTo 0
    
    Exit Function

Handler:
  OpenDataBase = False
    Exit Function
    
End Function

'***************************************************************************
' Function FormatPhone
'   Routine to format a phone number. It will call StripPhone to remove
'   any formatting it may already have.
'
' INPUT:
'       Txt - Phone Number to be formatted
'
' OUTPUT:
'       Formatted Phone Number   (xxx) xxx-xxxx
'***************************************************************************
Public Function FormatPhone(ByVal Txt As String) As String
  Dim s As String
    Txt = StripPhone(Txt)   'If it is already formatted then strip formatting
    Txt = Left$(Txt & Space$(10), 10)
    s = "(" & Left$(Txt, 3) & ") "
    s = s & Mid$(Txt, 4, 3) & "-"
    s = s & Right$(Txt, 4)
    FormatPhone = s
End Function

'***************************************************************************
' Function StripPhone
'   Routine to strip all non-numeric characters from a phone number.
'
' INPUT:
'       Txt - Phone Number to be stripped
'
' OUTPUT:
'       Stripped Phone Number   xxxxxxxxxx
'***************************************************************************
Public Function StripPhone(ByVal Txt As String) As String
  Dim i As Integer
  Dim j As Integer
  Dim C As Integer
  Dim W As String
    
    j = Len(Txt)
    For i = 1 To j
      C = Asc(Mid$(Txt, i, 1))
      If C > 46 And C < 58 Then
        W = W & Mid$(Txt, i, 1)
      End If
    Next
    StripPhone = W
End Function
