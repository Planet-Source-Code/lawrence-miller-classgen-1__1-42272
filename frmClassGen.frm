VERSION 5.00
Begin VB.Form frmClassGen 
   Caption         =   "Class Generator"
   ClientHeight    =   4770
   ClientLeft      =   4920
   ClientTop       =   3090
   ClientWidth     =   4395
   Icon            =   "frmClassGen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optLine 
      Caption         =   "Use String Concatenation"
      Height          =   225
      Index           =   1
      Left            =   2100
      TabIndex        =   10
      Top             =   3510
      Width           =   2235
   End
   Begin VB.OptionButton optLine 
      Caption         =   "Use Line Continuation"
      Height          =   225
      Index           =   0
      Left            =   2100
      TabIndex        =   9
      Top             =   3300
      Value           =   -1  'True
      Width           =   2235
   End
   Begin VB.CommandButton cmdDB 
      Caption         =   "Select Database"
      Height          =   735
      Left            =   2310
      TabIndex        =   8
      Top             =   3870
      Width           =   1425
   End
   Begin VB.CommandButton cmdCollection 
      Caption         =   "Create Collection"
      Height          =   525
      Left            =   2430
      TabIndex        =   7
      Top             =   1350
      Width           =   1245
   End
   Begin VB.CommandButton cmdClass 
      Caption         =   "Create Class"
      Height          =   525
      Left            =   2430
      TabIndex        =   2
      Top             =   780
      Width           =   1245
   End
   Begin VB.TextBox txtClass 
      Height          =   345
      Left            =   2010
      TabIndex        =   1
      Top             =   300
      Width           =   1695
   End
   Begin VB.ListBox lstTables 
      Height          =   4350
      Left            =   90
      TabIndex        =   0
      Top             =   300
      Width           =   1875
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   2010
      TabIndex        =   6
      Top             =   1920
      Width           =   1905
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2010
      TabIndex        =   5
      Top             =   2280
      Width           =   1905
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Name"
      Height          =   225
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      Top             =   90
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Table Name"
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   60
      Width           =   1245
   End
End
Attribute VB_Name = "frmClassGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mContinue As Boolean

Dim rst As ADODB.Recordset
Dim fld As ADODB.Fields

'******************************************************************
' cmdClass_Click
'
'   Creates a Class from the selected table.
'   This routine will get the name of the Class to build
'   from txtClass.Text (the class name text box).
'******************************************************************
Private Sub cmdClass_Click()
  Dim i As Integer
  Dim j As Integer
  Dim N As String
  Dim T As Integer
  Dim s As String
  Dim Tbl As String
  
  Dim FNames() As FieldData
  
  Dim iFile As Integer
  Dim sFile As String
  
    If lstTables.ListIndex <> -1 And txtClass.Text <> "" Then
      Me.MousePointer = vbHourglass
      
      Tbl = Trim$(lstTables.List(lstTables.ListIndex))
      
      'If there is a space in the TableName then surround it in brackets
      If InStr(Tbl, " ") <> 0 Then Tbl = "[" & Tbl & "]"
      
      lblGen.Caption = "Opening Table " & Tbl
      lblGen.Refresh
      Set rst = New ADODB.Recordset
      With rst
        .ActiveConnection = con
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "SELECT * FROM " & Tbl
        .Open
        j = .Fields.Count - 1
        
        For i = 0 To j
          DoEvents
          ReDim Preserve FNames(i)
          FNames(i).Name = .Fields.Item(i).Name
          FNames(i).Type = .Fields.Item(i).Type
          FNames(i).Size = .Fields.Item(i).DefinedSize
          Select Case .Fields.Item(i).Type
            Case adTinyInt, adUnsignedTinyInt
              FNames(i).Desc = "Byte"
            Case adSmallInt, adUnsignedSmallInt
              FNames(i).Desc = "Integer"
            Case adInteger, adUnsignedInt
              FNames(i).Desc = "Long"
            Case adBoolean
              FNames(i).Desc = "Boolean"
            Case adBSTR, adDBDate, adDBTime, adDBTimeStamp, adVarChar, adVarWChar, adWChar, adChar
              FNames(i).Desc = "String"
            Case adSingle
              FNames(i).Desc = "Single"
            Case adDecimal, adDouble, adNumeric, adCurrency
              FNames(i).Desc = "Double"
            Case adVariant
              FNames(i).Desc = "Variant"
            Case Else
              FNames(i).Desc = "String"
          End Select
        Next
        .Close
      End With
      Set rst = Nothing
      Close iFile
      
      sFile = ApPath & txtClass.Text & ".cls"
      iFile = FreeFile
      Open sFile For Output As iFile
      Print #iFile, "VERSION 1.0 CLASS"
      Print #iFile, "BEGIN"
      Print #iFile, "  MultiUse = -1  'True"
      Print #iFile, "  Persistable = 0  'NotPersistable"
      Print #iFile, "  DataBindingBehavior = 0  'vbNone"
      Print #iFile, "  DataSourceBehavior = 0   'vbNone"
      Print #iFile, "  MTSTransactionMode = 0   'NotAnMTSObject"
      Print #iFile, "End"
      Print #iFile, "Attribute VB_Name = " & QUOTE & txtClass.Text & QUOTE
      Print #iFile, "Attribute VB_GlobalNameSpace = False"
      Print #iFile, "Attribute VB_Creatable = True"
      Print #iFile, "Attribute VB_PredeclaredId = False"
      Print #iFile, "Attribute VB_Exposed = False"
      Print #iFile, "Attribute VB_Ext_KEY = " & QUOTE & "SavedWithClassBuilder6" & QUOTE & " ," & QUOTE & "No" & QUOTE
      Print #iFile, "Attribute VB_Ext_KEY = " & QUOTE & "Top_Level" & QUOTE & ", " & QUOTE & "Yes" & QUOTE
      Print #iFile, "Option Explicit"
      Print #iFile, ""
      
'*******************************************************
'****   Create the Private Variable Declarations    ****
'*******************************************************
      lblGen.Caption = "Generating"
      lblStatus.Caption = "Private Variables"
      lblStatus.Refresh
      For i = 0 To j
        If UCase$(FNames(i).Name) = "ID" Then
          s = "Private m_" & FNames(i).Name
        Else
          s = "Private m" & FNames(i).Name
        End If
        If Len(s) < 35 Then s = Left$(s & Space$(35), 35)
        s = s & " AS " & FNames(i).Desc
        Print #iFile, s
      Next
      Print #iFile, ""
'*******************************************************
'****           Create the Get Properties           ****
'*******************************************************
      lblStatus.Caption = "Public Properties"
      lblStatus.Refresh
      For i = 0 To j
        DoEvents
        Print #iFile, "Public Property GET " & FNames(i).Name & "() AS " & FNames(i).Desc
        If UCase$(FNames(i).Name) = "ID" Then
          Print #iFile, "   " & FNames(i).Name & " = m_" & FNames(i).Name
        Else
          Print #iFile, "   " & FNames(i).Name & " = m" & FNames(i).Name
        End If
        Print #iFile, "End Property"
        Print #iFile, ""
'*******************************************************
'****           Create the LET Properties           ****
'*******************************************************
        Print #iFile, "Public Property LET " & FNames(i).Name & "(ByVal vData as " & FNames(i).Desc & ")"
        If FNames(i).Desc = "String" Then
          If UCase$(FNames(i).Name) = "ID" Then
            Print #iFile, "      m_" & FNames(i).Name & " = SetString(vData, " & FNames(i).Size & ")"
          Else
            If InStr(UCase$(FNames(i).Name), "PHONE") <> 0 Or _
               InStr(UCase$(FNames(i).Name), "FAX") <> 0 Or _
               InStr(UCase$(FNames(i).Name), "CELL") <> 0 Or _
               InStr(UCase$(FNames(i).Name), "PAGER") <> 0 Then
                  If InStr(UCase$(FNames(i).Name), "EXT") = 0 Then
                    Print #iFile, "    m" & FNames(i).Name & " = FormatPhone(vData)"
                  Else
                    Print #iFile, "    m" & FNames(i).Name & " = SetString(vData, " & FNames(i).Size & ")"
                  End If
            Else
              Print #iFile, "    m" & FNames(i).Name & " = SetString(vData, " & FNames(i).Size & ")"
            End If
          End If
        Else
          If UCase$(FNames(i).Name) = "ID" Then
            Print #iFile, "   m_" & FNames(i).Name & " = vData"
          Else
            Print #iFile, "   m" & FNames(i).Name & " = vData"
          End If
        End If
        Print #iFile, "End Property"
        Print #iFile, ""
      Next

'*******************************************************
'****         Create the Read Method                ****
'*******************************************************
      lblStatus.Caption = "Public Read Method"
      lblStatus.Refresh
      Print #iFile, "Public Sub Read(ByVal ID AS Long)"
      Print #iFile, "  Dim s  As String"
      Print #iFile, "  Dim RS As ADODB.Recordset"
      Print #iFile, ""
      Print #iFile, "    If ID > 0 Then"
      Print #iFile, "      s = " & QUOTE & "SELECT * FROM " & Tbl & " WHERE " & FNames(0).Name & " = " & QUOTE & " & ID"
      Print #iFile, ""
      Print #iFile, "      Set RS = New ADODB.Recordset"
      Print #iFile, "      With RS"
      Print #iFile, "        .Source = s"
      Print #iFile, "        .ActiveConnection = DBCars"
      
      Print #iFile, "        .CursorLocation = adUseClient"
      Print #iFile, "        .CursorType = adOpenForwardOnly"
      Print #iFile, "        .LockType = adLockOptimistic"
      Print #iFile, "        .Open"
      Print #iFile, "        If .State = 1 Then"
      Print #iFile, "          If Not (.BOF And .EOF) Then"
      For i = 0 To j
        DoEvents
        Select Case FNames(i).Type
          Case adTinyInt, adUnsignedTinyInt, adSmallInt, adUnsignedSmallInt, adInteger, adUnsignedInt, adBoolean, adSingle, adDecimal, adDouble, adNumeric, adCurrency
            If UCase$(FNames(i).Name) = "ID" Then
              Print #iFile, "            m_" & FNames(i).Name & " = Val(" & QUOTE & QUOTE & " & !" & FNames(i).Name & ")"
            Else
              Print #iFile, "            m" & FNames(i).Name & " = Val(" & QUOTE & QUOTE & " & !" & FNames(i).Name & ")"
            End If
          Case Else
            If UCase$(FNames(i).Name) = "ID" Then
              Print #iFile, "            m_" & FNames(i).Name & " = Trim$(" & QUOTE & QUOTE & " & !" & FNames(i).Name & ")"
            Else
              Print #iFile, "            m" & FNames(i).Name & " = Trim$(" & QUOTE & QUOTE & " & !" & FNames(i).Name & ")"
            End If
        End Select
      Next
      Print #iFile, "          End If"
      Print #iFile, "        End If"
      Print #iFile, "        .Close"
      Print #iFile, "      End With"
      Print #iFile, "      Set RS = Nothing"
      Print #iFile, "    End If"
      Print #iFile, ""
      Print #iFile, "End Sub"
      Print #iFile, ""
      
'*******************************************************
'****         Create the Update Method              ****
'*******************************************************
      lblStatus.Caption = "Public Update Method"
      lblStatus.Refresh
      Print #iFile, "Public Sub Update()"
      Print #iFile, "  Dim s  As String"
      Print #iFile, "  Dim RS As ADODB.Recordset"
      Print #iFile, ""
'**** UPDATE ****
      Print #iFile, "    If m_" & FNames(0).Name & " <> 0 Then"
      s = Space$(6) & "s = " & QUOTE & "UPDATE " & Tbl & " SET "
      T = 2
      For i = 1 To j
        DoEvents
        Select Case FNames(i).Type
          Case adBoolean
            N = FNames(i).Name & " = " & QUOTE & CC & "Abs(m" & FNames(i).Name & ")"
          Case adTinyInt, adUnsignedTinyInt, adSmallInt, adUnsignedSmallInt, adInteger, adUnsignedInt, adSingle, adDecimal, adDouble, adNumeric, adCurrency
            N = FNames(i).Name & " = " & QUOTE & CC & "m" & FNames(i).Name
          Case Else       'String
            If (InStr(UCase$(FNames(i).Name), "PHONE") <> 0 Or _
               InStr(UCase$(FNames(i).Name), "FAX") <> 0 Or _
               InStr(UCase$(FNames(i).Name), "CELL") <> 0 Or _
               InStr(UCase$(FNames(i).Name), "PAGER") <> 0) And _
               InStr(UCase$(FNames(i).Name), "EXT") = 0 Then
              N = FNames(i).Name & " = '" & QUOTE & CC & "StripPhone(m" & FNames(i).Name & ")" & CC & QUOTE & "'"
            Else
              N = FNames(i).Name & " = '" & QUOTE & CC & "m" & FNames(i).Name & CC & QUOTE & "'"
            End If
        End Select
        If Len(s) = 0 Then
          s = QUOTE & N
        Else
          If (Right$(s, 1) <> QUOTE) And (Right$(s, 1) <> "'") Then
            If Trim$(Right$(s, 5)) = "SET" Then
              s = s & N
            Else
              s = s & CC & QUOTE & COMMA & N
            End If
          ElseIf Right$(s, 1) = "'" Then
            s = s & " " & N
          ElseIf Right$(s, 1) = QUOTE Then
            s = s & CC & N
          Else
            s = s & N
          End If
        End If
        
        If mContinue Then
          If i < j And T > 1 Then
            T = 0
            If Right$(s, 1) = "'" Then
              Print #iFile, s; COMMA; QUOTE; CC; "_"
            Else
              Print #iFile, s; CC; QUOTE; COMMA; QUOTE; CC; "_"
            End If
            Print #iFile, Space$(10);
            s = ""
          ElseIf i = j Then
            If Right$(s, 1) = "'" Then
              Print #iFile, s; QUOTE; CC; "_"
            Else
              Print #iFile, s; CC; "_"
            End If
          End If
          T = T + 1
        Else
          If i < j And T > 1 Then
            T = 0
            If Right$(s, 1) = "'" Then
              Print #iFile, s; COMMA; QUOTE
            Else
              Print #iFile, s; CC; QUOTE; COMMA; QUOTE
            End If
            Print #iFile, Space$(6); "s = s & ";
            s = ""
          ElseIf i = j Then
            If Right$(s, 1) = "'" Then
              Print #iFile, s; QUOTE
            Else
              Print #iFile, s
            End If
          End If
          T = T + 1
        End If
      Next
      If mContinue Then
        Print #iFile, Space$(10); QUOTE; " Where "; FNames(0).Name; " = " & QUOTE; CC; "m_"; FNames(0).Name
      Else
        Print #iFile, Space$(6); "s = s & "; QUOTE; " Where "; FNames(0).Name; " = " & QUOTE; CC; "m_"; FNames(0).Name
      End If
      Print #iFile, "      DbCars.Execute s"
'**** INSERT INTO ****
      Print #iFile, "    Else"
      If mContinue Then
        Print #iFile, Space$(6); "s = "; QUOTE; "INSERT INTO "; Tbl; QUOTE; CC; "_"
        Print #iFile, Space$(12); QUOTE; "(";
      Else
        Print #iFile, Space$(6); "s = "; QUOTE; "INSERT INTO "; Tbl; QUOTE
        Print #iFile, Space$(6); "s = s & "; QUOTE; "(";
      End If
      For i = 1 To j
        If i < j Then
          Print #iFile, FNames(i).Name & ", ";
          If (i Mod 6) = 0 And i <> j Then
            If mContinue Then
              Print #iFile, QUOTE; CC; "_"
              Print #iFile, Space$(12); QUOTE;
            Else
              Print #iFile, QUOTE
              Print #iFile, Space$(6); "s = s & "; QUOTE;
            End If
          End If
        Else
          If mContinue Then
            Print #iFile, FNames(i).Name & ") " & QUOTE & CC & "_"
          Else
            Print #iFile, FNames(i).Name & ") " & QUOTE
          End If
        End If
      Next
      If mContinue Then
        s = Space$(10) & QUOTE & "VALUES ("
      Else
        s = Space$(6) & "s = s & " & QUOTE & "VALUES ("
      End If
      T = 0
      For i = 1 To j
        DoEvents
        Select Case FNames(i).Type
          Case adBoolean
            N = "Abs(m" & FNames(i).Name & ")"
          Case adTinyInt, adUnsignedTinyInt, adSmallInt, adUnsignedSmallInt, adInteger, adUnsignedInt, adSingle, adDecimal, adDouble, adNumeric, adCurrency
            N = "m" & FNames(i).Name
          Case Else
            N = SQ & CC & "m" & FNames(i).Name & CC & SQ
        End Select
        If Len(s) = 12 Or Len(s) = 14 Then
          s = s & N
        Else
          If Right$(s, 1) = "(" Then
            If Left$(N, 1) = QUOTE Then
              s = s & Mid$(N, 2)
            Else
              If mContinue Then
                s = s & QUOTE & CC & N
              Else
                s = s & N
              End If
            End If
          Else
            If Right$(s, 1) = QUOTE Then
              If Left$(N, 1) = QUOTE Then
                s = Left$(s, Len(s) - 1) & COMMA & Mid$(N, 2)
              Else
                s = Left$(s, Len(s) - 1) & COMMA & QUOTE & CC & N
              End If
            Else
              If mContinue Then
                If Left$(N, 1) = QUOTE Then
                  s = s & CC & QUOTE & COMMA & Mid$(N, 2)
                Else
                  s = s & CC & QUOTE & COMMA & QUOTE & CC & N
                End If
              Else
                If Left$(N, 1) = QUOTE Then
                  s = s & CC & QUOTE & COMMA & Mid$(N, 2)
                Else
                  s = s & CC & QUOTE & COMMA & QUOTE & CC & N
                End If
              End If
            End If
          End If
        End If
        T = T + 1
        If i < j And T > 3 Then
          T = 0
          If mContinue Then
            If Right$(s, 1) = QUOTE Then
              Print #iFile, Left$(s, Len(s) - 1); COMMA; QUOTE; CC; "_"
            Else
              Print #iFile, s; CC; QUOTE; COMMA; QUOTE; "_"
            End If
            s = Space$(12)
          Else
            Print #iFile, s
            s = Space$(6) & "s = s & "
          End If
        ElseIf i = j Then
          Print #iFile, s; CC & QUOTE & ")"
        End If
      Next
      Print #iFile, ""
      Print #iFile, "      SET RS = New ADODB.Recordset"
      Print #iFile, "      Set RS = DbCars.Execute(s)"
      Print #iFile, "      Set RS = DbCars.Execute(" & QUOTE & "SELECT @@IDENTITY" & QUOTE & ")"
      Print #iFile, "      m" & FNames(0).Name & " = RS.Fields(0).Value"
      Print #iFile, "      RS.Close"
      Print #iFile, "      Set RS = Nothing"
      Print #iFile, "    End If"
      Print #iFile, "End Sub"
      
'*******************************************************
'****         Create the Delete Method              ****
'*******************************************************
      lblStatus.Caption = "Public Delete Method"
      Me.Refresh
      Print #iFile, ""
      Print #iFile, "Public Sub Delete(ByVal ID As Long)"
      Print #iFile, "  Dim s As String"
      Print #iFile, ""
      Print #iFile, "    s = " & QUOTE & "DELETE FROM " & Tbl;
      Print #iFile, " WHERE " & FNames(0).Name & " = " & QUOTE & " & m_" & FNames(0).Name
      Print #iFile, "    DbCars.Execute s"
      Print #iFile, "End Sub"
      
      Close iFile
      lblGen.Caption = ""
      lblStatus.Caption = ""
      Me.MousePointer = vbDefault
      s = "Class file has been generated:" & vbCrLf & vbCrLf & sFile
      MsgBox s, vbOKOnly, "Class Generator"
      
    End If
End Sub

'******************************************************************
' cmdCollection_Click
'
'   Creates a Collection class from the selected table.
'   This routine will get the name of the Class to build
'   from txtClass.Text (the class name text box).
'
'   It automatically add an "s" character to the end of
'   the class name.
'******************************************************************
Private Sub cmdCollection_Click()
  Dim i As Integer
  Dim j As Integer
  Dim N As String
  Dim T As Integer
  Dim s As String
  Dim BaseClass As String
  Dim Tbl As String
  
  Dim FNames() As FieldData
  Dim iFile As Integer
  Dim sFile As String
  
    If lstTables.ListIndex <> -1 And txtClass.Text <> "" Then
      Me.MousePointer = vbHourglass
      
      Tbl = lstTables.List(lstTables.ListIndex)
      
      'If there is a space in the TableName then surround it in brackets
      If InStr(Tbl, " ") <> 0 Then Tbl = "[" & Tbl & "]"
      
      Set rst = New ADODB.Recordset
      With rst
        .ActiveConnection = con
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "SELECT * FROM " & Tbl
        .Open
        j = .Fields.Count - 1
        
        For i = 0 To j
          DoEvents
          ReDim Preserve FNames(i)
          FNames(i).Name = .Fields.Item(i).Name
          FNames(i).Type = .Fields.Item(i).Type
          FNames(i).Size = .Fields.Item(i).DefinedSize
          Select Case .Fields.Item(i).Type
            Case adTinyInt, adUnsignedTinyInt
              FNames(i).Desc = "Byte"
            Case adSmallInt, adUnsignedSmallInt
              FNames(i).Desc = "Integer"
            Case adInteger, adUnsignedInt
              FNames(i).Desc = "Long"
            Case adBoolean
              FNames(i).Desc = "Boolean"
            Case adBSTR, adDBDate, adDBTime, adDBTimeStamp, adVarChar, adVarWChar, adWChar, adChar
              FNames(i).Desc = "String"
            Case adSingle
              FNames(i).Desc = "Single"
            Case adDecimal, adDouble, adNumeric, adCurrency
              FNames(i).Desc = "Double"
            Case adVariant
              FNames(i).Desc = "Variant"
            Case Else
              FNames(i).Desc = "String"
          End Select
        Next
        .Close
      End With
      Set rst = Nothing
      
      BaseClass = Trim$(txtClass.Text)
      N = BaseClass & "s"
      sFile = ApPath & N & ".cls"
      iFile = FreeFile
      Open sFile For Output As iFile
      Print #iFile, "VERSION 1.0 CLASS"
      Print #iFile, "BEGIN"
      Print #iFile, "  MultiUse = -1  'True"
      Print #iFile, "  Persistable = 0  'NotPersistable"
      Print #iFile, "  DataBindingBehavior = 0  'vbNone"
      Print #iFile, "  DataSourceBehavior = 0   'vbNone"
      Print #iFile, "  MTSTransactionMode = 0   'NotAnMTSObject"
      Print #iFile, "End"
      Print #iFile, "Attribute VB_Name = " & QUOTE & N & QUOTE
      Print #iFile, "Attribute VB_GlobalNameSpace = False"
      Print #iFile, "Attribute VB_Creatable = True"
      Print #iFile, "Attribute VB_PredeclaredId = False"
      Print #iFile, "Attribute VB_Exposed = False"
      Print #iFile, "Attribute VB_Ext_KEY = " & QUOTE & "SavedWithClassBuilder6" & QUOTE & " ," & QUOTE & "No" & QUOTE
      Print #iFile, "Attribute VB_Ext_KEY = " & QUOTE & "Collection" & QUOTE & " ," & QUOTE & BaseClass & QUOTE
      Print #iFile, "Attribute VB_Ext_KEY = " & QUOTE & "Member0" & QUOTE & " ," & QUOTE & BaseClass & QUOTE
      Print #iFile, "Attribute VB_Ext_KEY = " & QUOTE & "Top_Level" & QUOTE & ", " & QUOTE & "Yes" & QUOTE
      Print #iFile, "Option Explicit"
      Print #iFile, ""
      
'*******************************************************
'****   Create the Private Variable Declarations    ****
'*******************************************************
      lblGen.Caption = "Generating"
      lblStatus.Caption = "Private Variables"
      Me.Refresh
      Print #iFile, "Private mCol as Collection"
      Print #iFile, ""
      
'*******************************************************
'****         Create the Item Property              ****
'*******************************************************
      Print #iFile, "Public Property Get Item(vIndexKey As Variant) As " & BaseClass
      Print #iFile, "Attribute Item.VB_UserMemId = 0"
      Print #iFile, "    Set Item = mCol(vIndexKey)"
      Print #iFile, "End Property"
      Print #iFile, ""
      
'*******************************************************
'****         Create the Count Property             ****
'*******************************************************
      Print #iFile, "Public Property Get Count() As Long"
      Print #iFile, "    Count = mCol.Count"
      Print #iFile, "End Property"
      Print #iFile, ""
      
'*******************************************************
'****         Create the NewEnum Property           ****
'*******************************************************
      Print #iFile, "Public Property Get NewEnum() As IUnknown"
      Print #iFile, "Attribute NewEnum.VB_UserMemId = -4"
      Print #iFile, "Attribute NewEnum.VB_MemberFlags = " & QUOTE & "40" & QUOTE
      Print #iFile, "    Set NewEnum = mCol.[_NewEnum]"
      Print #iFile, "End Property"
      Print #iFile, ""

'*******************************************************
'****         Create the Add Method                 ****
'*******************************************************
      s = "Public Function Add("
      For i = 0 To j
        DoEvents
        s = s & "ByVal " & FNames(i).Name & " As " & FNames(i).Desc & ", "
      Next
      s = s & "Optional sKey as String) as " & BaseClass
      Print #iFile, s
      Print #iFile, "  Dim obj as " & BaseClass
      Print #iFile, ""
      Print #iFile, "    Set obj = New " & BaseClass
      Print #iFile, ""
      s = "    If " & FNames(0).Name & " = 0 Then " & FNames(0).Name & " = AddNew("
      For i = 1 To j
        DoEvents
        s = s & FNames(i).Name
        If i < j Then s = s & ", "
      Next
      s = s & ")"
      Print #iFile, s
      Print #iFile, ""
      Print #iFile, "    'Set the properties passed into the method"
      For i = 0 To j
        DoEvents
        Print #iFile, "    obj." & FNames(i).Name & " = " & FNames(i).Name
      Next
      Print #iFile, ""
      Print #iFile, "    If Len(sKey) = 0 Then"
      Print #iFile, "      mcol.Add obj"
      Print #iFile, "    Else"
      Print #iFile, "      mcol.Add obj, sKey"
      Print #iFile, "    End If"
      Print #iFile, ""
      Print #iFile, "    Set Add = obj                'Return the object created"
      Print #iFile, "    Set obj = Nothing"
      Print #iFile, "End Function"
      Print #iFile, ""
      
'*******************************************************
'****         Create the AddNew Method              ****
'*******************************************************
      s = "Public Function AddNew("
      For i = 1 To j
        DoEvents
        s = s & "ByVal " & FNames(i).Name & " As " & FNames(i).Desc
        If i < j Then s = s & ", "
      Next
      s = s & ") As Long"
      Print #iFile, s
      Print #iFile, "  Dim s As String"
      Print #iFile, "  Dim RS As ADODB.Recordset"
      Print #iFile, ""
      If mContinue Then
        Print #iFile, "'----[ Because we are using line continuations, you   ]----"
        Print #iFile, "'----[ may have to consolidate these before loading   ]----"
        Print #iFile, "'----[ this class into the environment.               ]----"
        Print #iFile, ""
        Print #iFile, Space$(6); "s = "; QUOTE; "INSERT INTO "; Tbl; QUOTE; CC; "_"
        Print #iFile, Space$(12); QUOTE; "(";
      Else
        Print #iFile, Space$(6); "s = "; QUOTE; "INSERT INTO "; Tbl; QUOTE
        Print #iFile, Space$(6); "s = s & "; QUOTE; "(";
      End If
      
      For i = 1 To j
        If i < j Then
          Print #iFile, FNames(i).Name & ", ";
          If (i Mod 6) = 0 And i <> j Then
            If mContinue Then
              Print #iFile, QUOTE; CC; "_"
              Print #iFile, Space$(12); QUOTE;
            Else
              Print #iFile, QUOTE
              Print #iFile, Space$(6); "s = s & "; QUOTE;
            End If
          End If
        Else
          If mContinue Then
            Print #iFile, FNames(i).Name & ") " & QUOTE & CC & "_"
          Else
            Print #iFile, FNames(i).Name & ") " & QUOTE
          End If
        End If
      Next
      If mContinue Then
        s = Space$(10) & QUOTE & "VALUES ("
      Else
        s = Space$(6) & "s = s & " & QUOTE & "VALUES ("
      End If
      T = 0
      For i = 1 To j
        DoEvents
        Select Case FNames(i).Type
          Case adBoolean
            N = "Abs(m" & FNames(i).Name & ")"
          Case adTinyInt, adUnsignedTinyInt, adSmallInt, adUnsignedSmallInt, adInteger, adUnsignedInt, adSingle, adDecimal, adDouble, adNumeric, adCurrency
            N = "m" & FNames(i).Name
          Case Else
            N = SQ & CC & "m" & FNames(i).Name & CC & SQ
        End Select
        If Len(s) = 12 Or Len(s) = 14 Then
          s = s & N
        Else
          If Right$(s, 1) = "(" Then
            If Left$(N, 1) = QUOTE Then
              s = s & Mid$(N, 2)
            Else
              If mContinue Then
                s = s & QUOTE & CC & N
              Else
                s = s & N
              End If
            End If
          Else
            If Right$(s, 1) = QUOTE Then
              If Left$(N, 1) = QUOTE Then
                s = Left$(s, Len(s) - 1) & COMMA & Mid$(N, 2)
              Else
                s = Left$(s, Len(s) - 1) & COMMA & QUOTE & CC & N
              End If
            Else
              If mContinue Then
                If Left$(N, 1) = QUOTE Then
                  s = s & CC & QUOTE & COMMA & Mid$(N, 2)
                Else
                  s = s & CC & QUOTE & COMMA & QUOTE & CC & N
                End If
              Else
                If Left$(N, 1) = QUOTE Then
                  s = s & CC & QUOTE & COMMA & Mid$(N, 2)
                Else
                  s = s & CC & QUOTE & COMMA & QUOTE & CC & N
                End If
              End If
            End If
          End If
        End If
        T = T + 1
        If i < j And T > 3 Then
          T = 0
          If mContinue Then
            If Right$(s, 1) = QUOTE Then
              Print #iFile, Left$(s, Len(s) - 1); COMMA; QUOTE; CC; "_"
            Else
              Print #iFile, s; CC; QUOTE; COMMA; QUOTE; "_"
            End If
            s = Space$(12)
          Else
            Print #iFile, s
            s = Space$(6) & "s = s & "
          End If
        ElseIf i = j Then
          Print #iFile, s; CC & QUOTE & ")"
        End If
      Next
      Print #iFile, ""
      Print #iFile, "      SET RS = New ADODB.Recordset"
      Print #iFile, "      Set RS = DbCars.Execute(s)"
      Print #iFile, "      Set RS = DbCars.Execute(" & QUOTE & "SELECT @@IDENTITY" & QUOTE & ")"
      Print #iFile, "      AddNew = RS.Fields(0).Value"
      Print #iFile, "      RS.Close"
      Print #iFile, "      Set RS = Nothing"
      Print #iFile, "End Function"
      Print #iFile, ""

'*******************************************************
'****         Create the Remove Method              ****
'*******************************************************
      Print #iFile, "Public Sub Remove(vIndexKey As Variant)"
      Print #iFile, "    DbCars.Execute " & QUOTE & "DELETE FROM " & Tbl & " WHERE " & FNames(0).Name & " = " & QUOTE & " & mCol.Item(vIndexKey)." & FNames(0).Name
      Print #iFile, "    mCol.Remove vIndexKey"
      Print #iFile, "End Sub"
      Print #iFile, ""
      
'*******************************************************
'****             Create the Read Method            ****
'*******************************************************
      Print #iFile, "Public Sub Read()"
      Print #iFile, "  Dim RS As ADODB.Recordset"
      Print #iFile, ""
      Print #iFile, "    Set RS = New ADODB.Recordset"
      Print #iFile, "    With RS"
      Print #iFile, "'----[You will probably need to change the ORDER BY clause ]----"
      Print #iFile, "      .Source = " & QUOTE & "SELECT * FROM " & Tbl & " ORDER BY " & FNames(1).Name & QUOTE
      Print #iFile, "      .ActiveConnection = DbCars"
      Print #iFile, "      .CursorLocation = adUseClient"
      Print #iFile, "      .CursorType = adOpenForwardOnly"
      Print #iFile, "      .LockType = adLockOptimistic"
      Print #iFile, "      .Open"
      Print #iFile, "      Do Until .EOF"
      s = "        Add "
      For i = 0 To j
        DoEvents
        If InStr(UCase$(FNames(i).Name), "PHONE") <> 0 Or _
           InStr(UCase$(FNames(i).Name), "FAX") <> 0 Or _
           InStr(UCase$(FNames(i).Name), "CELL") <> 0 Or _
           InStr(UCase$(FNames(i).Name), "PAGER") <> 0 Then
              If InStr(UCase$(FNames(i).Name), "EXT") = 0 Then
                s = s & "FormatPhone(" & QUOTE & QUOTE & " & !" & FNames(i).Name & ")"
              Else
                s = s & QUOTE & QUOTE & " & !" & FNames(i).Name
              End If
        Else
          s = s & QUOTE & QUOTE & " & !" & FNames(i).Name
        End If
        
        's = s & "!" & FNames(i).Name
        If i < j Then s = s & ", "
      Next
      Print #iFile, s
      Print #iFile, "        .MoveNext"
      Print #iFile, "      Loop"
      Print #iFile, "      .Close"
      Print #iFile, "    End With"
      Print #iFile, "    Set RS = Nothing"
      Print #iFile, ""
      Print #iFile, "End Sub"
      Print #iFile, ""
      
'*******************************************************
'****           Create the Update Method            ****
'*******************************************************
      Print #iFile, "Public Sub Update(ByVal Idx as Integer)"
      Print #iFile, "  Dim s As String"
      Print #iFile, ""
      Print #iFile, "    If Idx <> 0 Then"
      If mContinue Then
        Print #iFile, "'----[ Because we are using line continuations, you   ]----"
        Print #iFile, "'----[ may have to consolidate these before loading   ]----"
        Print #iFile, "'----[ this class into the environment.               ]----"
      End If
      Print #iFile, ""
      s = "      s = " & QUOTE & "UPDATE " & Tbl & " SET "
      T = 2
      For i = 1 To j
        DoEvents
        Select Case FNames(i).Type
          Case adBoolean
            N = FNames(i).Name & " = " & QUOTE & CC & "Abs(mcol(idx)." & FNames(i).Name & ")"
          Case adTinyInt, adUnsignedTinyInt, adSmallInt, adUnsignedSmallInt, adInteger, adUnsignedInt, adSingle, adDecimal, adDouble, adNumeric, adCurrency
            N = FNames(i).Name & " = " & QUOTE & CC & "Val(mcol(Idx)." & FNames(i).Name & ")"
          Case Else
            If InStr(UCase$(FNames(i).Name), "PHONE") <> 0 Or _
               InStr(UCase$(FNames(i).Name), "FAX") <> 0 Or _
               InStr(UCase$(FNames(i).Name), "CELL") <> 0 Or _
               InStr(UCase$(FNames(i).Name), "PAGER") <> 0 And _
               InStr(UCase$(FNames(i).Name), "EXT") = 0 Then
              N = FNames(i).Name & " = '" & QUOTE & CC & "StripPhone(mcol(Idx)." & FNames(i).Name & ")" & CC & QUOTE & "'"
            Else
              N = FNames(i).Name & " = '" & QUOTE & CC & "mcol(Idx)." & FNames(i).Name & CC & QUOTE & "'"
            End If
        End Select
        If Len(s) = 0 Then
          s = QUOTE & N
        Else
          If (Right$(s, 1) <> QUOTE) And (Right$(s, 1) <> "'") Then
            If Trim$(Right$(s, 5)) = "SET" Then
              s = s & N
            Else
              s = s & CC & QUOTE & COMMA & N
            End If
          ElseIf Right$(s, 1) = "'" Then
            s = s & " " & N
          ElseIf Right$(s, 1) = QUOTE Then
            s = s & CC & N
          Else
            s = s & N
          End If
        End If
        If mContinue Then
          If i < j And T > 1 Then
            T = 0
            If Right$(s, 1) = "'" Then
              Print #iFile, s; COMMA; QUOTE; CC; "_"
            Else
              Print #iFile, s; CC; QUOTE; COMMA; QUOTE; CC; "_"
            End If
            Print #iFile, Space$(10);
            s = ""
          ElseIf i = j Then
            If Right$(s, 1) = "'" Then
              Print #iFile, s; QUOTE; CC; "_"
            Else
              Print #iFile, s; CC; "_"
            End If
          End If
        Else
          If i < j And T > 1 Then
            T = 0
            If Right$(s, 1) = "'" Then
              Print #iFile, s; COMMA; QUOTE
            Else
              Print #iFile, s; CC; QUOTE; COMMA; QUOTE
            End If
            Print #iFile, Space$(6); "s = s & ";
            s = ""
          ElseIf i = j Then
            If Right$(s, 1) = "'" Then
              Print #iFile, s; QUOTE
            Else
              Print #iFile, s
            End If
          End If
        End If
        T = T + 1
      Next
      If mContinue Then
        Print #iFile, Space$(10) & QUOTE & " Where " & FNames(0).Name & " = " & QUOTE & " & mcol(Idx)." & FNames(0).Name
      Else
        Print #iFile, Space$(6); "s = s & " & QUOTE & " Where " & FNames(0).Name & " = " & QUOTE & " & mcol(Idx)." & FNames(0).Name
      End If
      Print #iFile, "      DbCars.Execute s"
      Print #iFile, "    End If"
      Print #iFile, "End Sub"
      Print #iFile, ""
    
'*******************************************************
'****     Create the Class_Initilaize Method        ****
'*******************************************************
      Print #iFile, "Private Sub Class_Initialize()"
      Print #iFile, "    Set mCol = New Collection"
      Print #iFile, "    Read"
      Print #iFile, "End Sub"
      Print #iFile, ""
      
'*******************************************************
'****     Create the Class_Terminate Method         ****
'*******************************************************
      Print #iFile, "Private Sub Class_Terminate()"
      Print #iFile, "    Set mCol = Nothing"
      Print #iFile, "End Sub"
      
      Close iFile
      lblGen.Caption = ""
      lblStatus.Caption = ""
      Me.MousePointer = vbDefault
      s = "Class file has been generated:" & vbCrLf & vbCrLf & sFile
      MsgBox s, vbOKOnly, "Class Generator"
      
    End If
End Sub

Private Sub cmdDB_Click()
    On Local Error Resume Next
    If con.State = 1 Then con.Close
    On Local Error GoTo 0
    frmDB.Show vbModal
    LoadTables
End Sub

Private Sub Form_Load()
      
    QUOTE = Chr$(34)
    CC = " & "
    COMMA = ", "
    SQ = QUOTE & "'" & QUOTE
    mContinue = True
    LoadTables
End Sub

Private Sub Form_Unload(Cancel As Integer)
    con.Close
    Set con = Nothing
End Sub

'******************************************************************
' lstTables_Click
'
'   Selects the table to create a class and/or collection for
'   This routine will strip any prefixed "tbl" from the table
'   name. Then prefix "c" in to the table name. This can be
'   edited in the table name text box, if desired.
'******************************************************************
Private Sub lstTables_Click()
  Dim s As String
  Dim i As Integer
    s = LCase$(Left$(lstTables.List(lstTables.ListIndex), 3))
    If s = "tbl" Then
      s = "c" & Mid$(lstTables.List(lstTables.ListIndex), 4)
    Else
      s = "c" & lstTables.List(lstTables.ListIndex)
    End If
  
  Do
    i = InStr(s, " ")
    If i <> 0 Then
      s = Left$(s, i - 1) & Mid$(s, i + 1)
    End If
  Loop Until i = 0
  
  'Strip any trailing s/S off the table name
  ' "s" will automatically be added for collections
  If Right$(s, 1) = "s" Or Right$(s, 1) = "S" Then s = Left$(s, Len(s) - 1)
  txtClass.Text = s
End Sub

Private Sub LoadTables()
  Dim cat As ADOX.Catalog
  Dim tb As ADOX.Table
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = con
    lstTables.Clear
    For Each tb In cat.Tables
      If tb.Type = "TABLE" Then lstTables.AddItem tb.Name
    Next
    Set tb = Nothing
    Set cat = Nothing
End Sub

Private Sub optLine_Click(Index As Integer)
    mContinue = optLine(0).Value
End Sub
