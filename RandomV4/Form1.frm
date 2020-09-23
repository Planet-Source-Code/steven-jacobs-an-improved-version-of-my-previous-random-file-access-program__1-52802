VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Random File Example (Beta V1.0.4)"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   163
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   2070
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   495
      Left            =   3600
      TabIndex        =   14
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Delete Record"
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add Record"
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<<"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update Record"
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Record"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Record #:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "First Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ID:"
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple example of Random File Access/Writing/Reading
'Complete with navigation, record jumping, adding a record, etc.
'Added: 02/10/2004: Very Simple Deletion Routine.
'Added: 03/30/2004: MDI form and Simple Search functionality and error checking
'                   routine (valMe())
'Added: 04/01/2004: Login form and Username/Password and Login validator text
'                   to an *.ini/config file for no "preemptive" or invalid
'                   logins for main form.
'Author:  Steven Jacobs
'Date: 02/09/2004
'Feel free to learn and improve
'Beta Version 1.0.4


Dim countRecs As Integer
Dim rt As Module1.tRec
Dim fileNum As Integer
Dim maxRecords As Integer
Dim id() As Double
Dim lName() As String
Dim fName() As String
Dim foundMe As Boolean

Function valMe() As Boolean

Dim badString As String

badString = "~`!@#$%^&*()-_=+|\}{[]';:""/?><qwertyuioplkjhgfdsazxcvbnm "

For t% = 1 To Len(Text1.Text)
If InStr(1, UCase(badString), Mid(UCase(Text1.Text), t%, 1)) > 0 Or Asc(Mid(Text1.Text, t%, 1)) > 128 Then
MsgBox "Not a valid record number.  Please try again", vbCritical, "Record Number Error"
valMe = False
Exit For
Else
valMe = True
End If
Next

End Function

Private Sub Command1_Click()
'Get Record By Record Number Function

If Text1.Text = "" Or Not valMe Then Text1.Text = "": Exit Sub
If CInt(Text1.Text) = 0 Then
MsgBox "Record Number '0' is not a valid record", 16, "Invalid Record Error"
Text1.Text = ""
Exit Sub
ElseIf CInt(Text1.Text) > (maxRecords - 1) Then
MsgBox "There are only " & maxRecords - 1 & " valid records in the file", 16, "Record Error"
Text1.Text = ""
Exit Sub
Else
Open "c:\test32.txt" For Random As fileNum% Len = Len(rt)
Get #fileNum%, CInt(Text1.Text), rt
Text2.Text = rt.id
Text3.Text = rt.lName
Text4.Text = rt.fName
Close fileNum%
End If
 
End Sub


Private Sub Command2_Click()
'Search function for record in file if available

Dim lastName As String
Dim foundRec As Boolean
lastName = InputBox("Search for:", "Search")

If lastName = "" Then Exit Sub

foundRec = False

Open "c:\test32.txt" For Random As fileNum% Len = Len(rt)
For rn% = 1 To maxRecords - 1
   Get #fileNum%, rn%, rt
   If lastName = Trim(rt.fName) Then
        foundRec = True
        Exit For
    End If
  Next

If foundRec = True Then
   Get #fileNum, rn%, rt
   Text1.Text = CStr(rn%)
   Text2.Text = rt.id
   Text3.Text = rt.lName
   Text4.Text = rt.fName
Else
    MsgBox "Name " + lastName + " not found!"
End If

Close fileNum%

End Sub

Private Sub Command3_Click()
'Update Current Record function

If Text1.Text = "" Or Not valMe Then Text1.Text = "": Exit Sub

Open "c:\test32.txt" For Random As fileNum% Len = Len(rt)

'Rewind back to first record/byte in file
Seek fileNum%, 1

rt.id = Text2.Text
rt.lName = Text3.Text
rt.fName = Text4.Text
Put #fileNum%, CInt(Text1.Text), rt
Close fileNum%


MsgBox "Record #" & Text1.Text & " updated", 32, "Updated Record Notification"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

End Sub

Private Sub Command4_Click()
'Move to next record function

If Text1.Text = "" Or Not valMe Then Text1.Text = "": Exit Sub

If (CInt(Text1.Text) + 1) > (maxRecords - 1) Then
MsgBox "No more records in file", 16, "Last Record Reached"
Else
Open "c:\test32.txt" For Random As fileNum% Len = Len(rt)
Get #fileNum%, CInt(Text1.Text + 1), rt
Text1.Text = CInt(Text1.Text) + 1
Text2.Text = rt.id
Text3.Text = rt.lName
Text4.Text = rt.fName
Close fileNum%
End If
End Sub

Private Sub Command5_Click()
'Move to Previous Record function

If Text1.Text = "" Or Not valMe Then Text1.Text = "": Exit Sub

If CInt(Text1.Text) = 0 Then
MsgBox "Record Number '0' is not a valid record", 16, "Invalid Record Error"
Exit Sub
ElseIf CInt(Text1.Text) = 1 Then
MsgBox "You have reached the first record in this file", 16, "First Record Reached"
Exit Sub
Else
Open "c:\test32.txt" For Random As fileNum% Len = Len(rt)
Get #fileNum%, CInt(Text1.Text - 1), rt
Text1.Text = CInt(Text1.Text) - 1
Text2.Text = rt.id
Text3.Text = rt.lName
Text4.Text = rt.fName
Close fileNum%
End If
End Sub

Private Sub Command6_Click()
'Add Record Function

If Text1.Text <> "" Then
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Exit Sub
End If

Open "c:\test32.txt" For Random As fileNum% Len = Len(rt)

'Rewind back to first record/byte in file
Seek fileNum%, 1

rt.id = CInt((((i * Rnd(4) Xor 9) / Rnd(5)) * 4) * Rnd(2))
rt.lName = Text3.Text
rt.fName = Text4.Text
Put #fileNum%, (maxRecords - 1) + 1, rt
Close fileNum%


MsgBox "Record #" & (maxRecords - 1) + 1 & " Added To File", 32, "Added Record Notification"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

maxRecords = maxRecords + 1

StatusBar1.SimpleText = "There are " & maxRecords - 1 & " files in c:\test32.txt"

End Sub

Private Sub Command7_Click()
'Delete Record Function

If Text1.Text = "" Or Not valMe Then Text1.Text = "": Exit Sub

If CInt(Text1.Text) = 0 Then
MsgBox "Record Number '0' is not a valid record", 16, "Invalid Record Error"
Text1.Text = ""
Exit Sub
ElseIf CInt(Text1.Text) > (maxRecords - 1) Then
MsgBox "There are only " & maxRecords - 1 & " valid records in the file", 16, "Record Error"
Text1.Text = ""
Exit Sub
Else
foundMe = False
Open "c:\test32.txt" For Random As fileNum% Len = Len(rt)
For i% = 1 To maxRecords - 1
If CInt(Text1.Text) = i% Then
foundMe = True
tempC% = i% - 1 'reset counter for array
ReDim Preserve id(tempC%)
ReDim Preserve lName(tempC%)
ReDim Preserve fName(tempC%)
Get #fileNum%, i% + 1, rt
id(tempC%) = rt.id
lName(tempC%) = rt.lName
fName(tempC%) = rt.fName
Else
ReDim Preserve id(i% - 1)
ReDim Preserve lName(i% - 1)
ReDim Preserve fName(i% - 1)
If foundMe Then
Get #fileNum%, i% + 1, rt
Else
Get #fileNum%, i%, rt
End If
id(i% - 1) = rt.id
lName(i% - 1) = rt.lName
fName(i% - 1) = rt.fName
End If
Next

Close fileNum%

'delete existing file
Kill "c:\test32.txt"

'write out new data to file
Open "c:\test32.txt" For Random As fileNum% Len = Len(rt)
For i% = 1 To UBound(id)
rt.id = CInt(id(i% - 1))
rt.lName = lName(i% - 1)
rt.fName = fName(i% - 1)
Put #fileNum%, i%, rt
Next

'if we got this far, we have created the new file w/o the deleted record.
MsgBox "Record Deleted", 64, "Deletion Notification"

'Reset maxRecords count
maxRecords = UBound(id) + 1

'Set label5 caption with number of records n file
StatusBar1.SimpleText = "There are " & UBound(id) & " files in c:\test32.txt"

'Reset fields
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""


Close fileNum%
Erase id
Erase lName
Erase fName
End If



End Sub

Private Sub Form_Load()
'Create new file with test data and get "valid" record count

Unload Form2

Module1.CenterChild MDIForm1, Form1

countRecs = 0

fileNum% = FreeFile()

Open "c:\test32.txt" For Random As #1 Len = Len(rt)

'Get number of records on file
Do While Not EOF(fileNum%)
   Get #fileNum%, , rt
   countRecs = countRecs + 1
Loop

Get #fileNum%, 1, rt
Text1.Text = 1
Text2.Text = rt.id
Text3.Text = rt.lName
Text4.Text = rt.fName
maxRecords = countRecs

'Set StatusBar caption with number of records n file
StatusBar1.SimpleText = "There are " & countRecs - 1 & " files in c:\test32.txt"

Close fileNum%

End Sub


Private Sub mnuExit_Click()
Unload Me
End Sub
