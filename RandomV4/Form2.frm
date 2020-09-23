VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Db Login Example"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Login"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1920
      PasswordChar    =   "^"
      TabIndex        =   6
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create File"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   3000
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Username:"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the DB Random Access Test Database Example."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple login form with reading and writing to an ini file along with
'file creation script/progress bar.

Dim i As Double

Private Sub Command1_Click()

Dim rt As Module1.tRec
'Delete existing file
If Dir("c:\test32.txt") <> "" Then
Kill "c:\test32.txt"
End If

ProgressBar1.Max = 1000
ProgressBar1.Min = 0
ProgressBar1.Enabled = True


'Recreate it with "random" data
Open "c:\test32.txt" For Random As #1 Len = Len(rt)
For i# = 1 To 1000
'Could Use this to generate a very random/unique ID
rt.id = CInt(i#)
rt.lName = "First Name" & CStr(i)
rt.fName = "Last Name" & CStr(i)
Put #1, i#, rt
ProgressBar1.Value = i#
Next

'Rewind back to first record/byte in file
Seek #1, 1
Close #1
Form1.Show
Unload Form2

End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0: loginMe
Case 1: End
End Select
End Sub

Sub loginMe()
If Text1(0).Text = "" Then MsgBox "Username required": Exit Sub
If Text1(1).Text = "" Then MsgBox "Password required": Exit Sub
If Text1(0).Text = "" And _
Text1(1).Text = "" Then MsgBox "Username and Password required": Exit Sub

Dim retVal(0 To 1) As String

retVal(0) = Module1.ReadINI("Uname", "username", App.Path & "\test.ini")
retVal(1) = Module1.ReadINI("Pword", "password", App.Path & "\test.ini")

If Text1(0).Text <> retVal(0) Then MsgBox "Incorrect username": Exit Sub
If Text1(1).Text <> retVal(1) Then MsgBox "Incorrect password": Exit Sub

ret = Module1.writeINI("LIN", "login", "yes", App.Path & "\test.ini")

Form2.Height = 4350
Module1.CenterChild MDIForm1, Form2
End Sub

Private Sub Form_Load()

Module1.CenterChild MDIForm1, Form2
Form2.Height = 3090

End Sub

