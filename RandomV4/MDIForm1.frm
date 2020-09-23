VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "DB Dialog SDI"
   ClientHeight    =   4950
   ClientLeft      =   4140
   ClientTop       =   3675
   ClientWidth     =   8220
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0442
            Key             =   "B1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":05EA
            Key             =   "B2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   794
      ButtonWidth     =   2858
      ButtonHeight    =   688
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open DB Dialog"
            Key             =   "B1"
            Object.ToolTipText     =   "Open DB Dialog"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close DB Dialog"
            Key             =   "B2"
            Object.ToolTipText     =   "Close DB Dialog"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open DB Dialog"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple MDI form with toolbar and menubar
'Parent form to child forms


Dim ret As Integer
Private Sub MDIForm_Activate()

ret = Module1.writeINI("Uname", "username", "student1", App.Path & "\test.ini")
ret = Module1.writeINI("Pword", "password", "pass1", App.Path & "\test.ini")
ret = Module1.writeINI("LIN", "login", "no", App.Path & "\test.ini")

Form2.Show
End Sub

Private Sub mnuExit_Click()
ret = Module1.writeINI("LIN", "login", "no", App.Path & "\test.ini")
End
End Sub

Private Sub mnuOpen_Click()
Form1.Show
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim logVal As String
logVal = Module1.ReadINI("LIN", "login", App.Path & "\test.ini")
    Select Case Button.Key
        Case "B1"
            If logVal = "yes" Then
            Unload Form2
            Form1.Show
            Else
            MsgBox "You must log in"
            End If
        Case "B2"
        Unload Form1
    End Select
End Sub
