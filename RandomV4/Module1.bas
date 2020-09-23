Attribute VB_Name = "Module1"

'API and file type structure and misc. routine functions

Option Explicit

'Record data structure
Type tRec
id As Double
lName As String * 30
fName As String * 30
End Type

Rem - API Decs for 16 bit and 32 bit computers
Rem - This can be found all over the net
#If Win16 Then
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
#Else
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Function ReadINI(Section, KeyName, filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function
Function writeINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Sub CenterChild(Parent As Form, Child As Form)
Rem - Center the child window in the MDI window upon load
    Dim iTop As Integer
    Dim iLeft As Integer
    If Parent.WindowState = 0 Then Exit Sub
    iTop = ((Parent.Height - Child.Height) \ 2)
    iLeft = ((Parent.Width - Child.Width) \ 2)
    Child.Move iLeft, iTop
End Sub

