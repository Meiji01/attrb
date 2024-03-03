Attribute VB_Name = "mdlpublic"
Public ysize As Integer
Public ypos As Integer


'0= folder name
'1=Attribute
Public f_attrib() As String

Sub initpublic()
ysize = 0
ypos = 0
End Sub


Public Sub followmainwindow() ' this sub is to relocate windows when moved around the screen
dlglistsub.Top = Form1.Top + Form1.Height
dlglistsub.Left = Form1.Left
End Sub
