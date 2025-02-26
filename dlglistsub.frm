VERSION 5.00
Begin VB.Form dlglistsub 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Files and Folders"
   ClientHeight    =   3465
   ClientLeft      =   8040
   ClientTop       =   5805
   ClientWidth     =   5775
   Icon            =   "dlglistsub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.DriveListBox drvbox 
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Frame frameProperties 
      Caption         =   "Properties"
      Height          =   3135
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.ListBox lstitems 
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Close"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2880
      Width           =   3015
   End
End
Attribute VB_Name = "dlglistsub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at

'    http://www.apache.org/licenses/LICENSE-2.0

'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.


Option Explicit


Private Sub cmdRefresh_Click()
Form1.printFolderList
End Sub

Private Sub drvbox_Change()

Dim newdrive As String

newdrive = drvbox.Drive


If InStr(newdrive, "[\\") = 0 Then
    mdlpublic.curdir = newdrive & "\"
    Else
    'get in between []
    Dim strdrive As String
    strdrive = mdlpublic.extractRegex(newdrive, "\\.+")
    strdrive = Left(strdrive, Len(strdrive) - 1)
    mdlpublic.curdir = strdrive
End If

Call cmdRefresh_Click
Debug.Print "Init Drive:" & mdlpublic.curdir
'Print drvbox.Drive
End Sub

Private Sub Form_Click()
'dirGoBack (mdlpublic.curdir)
'Debug.Print mdlpublic.extractRegex("test123", "\d+")
End Sub

Private Sub Form_Deactivate()
mdlpublic.followmainwindow
End Sub

Private Sub lstitems_Click()
'txtProperties.Text = txtProperties.Text & "Attribute:" & "Test"
'Debug.Print UBound(mdlpublic.f_attrib)
txtProperties.text = ""
If lstitems.ListIndex > 0 Then
    Debug.Print "Selected Item is"; lstitems.text
    findfolder lstitems.text, mdlpublic.f_attrib
    findfolder lstitems.text, mdlpublic.file_attrib
    Form1.txtfoldername.text = lstitems.text
Else
    txtProperties.text = "<Goto Previous Folder>"
    Form1.txtfoldername.text = ""
End If
End Sub

Private Sub lstitems_DblClick()
If lstitems.ListIndex = 0 Then
    dirGoBack (mdlpublic.curdir)
    Form1.printFolderList
Else
    On Error GoTo skipcheck
    If mdlpublic.f_attrib(lstitems.ListIndex, 3) = "Folder" Then
        dirGoUp (lstitems.text)
        Form1.printFolderList
     End If


skipcheck:
Debug.Print "Not a folder"
Exit Sub
    'Do nothing
End If

End Sub

Private Sub OKButton_Click()
Unload Me
End Sub

Private Sub findfolder(foldername As String, attribs() As String)

Dim o_attributes As String
Dim X

For X = 0 To UBound(attribs)
    Debug.Print "findfolder:" & foldername
    If attribs(X, 0) = foldername Then
        'Print ("Attribute:" & attribs(x, 1))
        txtProperties.text = txtProperties.text & "Name: " & attribs(X, 0) & vbCrLf
        txtProperties.text = txtProperties.text & "Type: " & attribs(X, 3) & vbCrLf
        'txtProperties.Text = txtProperties.Text & "Attribute: " & attribs(x, 1) & vbCrLf
        txtProperties.text = txtProperties.text & "Attribute: " & mdlpublic.getAttribValue(Val(attribs(X, 1))) & vbCrLf
        txtProperties.text = txtProperties.text & "Size: " & attribs(X, 2) & " bytes" & vbCrLf
    End If
Next X

End Sub

Private Function dirTrimDown(dir As String) As String

Dim curdir As String
curdir = dir

'sample
'C:\WorkingFOlder\Folder1

'Trim dir down
Dim newdir As String
Dim lastdirtextposition As Integer

lastdirtextposition = InStrRev(curdir, "\") - 1
newdir = Left(curdir, lastdirtextposition)
Debug.Print ("CurrentDir:" & dir)
Debug.Print ("NewDir:" & newdir)

If Len(newdir) < 3 And Right(newdir, 1) <> "\" Then ' handles if the currentdir is on root directory
    newdir = newdir & "\"
End If
dirTrimDown = newdir
End Function

Private Sub dirGoBack(dir As String)
Dim newdir As String
If Len(mdlpublic.curdir) > 3 Then
    newdir = dirTrimDown(dir)
    'Form1.lblcurdir = newdir
    mdlpublic.curdir = newdir
End If
End Sub

Private Sub dirGoUp(foldername As String)
Dim newdir As String
newdir = mdlpublic.curdir & "\" & foldername
Form1.lblcurdir = newdir
mdlpublic.curdir = newdir
End Sub
