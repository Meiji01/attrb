VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "attrb 2"
   ClientHeight    =   2085
   ClientLeft      =   8040
   ClientTop       =   3120
   ClientWidth     =   5775
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5775
   Visible         =   0   'False
   Begin VB.TextBox txtLabel 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":030A
      Top             =   1200
      Width           =   5535
   End
   Begin VB.CommandButton cmdlist 
      Caption         =   "Ls>"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Timer tmanim 
      Interval        =   1000
      Left            =   4440
      Top             =   840
   End
   Begin VB.ComboBox cmbattr 
      Height          =   315
      ItemData        =   "Form1.frx":0381
      Left            =   2040
      List            =   "Form1.frx":0397
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdunhide 
      Caption         =   "&Change Attribute"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtfoldername 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "New Attribute"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Enter File/Folder Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
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

'Public f_attrib() As String

Private Sub cmdlist_Click()
Dim fso As New FileSystemObject
Dim fld As Folder
Dim subfoldr As Folder


On Error GoTo errpage
Set fld = fso.GetFolder(App.Path)
Debug.Print App.Path

If dlglistsub.Visible = False Then
    dlglistsub.Show
Else
    Unload dlglistsub
End If


mdlpublic.followmainwindow
dlglistsub.lstitems.Clear
Erase mdlpublic.f_attrib

'Create array to hold the attributes

Dim i As Integer
'Dim f_attrib() As String

'count subfolders
Dim foldercount As Integer
foldercount = fld.SubFolders.Count
ReDim f_attrib(foldercount, 3)

'Iterate folders
i = 0
For Each subfoldr In fld.SubFolders
    Debug.Print ("Redim Array" & i + 1)
    Debug.Print "Array:" & i + 1 & "," & 1 & "=" & subfoldr.Name
    
    dlglistsub.lstitems.AddItem (subfoldr.Name) 'Add folder name in the listbox
    
    
    '0= folder name
    '1=Attribute
    mdlpublic.f_attrib(i, 0) = subfoldr.Name
    mdlpublic.f_attrib(i, 1) = subfoldr.Attributes
    mdlpublic.f_attrib(i, 2) = mdlpublic.getFolderSize(subfoldr.Path)
    mdlpublic.f_attrib(i, 3) = "Folder"
    i = i + 1
Next

'Iterate Files
Dim filecount As Integer
filecount = fld.Files.Count
ReDim mdlpublic.file_attrib(filecount, 3) ' create placeholder for the file attributes


Dim file As Object
Dim j As Integer
j = 0
For Each file In fld.Files
    dlglistsub.lstitems.AddItem (file.Name)
    mdlpublic.file_attrib(j, 0) = file.Name
    mdlpublic.file_attrib(j, 1) = file.Attributes
    mdlpublic.file_attrib(j, 2) = file.Size
    mdlpublic.file_attrib(j, 3) = "File"
    j = j + 1
Next

Exit Sub
errpage:
MsgBox Err.Description, vbCritical, "Error!"
'MsgBox App.Path

End Sub

Private Sub cmdunhide_Click()
Dim apdir As String
Dim fdname As String
fdname = txtfoldername.Text
apdir = App.Path & "\" & fdname

Debug.Print apdir
'Print apdir

On Error GoTo errpage

'set folder attribute
If cmbattr.Text = "Archive" Then
    SetAttr fdname, vbArchive
ElseIf cmbattr.Text = "Hidden" Then
    SetAttr fdname, vbHidden
ElseIf cmbattr.Text = "System" Then
    SetAttr fdname, vbSystem
ElseIf cmbattr.Text = "System(ReadOnly)" Then
    SetAttr fdname, vbReadOnly
ElseIf cmbattr.Text = "Default" Then
    SetAttr fdname, vbNormal
ElseIf cmbattr.Text = "System,Hidden" Then
    SetAttr fdname, vbSystem + vbHidden
Else
    Err.Description = "Invalid Operation!"
    GoTo errpage
    End If
    Exit Sub
errpage:
    MsgBox Err.Description, vbCritical, "ERROR!"
End Sub


Private Sub Form_Click()
'Temporary only
'Debug.Print getAttribValue(17)
End Sub

Private Sub Form_Load()
cmbattr.Text = "Default"
cmdunhide.Enabled = False

'initialize variables
mdlpublic.initpublic
Call getfilesystem
'Call getfs2
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub tmanim_Timer()

Form1.Visible = True
tmanim.Interval = 40

mdlpublic.ysize = mdlpublic.ysize + 300
Form1.Height = mdlpublic.ysize

mdlpublic.ypos = mdlpublic.ypos + 200
Form1.Top = mdlpublic.ypos


If Form1.Height >= 2385 Or Form1.Top >= 3000 Then
    tmanim.Enabled = False
End If


End Sub

Private Sub txtfoldername_Change()
If txtfoldername.Text = "" Then
    cmdunhide.Enabled = False
    Else
    cmdunhide.Enabled = True
End If
End Sub

Sub getfilesystem()
On Error GoTo myerr
Dim fs As Object
Dim d As Object
Dim apdir As String

apdir = Left(App.Path, 2)
Debug.Print "app drive: " & apdir

Set fs = CreateObject("Scripting.FileSystemObject")
Set d = fs.GetDrive(apdir)
Debug.Print d.FileSystem
Form1.Caption = Form1.Caption & " - " & d.FileSystem
Exit Sub
myerr:
MsgBox "Unable to get drive FileSystem!", vbOKOnly, "Error1"
End Sub

