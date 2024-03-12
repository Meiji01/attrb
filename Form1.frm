VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "attrb 2"
   ClientHeight    =   1995
   ClientLeft      =   8040
   ClientTop       =   3120
   ClientWidth     =   5775
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5775
   Visible         =   0   'False
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
      Left            =   5280
      Top             =   840
   End
   Begin VB.ComboBox cmbattr 
      Height          =   315
      ItemData        =   "Form1.frx":030A
      Left            =   2040
      List            =   "Form1.frx":0320
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
   Begin VB.Label lblcurdir 
      Caption         =   "x:\"
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label lblDirtitle 
      Caption         =   "Current Directory:"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblVersion 
      Caption         =   "v x.x"
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Set Attribute:"
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

On Error GoTo errpage


If dlglistsub.Visible = False Then
    dlglistsub.Show
    mdlpublic.followmainwindow
    Call printFolderList
Else
    Unload dlglistsub
End If


Exit Sub
errpage:
MsgBox Err.Description, vbCritical, "Error!"
'MsgBox App.Path

End Sub

Private Sub cmdunhide_Click()
Dim apdir As String
Dim fdname As String
'fdname = txtfoldername.Text
'apdir = App.Path & "\" & fdname
apdir = mdlpublic.curdir
fdname = apdir & "\" & txtfoldername.text



Debug.Print fdname
'Print apdir

On Error GoTo errpage

'set folder attribute
If cmbattr.text = "Archive" Then
    SetAttr fdname, vbArchive
ElseIf cmbattr.text = "Hidden" Then
    SetAttr fdname, vbHidden
ElseIf cmbattr.text = "System" Then
    SetAttr fdname, vbSystem
ElseIf cmbattr.text = "System(ReadOnly)" Then
    SetAttr fdname, vbReadOnly
ElseIf cmbattr.text = "Default" Then
    SetAttr fdname, vbNormal
ElseIf cmbattr.text = "System,Hidden" Then
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

lblVersion.Caption = "v. " & App.Major & "." & App.Minor & "." & App.Revision
mdlpublic.curdir = App.Path
Form1.lblcurdir = mdlpublic.curdir

cmbattr.text = "Default"
cmdunhide.Enabled = False

'initialize variables
mdlpublic.initpublic
Call getfilesystem
'Call getfs2
End Sub



Private Sub Form_Unload(Cancel As Integer)
End
End Sub







Private Sub lblcurdir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblcurdir.ToolTipText = mdlpublic.curdir
End Sub

Private Sub tmanim_Timer()

Form1.Visible = True
tmanim.Interval = 100

mdlpublic.ysize = mdlpublic.ysize + 300
Form1.Height = mdlpublic.ysize

mdlpublic.ypos = mdlpublic.ypos + 200
Form1.Top = mdlpublic.ypos


If Form1.Height >= 2475 Or Form1.Top >= 3000 Then
    tmanim.Enabled = False
End If


End Sub

Private Sub txtfoldername_Change()
If txtfoldername.text = "" Then
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

Public Sub printFolderList() ' this method to print value to listbox in dlglistsub

On Error GoTo ErrTab

Dim fso As New FileSystemObject
Dim fld As Folder
Dim subfoldr As Folder

'Set fld = fso.GetFolder(App.Path)
Set fld = fso.GetFolder(mdlpublic.curdir)
Debug.Print App.Path


dlglistsub.lstitems.Clear
Erase mdlpublic.f_attrib

'update directory label in form1
Form1.lblcurdir = mdlpublic.curdir

'Create array to hold the attributes

Dim i As Integer
'Dim f_attrib() As String

'count subfolders
Dim foldercount As Integer
foldercount = fld.SubFolders.Count
ReDim f_attrib(foldercount, 3)

'Iterate folders
i = 1 'iteration starts at 1 since 0 is reserved for the ... "goto prev dir"

'info for array 0
dlglistsub.lstitems.AddItem ("...")
mdlpublic.f_attrib(0, 3) = "Marker"


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
ErrTab:
MsgBox Err.Description
End Sub

