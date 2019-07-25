VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "attrb.exe by Meij"
   ClientHeight    =   1905
   ClientLeft      =   8040
   ClientTop       =   3120
   ClientWidth     =   5775
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5775
   Visible         =   0   'False
   Begin VB.CommandButton cmdlist 
      Caption         =   "Ls>"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
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
      ItemData        =   "Form1.frx":030A
      Left            =   1680
      List            =   "Form1.frx":0320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdunhide 
      Caption         =   "&Execute"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   1575
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
      Left            =   1680
      TabIndex        =   1
      Top             =   255
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Attribute"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "*Copy this app to desire directory where folder is seems to be exists"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Folder Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdlist_Click()
Dim fso As New FileSystemObject
Dim fld As Folder
Dim subfoldr As Folder


On Error GoTo errpage
Set fld = fso.GetFolder(App.Path)
Debug.Print App.Path

dlglistsub.Show
dlglistsub.Top = Form1.Top + Form1.Height
dlglistsub.Left = Form1.Left
dlglistsub.Text1.Text = ""


For Each subfoldr In fld.SubFolders
    Debug.Print subfoldr.Name & "-" & subfoldr.Attributes
    dlglistsub.Text1.Text = dlglistsub.Text1.Text & subfoldr.Name & " - " & subfoldr.Attributes & vbCrLf
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

