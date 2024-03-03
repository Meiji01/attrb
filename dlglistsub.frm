VERSION 5.00
Begin VB.Form dlglistsub 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Folder List"
   ClientHeight    =   2115
   ClientLeft      =   8040
   ClientTop       =   5805
   ClientWidth     =   5775
   Icon            =   "dlglistsub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameProperties 
      Caption         =   "Properties"
      Height          =   1935
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   1335
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
      Height          =   1230
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   3015
   End
End
Attribute VB_Name = "dlglistsub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Deactivate()
mdlpublic.followmainwindow
End Sub

Private Sub lstitems_Click()
'txtProperties.Text = txtProperties.Text & "Attribute:" & "Test"
'Debug.Print UBound(mdlpublic.f_attrib)
txtProperties.Text = ""
Debug.Print "Selected Item is"; lstitems.Text
findfolder lstitems.Text, mdlpublic.f_attrib
Form1.txtfoldername.Text = lstitems.Text
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
        txtProperties.Text = txtProperties.Text & "Folder Name:" & attribs(X, 0) & vbCrLf
        txtProperties.Text = txtProperties.Text & "Attribute:" & attribs(X, 1) & vbCrLf
    End If
Next X

End Sub

