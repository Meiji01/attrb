VERSION 5.00
Begin VB.Form dlglistsub 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sub Folders Exists"
   ClientHeight    =   1650
   ClientLeft      =   8040
   ClientTop       =   5805
   ClientWidth     =   5775
   Icon            =   "dlglistsub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "dlglistsub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
Unload Me
End Sub
