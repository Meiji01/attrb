Attribute VB_Name = "mdlpublic"
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at

'    http://www.apache.org/licenses/LICENSE-2.0

'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.


Public ysize As Integer
Public ypos As Integer
Public curdir As String


'0= folder name
'1=Attribute
'2=Folder Size
Public f_attrib() As String
Public file_attrib() As String

Sub initpublic()
ysize = 0
ypos = 0
End Sub


Public Sub followmainwindow() ' this sub is to relocate windows when moved around the screen
dlglistsub.Top = Form1.Top + Form1.Height
dlglistsub.Left = Form1.Left
End Sub

'Scripting File System Object module to get Folder Size
Public Function getFolderSize(folderpath As String)
Dim fso As FileSystemObject

Set fso = New FileSystemObject
getFolderSize = fso.GetFolder(folderpath).Size
End Function

Public Function getAttribValue(attrib As Integer) As String

Dim X As Integer
Dim bit As Integer
Dim attribstring As String
'Dim currbit As Integer


Dim modulus As Integer
modulus = 0

If attrib > 0 Then

Do
    For X = 0 To 5
        bit = 2 ^ X
        If bit = 0 Then
            quotient = 1
        Else
            quotient = attrib \ bit
            modulus = attrib Mod bit
        End If
        
        Debug.Print "Bit:" & bit & " Quotient:" & quotient & " Remainder:" & modulus
        
        If quotient = 1 Then
            If attribstring <> "" Then
                attribstring = attribstring & ", "
            End If
            attribstring = attribstring & getAttribText(bit)
            Exit For
        End If

    Next X
    
    'subtract attrib in remainder
    attrib = attrib - bit
Loop While modulus > 0
    Debug.Print "New Attrib:" & attrib
End If


'return attrib
getAttribValue = attribstring

End Function

Private Function getAttribText(attrib As Integer) As String

Const normal = 0
Const readOnly = 1
Const hidden = 2
Const system = 4
Const directory = 16
Const archive = 32
Dim attribvalue As String



Select Case attrib
Case normal
attribvalue = "Default"
Case readOnly
attribvalue = "Read-Only"
Case hidden
attribvalue = "Hidden"
Case system
attribvalue = "System"
Case directory
attribvalue = "Directory"
Case archive
attribvalue = "Archive"
Case Else
attribvalue = "Unknown"
End Select

'return the attribvalue
getAttribText = attribvalue

End Function

