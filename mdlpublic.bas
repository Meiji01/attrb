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


'0= folder name
'1=Attribute
'2=Folder Size
Public f_attrib() As String

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

