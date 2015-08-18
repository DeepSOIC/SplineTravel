Attribute VB_Name = "mdlFiles"
Option Explicit
'dependencies:
'* StringAccumulator.cls

'adds trailing backslash to path, if one is missing
Public Function ValFolder(ByRef strFolder As String) As String
If right$(strFolder, 1) = "\" Then
    ValFolder = strFolder
Else
    ValFolder = strFolder + "\"
End If
End Function

'returns list of full paths to files contained in folders specified in paths()
'folders specified in paths() must end with a backslash!
Public Function getListOfFiles(paths() As String, Optional matchString As String = "*") As String()
Dim iPath As Long
Dim files As New StringAccumulator
For iPath = 0 To UBound(paths)
  Dim fn As String
  fn = Dir(paths(iPath) + matchString)
  Do While Len(fn) > 0
    files.Append paths(iPath) + fn + vbNewLine
    fn = Dir()
  Loop
Next iPath
files.Backspace Len(vbNewLine)
getListOfFiles = split(files.content, vbNewLine)
End Function

'/common

'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'SplineTravel specific...

Function PresetsPaths() As String()
Static ret() As String
Static initialized As Boolean
'If Not initialized Then 'strange... ret is lost after each call, it shouldn't...
  ReDim ret(0 To 0)
  ret(0) = ValFolder(App.path) + "presets\"
  initialized = True
'End If
PresetsPaths = ret
End Function



'Example: D:\noname.bmp -> noname.bmp
Function GetFileName(strPath As String) As String
Dim Pos As Long
Pos = InStrRev(strPath, "\")
GetFileName = Mid$(strPath, Pos + 1)
End Function

'Example: D:\noname.bmp -> D:\noname
Function CropExt(path As String) As String
Dim dotPos As Long
Dim FilenameStartPos As Long
FilenameStartPos = InStrRev(path, "\") + 1
dotPos = InStrRev(path, ".")
If dotPos < FilenameStartPos Then dotPos = Len(path) + 1
If dotPos = 0 Then dotPos = Len(path) + 1
CropExt = Left$(path, dotPos - 1)
End Function

Function getFileTitle(path As String) As String
Dim dotPos As Long
Dim FilenameStartPos As Long
FilenameStartPos = InStrRev(path, "\") + 1
dotPos = InStrRev(path, ".")
If dotPos < FilenameStartPos Then dotPos = Len(path) + 1
If dotPos = 0 Then dotPos = Len(path) + 1
getFileTitle = Mid$(path, FilenameStartPos, dotPos - FilenameStartPos)
End Function
