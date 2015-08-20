Attribute VB_Name = "mdlCommon"
Option Explicit

Public Const Pi As Double = 3.14159265358979



Public Type typCurrentState
  Speed As Double
  Pos As typVector3D
  Epos As Double 'extrusion accumulator
  MoveRelative As Boolean
  ExtrusionRelative As Boolean
End Type

Function vtStr(val As Double) As String
vtStr = Trim(Str(val))
End Function


Function prepad(st As String, ByVal toLength As Long, _
                Optional ByVal padChar As String = "0") As String
If Len(st) > toLength Then
  Throw errInvalidArgument, "prepad", "pad can't shrink; the string """ + st + _
                            """ has more than " + CStr(toLength) + " characters."
End If
prepad = String$(toLength - Len(st), padChar) + st
End Function


Public Function EscapeString(st As String) As String
Dim outStr As New StringAccumulator
Dim i As Long
For i = 0 To Len(st) - 1
  Dim chcode As Long
  chcode = AscW(Mid$(st, i + 1, 1))
  If chcode < 32 Or chcode > 255 Or chcode = Asc("%") Then
    outStr.Append "%" + prepad(Hex$(chcode), 4)
  Else
    outStr.Append Mid$(st, i + 1, 1)
  End If
Next i
EscapeString = outStr.content
End Function

Public Function unEscapeString(st As String) As String
Dim outStr As New StringAccumulator
Dim i As Long
For i = 0 To Len(st) - 1
  Dim ch As String
  ch = Mid$(st, i + 1, 1)
  If ch = "%" Then
    outStr.Append ChrW(val("&H" + Mid$(st, i + 2, 4) + "&"))
    i = i + 4
  Else
    outStr.Append ch
  End If
Next i
unEscapeString = outStr.content
End Function


