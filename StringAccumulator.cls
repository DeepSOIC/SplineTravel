VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringAccumulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Buf As String
Dim nRealChars As Long
Const INIT_SIZE = 128
Const IncMult As Double = 1.5

Public Sub Clear()
Buf = Space$(INIT_SIZE)
nRealChars = 0
End Sub

Public Property Get Content() As String
Content = Left$(Buf, nRealChars)
End Property

Public Property Let Content(ByRef NewStr As String)
Buf = NewStr
nRealChars = Len(Buf)
End Property

Public Sub Append(ByRef StringToAppend As String)
Dim need As Long
need = nRealChars + Len(StringToAppend)
If need > Len(Buf) Then
  Buf = Buf + Space$(MaxL(need - Len(Buf), CLng(Len(Buf) * IncMult)))
End If
Mid$(Buf, nRealChars + 1, Len(StringToAppend)) = StringToAppend
nRealChars = nRealChars + Len(StringToAppend)
End Sub

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
If a > b Then MaxL = a Else MaxL = b
End Function

Public Property Get Length() As Long
Length = nRealChars
End Property
