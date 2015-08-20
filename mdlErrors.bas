Attribute VB_Name = "mdlErrors"
Option Explicit

Const DisableAssertions As Boolean = True

Public Type vtError
    Number As Long
    Source As String
    Description As String
End Type

Private ProjectName As String
Private ErrorStack() As vtError
Private nInStack As Long


Public Enum eErrors
  'project-specific errors
  errZeroTimeMove = 12345
  errTooSlow = 12346
  errClassNotInitialized = 12347
  errInvalidCommand = 12348
  errNotInChain = 12349
  errAlreadyInChain = 12350
  errVerificationFailed = 12351
  errWrongConfigLine = 12352
  
  'standard errors
  errCancel = 32755
  errIndexOutOfRange = 9
  errInvalidArgument = 5
  errWrongType = 13
End Enum
 
Public Sub Throw(Optional er As eErrors = 0, Optional Source As String, Optional extraMessage As String)
If er = 0 Then
  'analog of C++ throw, re-raise the handled error
  ErrRaise
End If

Dim Message As String
Select Case er
  'project-specific errors
  Case errZeroTimeMove
    Message = "Zero time move"
  Case errTooSlow
    Message = "Too slow, it doesn't make sense to fit a spline"
  Case errClassNotInitialized
    Message = "Class is not initialized properly"
  Case errInvalidCommand
    Message = "G-Command is invalid"
  Case errNotInChain
    Message = "Command is not in chain"
  Case errAlreadyInChain
    Message = "Command is already in a chain"
  Case errVerificationFailed
    Message = ""
  Case errWrongConfigLine
    Message = "Wrong config line"
    
  'standard errors
  Case errCancel
    Message = "Canceled by user"
  Case Else
    'hack - obtain standard error message.
    On Error Resume Next
    Err.Raise er
    Message = Err.Description
    On Error GoTo 0
End Select

If Len(extraMessage) > 0 Then
  Message = Message + ", " + extraMessage
End If
Err.Raise er, Source, Message
End Sub


Public Sub ReadError_Arg(ByRef vErr As vtError)
vErr.Number = Err.Number
vErr.Source = Err.Source
vErr.Description = Err.Description
End Sub

Public Function ReadError() As vtError
Dim vErr As vtError
ReadError_Arg vErr
ReadError = vErr
End Function

Public Sub vtRaiseError(ByRef aErr As vtError, _
                      Optional ByVal ProcedureName As String)
If Len(ProjectName) = 0 Then
    GetProjectName
End If
Debug.Assert aErr.Number = eErrors.errCancel Or DisableAssertions
If Len(ProcedureName) > 0 Then
    If aErr.Source = ProjectName Then
        Err.Raise aErr.Number, ProcedureName, aErr.Description
    Else
        Err.Raise aErr.Number, aErr.Source, aErr.Description
    End If
Else
    Err.Raise aErr.Number, aErr.Source, aErr.Description
End If
End Sub

Private Function GetProjectName()
PushError

'hack.
On Error Resume Next
Err.Raise 5
ProjectName = Err.Source

PopError
End Function

Public Sub RaiseError(Optional ByVal ProcedureName As String)
Dim vErr As vtError
ReadError_Arg vErr
vtRaiseError vErr, ProcedureName
End Sub

Public Sub ErrRaise(Optional ByVal ProcedureName As String)
RaiseError ProcedureName
End Sub

'displays an error message, according to info in global Err
Function MsgError(Optional ByVal Message As Variant = "", _
                  Optional ByVal Style As VbMsgBoxStyle = vbCritical, _
                  Optional ByVal Assertion As Boolean = False) As VbMsgBoxResult
Dim strMessage As String
Dim bAddErrDesc As Boolean
PushError
If Err.Number = eErrors.errCancel Then
    MsgError = vbCancel
    Exit Function
End If
On Error Resume Next
  strMessage = CStr(Message)
On Error GoTo 0
PopError
If Len(strMessage) = 0 Then
    strMessage = Err.Description
ElseIf InStr(1&, strMessage, "Err.Description", vbTextCompare) Then
    strMessage = Replace(strMessage, "Err.Description", Err.Description, Compare:=vbTextCompare)
Else
    strMessage = strMessage + vbNewLine + "(" + Err.Description + ")"
End If
MsgError = MsgBox(strMessage, Style)
Debug.Assert Assertion Or DisableAssertions
End Function

'pushes the error info (from global Err) into a stack for later re-use.
'PushError/PopError are useful when doing cleanup in case of
'an error, before raising it.
Public Sub PushError()
Dim vErr As vtError
ReadError_Arg vErr
If nInStack = 0 Then
    ReDim ErrorStack(0 To 0)
Else
    ReDim Preserve ErrorStack(0 To nInStack)
End If
ErrorStack(nInStack) = vErr
nInStack = nInStack + 1
End Sub

'pops the error info (writes to global Err object) from the error stack.
'PushError/PopError are useful when doing cleanup in case of
'an error, before raising it.
Public Sub PopError(Optional ByVal RaiseIt As Boolean = False)
Dim vErr As vtError
nInStack = nInStack - 1
If nInStack >= 0 Then
    vErr = ErrorStack(nInStack)
Else
    nInStack = 0
End If
If nInStack > 0 Then
    ReDim Preserve ErrorStack(0 To nInStack - 1)
Else
    Erase ErrorStack
End If
If Not RaiseIt Then On Error Resume Next
Err.Raise vErr.Number, vErr.Source, vErr.Description
End Sub

'puts the error into the global Err object, but doesn't raise it
Public Sub WriteError(ByVal ErrNumber As Long, _
                      ByVal ErrSource As String, _
                      ByVal ErrDescription As String)
On Error Resume Next
Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub




