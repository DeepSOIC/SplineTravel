VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'chain of g commands

Dim myFirst As clsGCommand
Dim myLast As clsGCommand

Dim mySize As Long

Const AUTO_VERIFY As Boolean = True 'set to true for debugging; will cause serious slowdown

Private Sub autoVerify(strMethod As String)
If Not AUTO_VERIFY Then Exit Sub
If Not verify Then
  Throw errVerificationFailed, "clsChain." + strMethod, _
    "Verifying the chain indicated that there is an error. See debug" + _
    " output for more info on what's wrong."
End If
End Sub

Public Sub Add(ByVal cmd As clsGCommand, Optional ByVal Before As clsGCommand = Nothing, Optional ByVal After As clsGCommand = Nothing)

If Not (Before Is Nothing) And Not (After Is Nothing) Then
  Throw errInvalidArgument, "clsChain::Add", extraMessage:="only one of Before, After can be specified, not both"
End If

If Before Is Nothing And After Is Nothing Then Set After = myLast

If Before Is Nothing And Not After Is Nothing Then
  Set Before = After.nextCommand
ElseIf After Is Nothing And Not Before Is Nothing Then
  Set After = Before.prevCommand
End If

CheckInChain Before, "clsChain.Add", extraMessage:="can't use it as a place to insert before"
CheckInChain After, "clsChain.Add", extraMessage:="can't use it as a place to insert after"

If Not (cmd.inChain Is Nothing) Then
  Throw errAlreadyInChain, extraMessage:="can't add it to another chain"
End If
Set cmd.inChain = Me
If myFirst Is Nothing Then
  'chain is empty. this is the first command.
  Set myFirst = cmd
  Set myLast = cmd
  mySize = 1
Else
  MakeLink cmd, Before
  MakeLink After, cmd
  If cmd.nextCommand Is myFirst Then Set myFirst = cmd
  If cmd.prevCommand Is myLast Then Set myLast = cmd
  mySize = mySize + 1
End If

'autoVerify "Add"
End Sub


Public Sub AddChain(ByVal chain As clsChain, Optional ByVal Before As clsGCommand = Nothing, Optional ByVal After As clsGCommand = Nothing)

If chain.size <= 0 Then Exit Sub

If Not (Before Is Nothing) And Not (After Is Nothing) Then
  Throw errInvalidArgument, "clsChain::Add", extraMessage:="only one of Before, After can be specified, not both"
End If

If Before Is Nothing And After Is Nothing Then Set After = myLast

If Before Is Nothing And Not After Is Nothing Then
  Set Before = After.nextCommand
ElseIf After Is Nothing And Not Before Is Nothing Then
  Set After = Before.nextCommand
End If

CheckInChain Before, "clsChain.AddChain", extraMessage:="can't use it as a place to insert before"
CheckInChain After, "clsChain.AddChain", extraMessage:="can't use it as a place to insert after"

changeTrainChain chain.first, chain.last, Me

If myFirst Is Nothing Then
  'me is empty chain.
  Set myFirst = chain.first
  Set myLast = chain.last
  mySize = chain.size
Else
  MakeLink chain.last, Before
  MakeLink After, chain.first
  If chain.last.nextCommand Is myFirst Then Set myFirst = chain.first
  If chain.first.prevCommand Is myLast Then Set myLast = chain.last
  mySize = mySize + chain.size
End If
chain.unwrapMe

autoVerify "AddChain"
End Sub

Private Sub CheckInChain(ByVal cmd As clsGCommand, source As String, extraMessage As String)
If Not cmd Is Nothing Then
  If Not cmd.inChain Is Me Then
    Throw errNotInChain, source, extraMessage
  End If
End If
End Sub

Private Function changeTrainChain(cmdFirst As clsGCommand, cmdLast As clsGCommand, newChainRef As clsChain) As Long
Dim cnt As Long
Dim cmd As clsGCommand
Set cmd = cmdFirst
Do
  Set cmd.inChain = newChainRef
  cnt = cnt + 1
  If cmd Is cmdLast Then Exit Do
  Set cmd = cmd.nextCommand
Loop
changeTrainChain = cnt
End Function

Public Property Get size() As Long
size = mySize
autoVerify "size"
End Property

'leaves the list floating, used for transferring all items to another list
'don't use from outside
Friend Sub unwrapMe()
Set myFirst = Nothing
Set myLast = Nothing
mySize = 0

autoVerify "unwrapMe"
End Sub

Public Property Get first() As clsGCommand
Set first = myFirst
End Property

Public Property Get last() As clsGCommand
Set last = myLast
End Property

Public Sub delete()
Dim cmd As clsGCommand
If mySize = 0 Then Exit Sub
Set cmd = myFirst
Do
  Set cmd.inChain = Nothing
  If cmd Is myLast Then Exit Do
  Set cmd = cmd.nextCommand
  BreakLink_Prev cmd
Loop

Set myFirst = Nothing
Set myLast = Nothing
mySize = 0

'autoVerify "delete"
End Sub

'tolerant to empty values
Private Sub MakeLink(cmd1 As clsGCommand, cmd2 As clsGCommand)
If Not cmd1 Is Nothing Then
  Set cmd1.nextCommand = cmd2
End If
If Not cmd2 Is Nothing Then
  Set cmd2.prevCommand = cmd1
End If
End Sub

Private Sub BreakLink_Next(cmd As clsGCommand)
If Not cmd.nextCommand Is Nothing Then
  Set cmd.nextCommand.prevCommand = Nothing
  Set cmd.nextCommand = Nothing
End If
End Sub

Private Sub BreakLink_Prev(cmd As clsGCommand)
If Not cmd.prevCommand Is Nothing Then
  Set cmd.prevCommand.nextCommand = Nothing
  Set cmd.prevCommand = Nothing
End If
End Sub

'removes a command from the list. If keepRefs is true, the withdrawn
'command's links will be kept, to assist locating where it was later.
'If removing a series of commands, make sure to purge their references,
'otherwise memory leaks will result
Public Sub withdraw(ByVal cmd As clsGCommand, Optional ByVal keepRefs As Boolean = False)
CheckInChain cmd, "clsChain.withdraw", "can't withdraw"
Dim prv As clsGCommand, nxt As clsGCommand
Set prv = cmd.prevCommand
Set nxt = cmd.nextCommand
If cmd Is myFirst Then Set myFirst = cmd.nextCommand
If cmd Is myLast Then Set myLast = cmd.prevCommand

'break-before-make, otherwise break would corrupt the just-made link
If Not keepRefs Then
  BreakLink_Prev cmd
  BreakLink_Next cmd
End If

Set cmd.inChain = Nothing
mySize = mySize - 1
MakeLink prv, nxt

autoVerify "withdraw"
End Sub

Public Function withdrawChain(ByVal cmdFrom As clsGCommand, _
                              ByVal cmdTo As clsGCommand, _
                              Optional ByVal preserveLinks As Boolean) As clsChain
CheckInChain cmdFrom, "clsChain.withdrawChain", "can't withdraw"
CheckInChain cmdTo, "clsChain.withdrawChain", "can't withdraw"
Dim prv As clsGCommand, nxt As clsGCommand
Set prv = cmdFrom.prevCommand
Set nxt = cmdTo.nextCommand
If cmdFrom Is myFirst Then Set myFirst = cmdTo.nextCommand
If cmdTo Is myLast Then Set myLast = cmdFrom.prevCommand

Dim chain As New clsChain
chain.wrap cmdFrom, cmdTo
If Not preserveLinks Then
  BreakLink_Prev cmdFrom
  BreakLink_Next cmdTo
End If
MakeLink prv, nxt

mySize = mySize - chain.size

Set withdrawChain = chain

'autoVerify "withdrawChain"
End Function

Friend Sub wrap(cmdFirst As clsGCommand, cmdLast As clsGCommand)
If mySize > 0 Then
  Me.delete
End If
Set myFirst = cmdFirst
Set myLast = cmdLast
mySize = changeTrainChain(cmdFirst, cmdLast, Me)

autoVerify "wrap"
End Sub


Public Function verify()
Const MAX_ALLOWED_LENGTH As Long = 30000000
Dim valid As Boolean: valid = True
If mySize = 0 Then
  If Me.first Is Nothing And Me.last Is Nothing Then
    'ok
  Else
    Debug.Print "List size is zero, but it has pointers to first or last"
    valid = False
  End If
  verify = valid
  Exit Function
Else
  If Me.first Is Nothing Or Me.last Is Nothing Then
    Debug.Print "List size is nonzero, but one of first/last pointers is missing"
    valid = False
  End If
End If

Dim cmd As clsGCommand
Dim cmdPrev As clsGCommand
Dim cnt As Long
Set cmdPrev = Nothing
Set cmd = myFirst
Do
  If Not cmdPrev Is Nothing Then
    If cmd.prevCommand Is cmdPrev Then
      'ok
    Else
      Debug.Print "GCommand is not linked to previous command in the chain"
      valid = False
    End If
  End If
  If cmd.inChain Is Me Then
    'ok
  Else
    If cmd.inChain Is Nothing Then
      Debug.Print "GCommand's chain is not set."
    Else
      Debug.Print "GCommand is in another chain."
    End If
    valid = False
  End If
  cnt = cnt + 1
  If cmd Is myLast Then Exit Do
  Set cmdPrev = cmd
  Set cmd = cmd.nextCommand
  If cmd Is Nothing Then
    Debug.Print "Chain's end was reached, but element pointed by last was not encountered. Broken links?"
    valid = False
    verify = valid
    Exit Function
  End If
Loop While cnt < MAX_ALLOWED_LENGTH
If cnt = MAX_ALLOWED_LENGTH Then
  Debug.Print "Chain count is above " + Str(MAX_ALLOWED_LENGTH) + ". The chain is probably looped."
  valid = False
ElseIf cnt <> mySize Then
  Debug.Print "Chain's count property is out of sync."
  valid = False
End If
verify = valid
End Function
