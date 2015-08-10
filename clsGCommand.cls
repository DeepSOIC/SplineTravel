VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public strLine As String

Private stateBefore As typCurrentState
Private stateAfter As typCurrentState
Public touched As Boolean
Private myCmdType As eGCommand
Private myMove As clsGMove

Private refPrevCommand As clsGCommand
Private refNextCommand As clsGCommand
Public inChain As clsChain 'dumb member, to be managed by the chain itself

Public Enum eGCommand
'low byte = letter:
'G = &H47
'M = &H4D
'high byte = number after letter
  egcEmpty = 0
  G0_quickMove = &H47&
  G1_controlledMove = &H147&
  G4_dwell = &H447&
  G21_setUnitMm = &H1547&
  M82_absoluteE = &H524D&
  M83_relativeE = &H534D&
  G90_absolutePos = &H5A47&
  G91_relativePos = &H5B47&
  G92_overridePos = &H5C47&
End Enum

Private Type typLetVal 'letter + value, argument of g command
  'letter as byte = index of array
  Value As Double
  specified As Boolean
End Type

Dim arguments(65 To 90) As typLetVal '65 to 90 = A to Z

'recompute states: update stateBefore from previous command;
'reparse arguments of a command and compute the new state.
'if preserveDeltaE is specified, current extrusion length as
'derived from current difference in states, is preserved
'(unless the command is G92)
Public Sub RecomputeStates(Optional ByVal preserveDeltaE = False, Optional ByVal keepStateBefore As Boolean)

'save current deltaE, we'll need it if preserveDeltaE = true
Dim oldDeltaE As Double
oldDeltaE = stateAfter.Epos - stateBefore.Epos

If Not (refPrevCommand Is Nothing) And Not keepStateBefore Then
  stateBefore = refPrevCommand.CompleteStateAfter
End If

'reset stateAfter to match stateBefore, then apply changes.
stateAfter = stateBefore

'these values are handled in a command-independent way, assume they are global-persistent
Dim arg As typLetVal
arg = arguments(Asc("X"))
If arg.specified Then
  If stateBefore.MoveRelative Then
    stateAfter.pos.X = stateAfter.pos.X + arg.Value
  Else
    stateAfter.pos.X = arg.Value
  End If
End If
arg = arguments(Asc("Y"))
If arg.specified Then
  If stateBefore.MoveRelative Then
    stateAfter.pos.Y = stateAfter.pos.Y + arg.Value
  Else
    stateAfter.pos.Y = arg.Value
  End If
End If
arg = arguments(Asc("Z"))
If arg.specified Then
  If stateBefore.MoveRelative Then
    stateAfter.pos.Z = stateAfter.pos.Z + arg.Value
  Else
    stateAfter.pos.Z = arg.Value
  End If
End If

arg = arguments(Asc("E"))
If arg.specified Then
  If stateBefore.ExtrusionRelative Then
    stateAfter.Epos = stateBefore.Epos + arg.Value
  Else
    stateAfter.Epos = arg.Value
  End If
  If preserveDeltaE Then
    stateAfter.Epos = stateBefore.Epos + oldDeltaE
    If Not stateBefore.ExtrusionRelative Then
      'update E argument
      arg.Value = vtStr(Round(stateAfter.Epos, mdlPrecision.extrDecimals))
    End If
  End If
End If

arg = arguments(Asc("F"))
If arg.specified Then
  stateAfter.Speed = arg.Value / 60 'convert to mm/s
End If

Select Case myCmdType
  Case eGCommand.G90_absolutePos
    stateAfter.MoveRelative = False
  Case eGCommand.G91_relativePos
    stateAfter.MoveRelative = True
  Case eGCommand.M82_absoluteE
    stateAfter.ExtrusionRelative = False
  Case eGCommand.M83_relativeE
    stateAfter.ExtrusionRelative = True
  Case eGCommand.G92_overridePos
    If arguments(Asc("E")).specified Then
      stateAfter.Epos = arguments(Asc("E")).Value
    End If
    If arguments(Asc("X")).specified Then
      stateAfter.pos.X = arguments(Asc("X")).Value
    End If
    If arguments(Asc("Y")).specified Then
      stateAfter.pos.Y = arguments(Asc("Y")).Value
    End If
    If arguments(Asc("Z")).specified Then
      stateAfter.pos.Z = arguments(Asc("Z")).Value
    End If
End Select
End Sub

Public Sub regenerateString()
Me.strLine = ""
If Me.isEmpty Then Exit Sub

Me.strLine = Me.strLine + Chr$(myCmdType And &HFF&) + vtStr((myCmdType And &HFF00&) / &H100)

Dim i As Long
For i = LBound(arguments) To UBound(arguments)
  If arguments(i).specified Then
    Me.strLine = Me.strLine + " " + Chr$(i) + vtStr(arguments(i).Value)
  End If
Next i
End Sub

Friend Property Get CompleteStateAfter() As typCurrentState
CompleteStateAfter = stateAfter
End Property

Friend Property Get CompleteStateBefore() As typCurrentState
CompleteStateBefore = stateBefore
End Property

Public Function getEnterSpeed() As clsVector3D
'FIXME: support arcs
Set getEnterSpeed = _
  makeClsVector(stateAfter.pos) _
  .Subtrd(makeClsVector(stateBefore.pos)) _
  .Normalized.Multed(stateAfter.Speed)
End Function

Public Function getExitSpeed() As clsVector3D
'FIXME: support arcs
Set getExitSpeed = getEnterSpeed
End Function

Public Sub ParseString(Optional ByVal throwIfInvalid = False)
Erase arguments
myCmdType = 0

On Error GoTo eh

  Dim lineText As String
  lineText = strLine
  
  'get rid of comment
  Dim pos As Long
  pos = InStr(lineText, ";")
  If pos > 0 Then
    lineText = Trim$(Left$(lineText, pos - 1))
  End If
  'get rid of checksum
  pos = InStrRev(lineText, "*")
  If pos > 0 Then
    lineText = Trim$(Left$(lineText, pos - 1))
  End If
  
  If Len(lineText) = 0 Then Exit Sub
  Dim words() As String
  words = VBA.split(lineText, " ")
  Dim iword As Long
  For iword = 0 To UBound(words)
    If Len(words(iword)) > 0 Then 'this check is needed to deal with double spaces
      Dim firstletter As Long
      firstletter = Asc(UCase$(Mid$(words(iword), 1, 1)))
      If firstletter < Asc("A") Or firstletter > Asc("Z") Then
        Throw errInvalidCommand, extraMessage:="first character of a word in a command must be a latin letter; it is """ + Chr$(firstletter) + """"
      End If
      Dim Value As Double
      Value = val(Mid$(words(iword), 2))
      
      If iword = 0 Then
        myCmdType = CLng(Value) * &H100 + firstletter
      Else
        arguments(firstletter).specified = True
        arguments(firstletter).Value = Value
      End If
    End If
  Next iword
  processArguments
Exit Sub
eh:
If throwIfInvalid Then
  Throw
Else
  Debug.Print "Error parsing G-code line. Line = """ + strLine + """"
  Debug.Print "Error is: " + Err.Description
End If

End Sub

Public Function isEmpty()
isEmpty = (myCmdType = egcEmpty)
End Function

'the command involves moving of head; the move can be constructed
'note that this will be true for speed setting command; checking for this is a todo
Public Function isMove()
isMove = (myCmdType = G0_quickMove) Or (myCmdType = G1_controlledMove)
End Function

Public Sub processArguments()
Select Case myCmdType
  Case eGCommand.G0_quickMove, eGCommand.G1_controlledMove
    constructMove
End Select
End Sub

Public Sub constructMove()
If Not isMove Then
  Set myMove = Nothing
  Exit Sub
End If

If myMove Is Nothing Then
  Set myMove = New clsGMove
End If
myMove.p1.copyFromT stateBefore.pos
myMove.p2.copyFromT stateAfter.pos
myMove.Extrusion = stateAfter.Epos - stateBefore.Epos

If myMove.traveldist > 1E-100 Then
  myMove.Speed = stateAfter.Speed
ElseIf Abs(myMove.Extrusion) > 1E-100 Then
  myMove.ExtrusionSpeed = stateAfter.Speed
Else
  'move is invalid, it is probably a set speed command
  Set myMove = Nothing
End If
End Sub

'make sure to have called constructMove!
Public Function getMove() As clsGMove
Set getMove = myMove
End Function

'sets the move and updates the command
'EError is rounding error tracker variable for extrusion values.
'note: doesn't update states automatically!
Public Sub setMove(ByVal newMove As clsGMove, Optional ByRef EError As Double)
If Not Me.isMove And Not Me.isEmpty Then Throw errWrongType, "clsGCommand.setMove", "command """ + Me.strLine + """ is not a move command"
Me.strLine = newMove.GenerateGCode(stateBefore, EError)
Me.ParseString
End Sub

Public Property Get prevCommand() As clsGCommand
Set prevCommand = refPrevCommand
End Property

Public Property Set prevCommand(newVal As clsGCommand)
Set refPrevCommand = newVal
End Property


Public Property Get nextCommand() As clsGCommand
Set nextCommand = refNextCommand
End Property

Public Property Set nextCommand(newVal As clsGCommand)
Set refNextCommand = newVal
End Property

'
'Public Sub DiscardChainBeforeMe()
'Dim cmd As clsGCommand
''disassemble chain preceding this command
'Set cmd = Me.prevCommand
'Do Until cmd Is Nothing
'  Set cmd.nextCommand.prevCommand = Nothing
'  Set cmd.nextCommand = Nothing
'  Set cmd = cmd.prevCommand
'Loop Until cmd Is Nothing
'End Sub
'
'Public Sub DiscardChainAfterMe()
'Dim cmd As clsGCommand
''disassemble chain preceding this command
'Set cmd = Me.nextCommand
'Do Until cmd Is Nothing
'  Set cmd.prevCommand.nextCommand = Nothing
'  Set cmd.prevCommand = Nothing
'  Set cmd = cmd.nextCommand
'Loop
'End Sub

'also true for hop moves
Public Property Get isTravelMove() As Boolean
If Not isMove Then
  isTravelMove = False
  Exit Property
End If
Dim traveldist As Double
traveldist = Vector3D.Dist(stateBefore.pos, stateAfter.pos)
If traveldist < 1E-100 Then
  isTravelMove = False
  Exit Function
End If
If Abs(stateAfter.Epos - stateBefore.Epos) > 1E-100 Then
  isTravelMove = False
  Exit Function
End If
isTravelMove = True
End Property

Public Property Get isExtruderMove() As Boolean
If Not isMove Then
  isExtruderMove = False
  Exit Property
End If
Dim traveldist As Double
traveldist = Vector3D.Dist(stateBefore.pos, stateAfter.pos)
If traveldist > 1E-100 Then
  isExtruderMove = False
  Exit Property
End If
If Abs(stateAfter.Epos - stateBefore.Epos) < 1E-100 Then
  isExtruderMove = False
  Exit Property
End If
isExtruderMove = True
End Property

'note: if there is a position change too, this is still a retract.
'If you want to filter that out, check for isExtruderMove
Public Property Get isRetract() As Boolean
If Not isMove Then 'FIXME: support retract G-command
  isRetract = False
  Exit Property
End If
If stateAfter.Epos - stateBefore.Epos > -1E-100 Then
  isRetract = False
  Exit Property
End If
isRetract = True
End Property

'returns true if the move includes positive extusion and some motion.
Public Property Get isBuildMove() As Boolean
If Not isMove Then
  isBuildMove = False
  Exit Property
End If
Dim traveldist As Double
traveldist = Vector3D.Dist(stateBefore.pos, stateAfter.pos)
If traveldist < 1E-100 Then
  isBuildMove = False
  Exit Property
End If
If stateAfter.Epos - stateBefore.Epos < 1E-100 Then
  isBuildMove = False
  Exit Property
End If
isBuildMove = True
End Property

Public Property Get Echange() As Double
Echange = stateAfter.Epos - stateBefore.Epos
End Property

Public Property Get PosChange() As clsVector3D
Set PosChange = makeClsVector(stateAfter.pos).Subtrd(makeClsVector(stateBefore.pos))
End Property

Public Property Get ZChange() As Double
ZChange = stateAfter.pos.Z - stateBefore.pos.Z
End Property

'how much time it takes to execute the command?
'(only valid for move commands; make sure to have called constructMove before using)
'Returns the time assumming the move is done linearly with current feedrate.
Public Property Get execTime() As Double
If Not myMove Is Nothing Then
  execTime = myMove.time
End If
End Property

'make sure to have called constructMove
Public Sub split(ByVal timePoint As Double, Optional ByRef EError As Double)
Dim Part1 As clsGMove, Part2 As clsGMove
If myMove.split(timePoint, Part1, Part2) Then
  Dim cmd2 As New clsGCommand
  Me.setMove Part1, EError
  Me.RecomputeStates keepStateBefore:=True
  cmd2.strLine = Part2.GenerateGCode(Me.CompleteStateAfter, EError)
  cmd2.ParseString
  Me.inChain.Add cmd2, After:=Me
Else
  Throw errInvalidArgument, extraMessage:="timePoint is not within 0..execTime, can't split"
End If
End Sub

Public Function getPrevMove() As clsGCommand
Dim cmd As clsGCommand
Set cmd = Me
Do
  If cmd Is Me.inChain.first Then Set cmd = Nothing: Exit Do
  Set cmd = cmd.prevCommand
Loop Until cmd.isMove
Set getPrevMove = cmd
End Function

Public Function getNextMove() As clsGCommand
Dim cmd As clsGCommand
Set cmd = Me
Do
  If cmd Is Me.inChain.last Then Set cmd = Nothing: Exit Do
  Set cmd = cmd.nextCommand
Loop Until cmd.isMove
Set getNextMove = cmd
End Function

