VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBlokada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type typThis
  blockCount As Long
  Master As clsBlokada
End Type
Dim this As typThis

Public Event Blocked()
Public Event Unblocked()

Public Property Get isBlocked() As Boolean
Attribute isBlocked.VB_UserMemId = 0
isBlocked = this.blockCount > 0
End Property

Public Property Get Master() As clsBlokada
Set Master = this.Master
End Property

Friend Property Set Master(ByVal newMaster As clsBlokada)
If Not this.Master Is Nothing Then Throw errInvalidArgument, , extraMessage:="master of blockada can be set only once!"
Set this.Master = newMaster
End Property

'don't call, it is for internal use
Friend Sub Decr(ByVal keeper As clsBlokada)
If keeper.Master Is Me Then
  Debug.Assert this.blockCount > 0
  this.blockCount = this.blockCount - 1
  If this.blockCount = 0 Then RaiseEvent Unblocked
End If
End Sub

'returns a keeper
Public Function block() As clsBlokada
If Not this.Master Is Nothing Then Throw errInvalidArgument, , "block can only be called for a master object, not for keeper"
Dim keeper As clsBlokada
Set keeper = New clsBlokada
Set keeper.Master = Me
this.blockCount = this.blockCount + 1
Set block = keeper
If this.blockCount = 1 Then RaiseEvent Blocked
End Function

Private Sub Class_Terminate()
If Not this.Master Is Nothing Then
  this.Master.Decr Me
  Set this.Master = Nothing
End If
End Sub

Public Sub Unblock()
If this.Master Is Nothing Then Throw errInvalidArgument, , "unblock can only be called for a keeper object, not for master"
If Not this.Master Is Nothing Then
  this.Master.Decr Me
  Set this.Master = Nothing
End If
End Sub
