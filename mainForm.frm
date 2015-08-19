VERSION 5.00
Begin VB.Form mainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SplineTravel"
   ClientHeight    =   8180
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   11990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8180
   ScaleWidth      =   11990
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Travel moves and retraction"
      Height          =   5750
      Left            =   50
      TabIndex        =   21
      Top             =   900
      Width           =   7800
      Begin VB.TextBox txtRetract 
         Height          =   370
         Left            =   5270
         TabIndex        =   45
         Text            =   "1.5"
         Top             =   270
         Width           =   1450
      End
      Begin VB.OptionButton optTravelStraight 
         BackColor       =   &H0080FFFF&
         Caption         =   "Straight travel"
         Height          =   300
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   4210
         Width           =   3210
      End
      Begin VB.OptionButton optTravelSpline 
         BackColor       =   &H0080FFFF&
         Caption         =   "Spline travel"
         Height          =   300
         Left            =   280
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   660
         Value           =   -1  'True
         Width           =   3210
      End
      Begin VB.Frame Frame4 
         Height          =   1380
         Left            =   180
         TabIndex        =   42
         Top             =   4220
         Width           =   7480
         Begin VB.TextBox txtZHop 
            Height          =   370
            Left            =   1160
            TabIndex        =   54
            Text            =   "1"
            Top             =   860
            Width           =   1450
         End
         Begin VB.TextBox txtRSpeedStraight 
            Height          =   370
            Left            =   5040
            TabIndex        =   51
            Text            =   "300"
            Top             =   370
            Width           =   1450
         End
         Begin VB.TextBox txtSpeedStraight 
            Height          =   370
            Left            =   1160
            TabIndex        =   48
            Text            =   "200"
            Top             =   390
            Width           =   1450
         End
         Begin VB.Label Label31 
            Caption         =   "mm"
            Height          =   240
            Left            =   2710
            TabIndex        =   56
            Top             =   930
            Width           =   920
         End
         Begin VB.Label Label30 
            Caption         =   "Z-hop"
            Height          =   340
            Left            =   160
            TabIndex        =   55
            Top             =   920
            Width           =   1210
         End
         Begin VB.Label Label29 
            Caption         =   "mm/s"
            Height          =   240
            Left            =   6540
            TabIndex        =   53
            Top             =   440
            Width           =   630
         End
         Begin VB.Label Label28 
            Caption         =   "retract speed"
            Height          =   340
            Left            =   3940
            TabIndex        =   52
            Top             =   430
            Width           =   1020
         End
         Begin VB.Label Label27 
            Caption         =   "mm/s"
            Height          =   240
            Left            =   2710
            TabIndex        =   50
            Top             =   460
            Width           =   920
         End
         Begin VB.Label Label26 
            Caption         =   "speed"
            Height          =   340
            Left            =   160
            TabIndex        =   49
            Top             =   450
            Width           =   1210
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3500
         Left            =   180
         TabIndex        =   22
         Top             =   700
         Width           =   7460
         Begin VB.TextBox txtEAccel 
            Height          =   370
            Left            =   5070
            TabIndex        =   28
            Text            =   "1000"
            Top             =   1580
            Width           =   1450
         End
         Begin VB.TextBox txtEJerk 
            Height          =   370
            Left            =   5080
            TabIndex        =   27
            Text            =   "8"
            Top             =   2220
            Width           =   1450
         End
         Begin VB.TextBox txtZJerk 
            Height          =   370
            Left            =   1120
            TabIndex        =   26
            Text            =   "0"
            Top             =   2860
            Width           =   1450
         End
         Begin VB.TextBox txtSpeedLimit 
            Height          =   370
            Left            =   1120
            TabIndex        =   25
            Text            =   "200"
            Top             =   990
            Width           =   1450
         End
         Begin VB.TextBox txtCurveJerk 
            Height          =   370
            Left            =   1120
            TabIndex        =   24
            Text            =   "2"
            Top             =   2200
            Width           =   1450
         End
         Begin VB.TextBox txtAcceleration 
            Height          =   370
            Left            =   1120
            TabIndex        =   23
            Text            =   "800"
            Top             =   1610
            Width           =   1450
         End
         Begin VB.Label Label16 
            Caption         =   "E acceleration"
            Height          =   610
            Left            =   3890
            TabIndex        =   41
            Top             =   1610
            Width           =   1070
         End
         Begin VB.Label Label15 
            Caption         =   "mm/s2"
            Height          =   240
            Left            =   6620
            TabIndex        =   40
            Top             =   1650
            Width           =   760
         End
         Begin VB.Label Label12 
            Caption         =   "E jerk (for retraction)"
            Height          =   610
            Left            =   3910
            TabIndex        =   39
            Top             =   2180
            Width           =   1070
         End
         Begin VB.Label Label11 
            Caption         =   "mm/s"
            Height          =   240
            Left            =   6630
            TabIndex        =   38
            Top             =   2290
            Width           =   760
         End
         Begin VB.Label Label10 
            Caption         =   "Z jerk (for hopping)"
            Height          =   610
            Left            =   70
            TabIndex        =   37
            Top             =   2870
            Width           =   1070
         End
         Begin VB.Label Label9 
            Caption         =   "mm/s"
            Height          =   240
            Left            =   2670
            TabIndex        =   36
            Top             =   2930
            Width           =   760
         End
         Begin VB.Label label8 
            Caption         =   "mm/s"
            Height          =   240
            Left            =   2670
            TabIndex        =   35
            Top             =   1060
            Width           =   920
         End
         Begin VB.Label Label7 
            Caption         =   "speed limit"
            Height          =   340
            Left            =   120
            TabIndex        =   34
            Top             =   1050
            Width           =   1210
         End
         Begin VB.Label Label6 
            Caption         =   "mm/s"
            Height          =   240
            Left            =   2670
            TabIndex        =   33
            Top             =   2270
            Width           =   760
         End
         Begin VB.Label Label5 
            Caption         =   "curve tesellation (jerk)"
            Height          =   610
            Left            =   50
            TabIndex        =   32
            Top             =   2060
            Width           =   1070
         End
         Begin VB.Label Label4 
            Caption         =   "mm/s2"
            Height          =   240
            Left            =   2670
            TabIndex        =   31
            Top             =   1680
            Width           =   920
         End
         Begin VB.Label Label3 
            Caption         =   "acceleration"
            Height          =   340
            Left            =   30
            TabIndex        =   30
            Top             =   1630
            Width           =   1210
         End
         Begin VB.Label Label20 
            Caption         =   $"mainForm.frx":0000
            Height          =   700
            Left            =   170
            TabIndex        =   29
            Top             =   300
            Width           =   7060
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label Label14 
         Caption         =   "retract"
         Height          =   610
         Left            =   4170
         TabIndex        =   47
         Top             =   310
         Width           =   1070
      End
      Begin VB.Label Label13 
         Caption         =   "mm"
         Height          =   240
         Left            =   6820
         TabIndex        =   46
         Top             =   340
         Width           =   760
      End
   End
   Begin VB.TextBox txtNotes 
      Height          =   700
      Left            =   7310
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   140
      Width           =   4220
   End
   Begin VB.Frame Frame3 
      Height          =   2460
      Left            =   8040
      TabIndex        =   9
      Top             =   1000
      Width           =   3830
      Begin VB.TextBox txtRSpeedSC 
         Height          =   370
         Left            =   1280
         TabIndex        =   17
         Text            =   "8"
         Top             =   1780
         Width           =   1450
      End
      Begin VB.TextBox txtLoopTol 
         Height          =   370
         Left            =   1330
         TabIndex        =   11
         Text            =   "0.3"
         Top             =   1250
         Width           =   1450
      End
      Begin VB.CheckBox chkSeamConceal 
         BackColor       =   &H00FF80FF&
         Caption         =   "seam concealment"
         Height          =   310
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Value           =   1  'Checked
         Width           =   1790
      End
      Begin VB.Label Label23 
         Caption         =   "speed of concealed retraction"
         Height          =   570
         Left            =   230
         TabIndex        =   19
         Top             =   1650
         Width           =   1070
      End
      Begin VB.Label Label22 
         Caption         =   "mm/s"
         Height          =   240
         Left            =   2850
         TabIndex        =   18
         Top             =   1850
         Width           =   760
      End
      Begin VB.Label Label19 
         Caption         =   $"mainForm.frx":00D0
         Height          =   630
         Left            =   110
         TabIndex        =   14
         Top             =   350
         Width           =   3600
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         Caption         =   "mm"
         Height          =   240
         Left            =   2900
         TabIndex        =   13
         Top             =   1320
         Width           =   760
      End
      Begin VB.Label Label17 
         Caption         =   "loop detection tolerance"
         Height          =   610
         Left            =   230
         TabIndex        =   12
         Top             =   1200
         Width           =   1070
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Presets"
      Height          =   730
      Left            =   20
      TabIndex        =   5
      Top             =   40
      Width           =   6290
      Begin VB.CommandButton cmdResetSettings 
         BackColor       =   &H008080FF&
         Caption         =   "*"
         Height          =   280
         Left            =   5890
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "load defaults"
         Top             =   390
         Visible         =   0   'False
         Width           =   340
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H008080FF&
         Caption         =   "Delete"
         Height          =   280
         Left            =   5450
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   270
         Width           =   720
      End
      Begin VB.ComboBox cmbPreset 
         Height          =   280
         ItemData        =   "mainForm.frx":015B
         Left            =   100
         List            =   "mainForm.frx":015D
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   270
         Width           =   4110
      End
      Begin VB.CommandButton cmdSaveAs 
         Caption         =   "Save as..."
         Height          =   280
         Left            =   4330
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   270
         Width           =   1060
      End
   End
   Begin VB.CommandButton cmdProcessFile 
      BackColor       =   &H0080FF80&
      Caption         =   "Go"
      CausesValidation=   0   'False
      Height          =   840
      Left            =   8380
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7000
      Width           =   3520
   End
   Begin VB.TextBox txtFNOut 
      Height          =   410
      Left            =   1500
      TabIndex        =   3
      Tag             =   "!f"
      Top             =   7470
      Width           =   6750
   End
   Begin VB.TextBox txtFNIn 
      Height          =   410
      Left            =   1500
      TabIndex        =   0
      Tag             =   "!f"
      Top             =   6990
      Width           =   6730
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "notes"
      Height          =   230
      Left            =   6270
      TabIndex        =   15
      Top             =   330
      Width           =   940
   End
   Begin VB.Label Label2 
      Caption         =   "output"
      Height          =   380
      Left            =   490
      TabIndex        =   2
      Top             =   7500
      Width           =   1090
   End
   Begin VB.Label Label1 
      Caption         =   "input"
      Height          =   270
      Left            =   490
      TabIndex        =   1
      Top             =   7000
      Width           =   1020
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typPresetManagerGlobals
  filesInCombo() As String 'full paths
  block As New clsBlokada  'locks response to events, for when filling preset list
  curPresetIsModified As Boolean
  curPresetFN As String
End Type
Dim pm As typPresetManagerGlobals

Private Sub chkSeamConceal_Click()
ChangeWasMade
End Sub

Private Sub cmbPreset_Click()
If pm.block Then Exit Sub
On Error GoTo eh
If Me.getSelectedPreset = pm.curPresetFN Then Exit Sub  'no change
If pm.curPresetIsModified Then
  Dim answ As VbMsgBoxResult
  answ = MsgBox("Current preset was modified. The changes will be lost. Continue anyway?", vbYesNo Or vbDefaultButton2)
  If answ = vbNo Then
    Me.SelectPreset pm.curPresetFN
    Exit Sub
  End If
End If
Me.LoadPreset Me.getSelectedPreset
Exit Sub
eh:
MsgError
Me.SelectPreset pm.curPresetFN
End Sub

Private Sub cmdDelete_Click()
On Error GoTo eh
If Len(pm.curPresetFN) = 0 Then Throw errInvalidArgument, , "No preset is selected, can't delete"
Dim answ As VbMsgBoxResult
answ = MsgBox("You are about to delete preset " + getFileTitle(pm.curPresetFN) + ", stored in " + pm.curPresetFN + ". Continue?", vbYesNo Or vbDefaultButton2)
If answ = vbNo Then Throw errCancel
Kill pm.curPresetFN
MsgBox "Preset " + getFileTitle(pm.curPresetFN) + " was deleted."
Me.purgeModified
pm.curPresetFN = ""
Me.RefillPresets
Exit Sub
eh:
MsgError
End Sub

Private Sub cmdProcessFile_Click()
cmdProcessFile.Enabled = False
On Error GoTo eh
SavePreset "(last used)"

mdlWorker.Process Me.txtFNIn.Text, Me.txtFNOut.Text, Me

Me.cmdProcessFile.Caption = "Done."
Me.cmdProcessFile.Enabled = True
Exit Sub
eh:
  Me.cmdProcessFile.Caption = "Failed." + vbNewLine + Err.Description
  Me.cmdProcessFile.Enabled = True
End Sub


Private Sub cmdResetSettings_Click()
Me.ResetSettings includeFilenames:=False
End Sub

Private Sub cmdSaveAs_Click()
On Error GoTo eh
Dim presetName As String
presetName = getFileTitle(pm.curPresetFN)
Dim newPresetName As String
again:
newPresetName = InputBox("Name the preset. Entering a new name will keep current preset intact.", presetName)
If Len(newPresetName) = 0 Then Exit Sub
verifyPresetName newPresetName
Dim newPresetFilename As String
If StrComp(newPresetName, presetName, vbTextCompare) = 0 Then
  newPresetFilename = pm.curPresetFN
Else
  Dim paths() As String
  paths = mdlFiles.PresetsPaths
  newPresetFilename = paths(UBound(paths)) + newPresetName + ".ini"
  If Me.FindInCombo(newPresetFilename) <> -1 Then
    If MsgBox("Overwriting another existing preset. Continue?", vbYesNo Or vbDefaultButton2) = vbNo Then
      Exit Sub
    End If
  End If
End If
WritePresetFile newPresetFilename, Me.GetConfigString
pm.curPresetFN = newPresetFilename
pm.curPresetIsModified = False
Me.RefillPresets
Exit Sub
eh:
MsgError

End Sub

Public Sub verifyPresetName(ByRef newName As String)
Dim i As Long
Dim ch As String
Const allowedCharacters As String = " ,.-+=!()'1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
newName = Trim(newName)
If Len(newName) = 0 Then Throw errInvalidArgument, extraMessage:="preset name is empty"
For i = 1 To Len(newName)
  ch = Mid$(newName, i, 1)
  If InStr(allowedCharacters, ch) = 0 Then
    Throw errInvalidArgument, extraMessage:="preset name contains an invalid character " + ch
  End If
Next i
End Sub

Private Sub Form_Activate()
Me.RefillPresets
End Sub

Private Sub Form_Load()
mdlPrecision.InitModule
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = QueryUnloadConstants.vbFormCode Then
  'a temporary form was created and is being destroyed, nothing special needs to be done
Else
  SavePreset "(since last close)"
  'explicitly terminate the application. This will prevent the app from keeping running if it is closed while processing a file
  End
End If
End Sub

'a shortcut. It doesn't update current preset, it merely writes a new preset file
Public Sub SavePreset(presetName As String)
Dim newPresetFilename As String
Dim paths() As String
paths = mdlFiles.PresetsPaths
newPresetFilename = paths(UBound(paths)) + presetName + ".ini"
WritePresetFile newPresetFilename, Me.GetConfigString
End Sub







Public Function GetConfigString(Optional ByVal includeFilenames As Boolean) As String
Dim ctr As Control
Dim configStr As New StringAccumulator
For Each ctr In Me
  Dim strLine As String
  strLine = ""
  
  If ctr.Tag = "!f" And Not includeFilenames Then GoTo continue
  
  If TypeOf ctr Is TextBox Then
    Dim txt As TextBox
    Set txt = ctr
    strLine = ctr.Name + ".Text" + " = " + EscapeString(txt.Text)
  ElseIf TypeOf ctr Is CheckBox Then
    Dim chk As CheckBox
    Set chk = ctr
    strLine = ctr.Name + ".Value" + " = " + Trim(Str(chk.Value))
  ElseIf TypeOf ctr Is OptionButton Then
    Dim opt As OptionButton
    Set opt = ctr
    strLine = ctr.Name + ".Value" + " = " + CStr(opt.Value)
  End If
  
  If Len(strLine) > 0 Then
    configStr.Append strLine + vbNewLine
  End If
continue:
Next
GetConfigString = configStr.content
End Function

'fills the settings from the config string
Public Sub ApplyConfigStr(configStr As String, Optional ByVal suppressErrorMessages As Boolean = False, Optional ByVal includeFilenames As Boolean)
On Error GoTo eh
Dim lines() As String
lines = split(configStr, vbNewLine)
Dim i As Long
For i = 0 To UBound(lines)
  Dim strLine As String
  strLine = lines(i)
  If Not Len(strLine) > 0 Then GoTo continue
  
  Dim hs() As String
  hs = split(strLine, "=", limit:=2)
  If UBound(hs) <> 1 Then Throw errWrongConfigLine, "ApplyConfigStr", extraMessage:="Failed to split left-hand-side and right-hand-side in config line number " + Str(i)
  hs(0) = Trim(hs(0))
  'l-trim the value string, but only one space that is attached to " = "
  If Len(hs(1)) > 0 Then
    If Mid$(hs(1), 1, 1) = " " Then
      hs(1) = Mid$(hs(1), 2)
    End If
  End If
  
  Dim objprop() As String 'split object name and property name
  objprop = split(hs(0), ".", limit:=2)
  If UBound(objprop) <> 1 Then Throw errWrongConfigLine, "ApplyConfigStr", extraMessage:="Failed to split object and property in config line number " + Str(i)
  
  Dim objName As String: objName = Trim(objprop(0))
  Dim propName As String: propName = Trim(objprop(1))
  Dim obj As Control:  Set obj = Nothing
  Set obj = CallByName(Me, objName, VbGet)
  If Not obj Is Nothing Then
    If obj.Tag = "!f" And Not includeFilenames Then GoTo continue
    CallByName obj, propName, VbLet, unEscapeString(hs(1))
  End If
continue:
Next i

Exit Sub
eh:
  Dim answ As VbMsgBoxResult
  If Not suppressErrorMessages Then
    answ = MsgError(Style:=vbAbortRetryIgnore)
  Else
    answ = vbIgnore
    Debug.Print Err.Source, Err.Description
  End If
  
  If answ = vbIgnore Then
    Resume continue
  ElseIf answ = vbRetry Then
    Resume
  Else
    Throw errCancel
  End If
End Sub


Public Sub ResetSettings(Optional ByVal includeFilenames As Boolean = False)
Dim tmpForm As mainForm
Set tmpForm = New mainForm
Load tmpForm
On Error GoTo cleanup
Me.ApplyConfigStr tmpForm.GetConfigString(includeFilenames:=includeFilenames), includeFilenames:=includeFilenames
Unload tmpForm
Exit Sub
cleanup:
PushError
Unload tmpForm
PopError
Throw
End Sub

Public Sub RefillPresets()
Dim keeper As clsBlokada: Set keeper = pm.block.block

pm.filesInCombo = mdlFiles.getListOfFiles(mdlFiles.PresetsPaths, "*.ini")
cmbPreset.Clear
Dim i As Long
For i = 0 To ArrLen(pm.filesInCombo) - 1
  cmbPreset.AddItem mdlFiles.getFileTitle(pm.filesInCombo(i))
  cmbPreset.ItemData(cmbPreset.NewIndex) = i
  If pm.curPresetFN = pm.filesInCombo(i) And pm.curPresetIsModified Then
    cmbPreset.List(cmbPreset.NewIndex) = cmbPreset.List(cmbPreset.NewIndex) + "*"
  End If
Next i

Me.SelectPreset pm.curPresetFN

keeper.Unblock
End Sub

'returns -1 if item not found
Public Function FindInCombo(presetFilePath As String) As Long
FindInCombo = -1
Debug.Assert cmbPreset.ListCount = ArrLen(pm.filesInCombo)
Dim i As Long
For i = 0 To cmbPreset.ListCount - 1
  If StrComp(pm.filesInCombo(cmbPreset.ItemData(i)), presetFilePath, vbTextCompare) = 0 Then
    FindInCombo = i
  End If
Next i
End Function

Private Function ArrLen(arr As Variant) As Long
Dim ub As Long: ub = -1
On Error Resume Next
ub = UBound(arr)
ArrLen = ub + 1
End Function

Public Sub LoadPreset(FilePath As String)
Dim tmp As String
tmp = ReadPresetFile(FilePath)
Dim keeper As clsBlokada: Set keeper = pm.block.block
Me.ResetSettings
Me.ApplyConfigStr tmp
Me.purgeModified
Me.SelectPreset FilePath
pm.curPresetFN = FilePath
keeper.Unblock
End Sub

Public Function ReadPresetFile(FilePath As String) As String
Dim f As Long: f = FreeFile
Open FilePath For Input As f
On Error GoTo cleanup
  ReadPresetFile = Input$(LOF(f), #(f))
Close f
Exit Function
cleanup:
  PushError
  Close f
  PopError
  Throw
End Function

Public Sub WritePresetFile(FilePath As String, content As String)
Dim f As Long: f = FreeFile
Open FilePath For Output As f
On Error GoTo cleanup
  Print #(f), content;
Close f
Exit Sub
cleanup:
  PushError
  Close f
  PopError
  Throw
End Sub

Public Function getSelectedPreset(Optional ByVal index As Long = -2) As String
If index = -2 Then index = cmbPreset.ListIndex
If index = -1 Then getSelectedPreset = "": Exit Function
Debug.Assert ArrLen(pm.filesInCombo) = cmbPreset.ListCount
getSelectedPreset = pm.filesInCombo(cmbPreset.ItemData(index))
End Function

'only selects; no file is opened
Public Function SelectPreset(presetFilePath As String) As Long
Dim idx As Long
idx = Me.FindInCombo(presetFilePath)
If Len(presetFilePath) > 0 And idx = -1 Then
  'preset isn't in the list, add it...
  Me.addPresetToList presetFilePath
  idx = Me.FindInCombo(presetFilePath)
  Debug.Assert idx <> -1
End If
Dim keeper As clsBlokada: Set keeper = pm.block.block
cmbPreset.ListIndex = idx
keeper.Unblock
End Function

Public Function addPresetToList(presetFilePath As String, Optional ByVal checkIfExists As Boolean = False) As Long
If checkIfExists Then
  If Me.FindInCombo(presetFilePath) <> -1 Then
    'already in list!
    Exit Function
  End If
End If
If ArrLen(pm.filesInCombo) = 0 Then ReDim pm.filesInCombo(0 To 0) Else ReDim Preserve pm.filesInCombo(0 To UBound(pm.filesInCombo) + 1)
Dim i As Long: i = UBound(pm.filesInCombo)
pm.filesInCombo(i) = presetFilePath
cmbPreset.AddItem mdlFiles.getFileTitle(pm.filesInCombo(i))
cmbPreset.ItemData(cmbPreset.NewIndex) = i
End Function

Public Sub ChangeWasMade()
If pm.block Then Exit Sub
If cmbPreset.ListIndex = -1 Then Exit Sub
pm.curPresetIsModified = True
Dim txt As String
txt = cmbPreset.List(cmbPreset.ListIndex)
If right$(txt, 1) <> "*" Then
  txt = txt + "*"
  cmbPreset.List(cmbPreset.ListIndex) = txt
End If

End Sub

Private Sub optTravelSpline_Click()
ChangeWasMade
End Sub

Private Sub optTravelStraight_Click()
ChangeWasMade
End Sub

Private Sub txtAcceleration_Change()
ChangeWasMade
End Sub

Private Sub txtCurveJerk_Change()
ChangeWasMade
End Sub

Private Sub txtEAccel_Change()
ChangeWasMade
End Sub

Private Sub txtEJerk_Change()
ChangeWasMade
End Sub

Private Sub txtLoopTol_Change()
ChangeWasMade
End Sub

Private Sub txtRetract_Change()
ChangeWasMade
End Sub

Private Sub txtRSpeedSC_Change()
ChangeWasMade
End Sub

Private Sub txtRSpeedStraight_Change()
ChangeWasMade
End Sub

Private Sub txtSpeedLimit_Change()
ChangeWasMade
End Sub

Private Sub txtSpeedStraight_Change()
ChangeWasMade
End Sub

Private Sub txtZHop_Change()
ChangeWasMade
End Sub

Private Sub txtZJerk_Change()
ChangeWasMade
End Sub

Public Sub purgeModified(Optional ByVal index As Long = -2)
If index = -2 Then
  index = Me.FindInCombo(pm.curPresetFN)
  pm.curPresetIsModified = False
End If
If index = -1 Then Exit Sub
Dim txt As String
txt = cmbPreset.List(index)
If right$(txt, 1) = "*" Then
  txt = Left$(txt, Len(txt) - 1)
  cmbPreset.List(index) = txt
End If
End Sub
