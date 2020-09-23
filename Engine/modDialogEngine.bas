Attribute VB_Name = "modDialogEngine"
' The Dialog Engine, with KAPRE® Technology
' All the code-crunching code are in here.

Private Declare Function GetTickCount Lib "kernel32" () As Long

Type DialogScript
  theCode() As String
  CLine As Long
  include(50) As String
  Length As Long
End Type

Dim CLine As Integer
Dim Run2Prog As String

Sub ReadDialog(wDg As String)
' This is where we read and execute the dialog script
Dim dLine As String
Dim DgType As String
Dim FNum

FNum = FreeFile

' We...um...read all code first...
Echo vbCrLf + "*** Loading Dialog Script... ***"
frmLstFiles.lstCode.Clear
frmLstFiles.lstLbl.Clear
CLine = 0
uChoiceMode = False


Open WrkDir + wDg For Input As #FNum
  Run2Prog = ""
  Do Until EOF(FNum)
    Line Input #FNum, dLine
    frmLstFiles.lstCode.AddItem dLine
    Echo "[" & CLine + 1 & "]  " & dLine
    
    ' If the line is a label...we'll store it for later
    ' purposes....
    If Left$(LTrim$(dLine), 1) = ":" Then
      frmLstFiles.lstLbl.AddItem dLine
      frmLstFiles.lstLbl.ItemData(frmLstFiles.lstLbl.ListCount - 1) = CLine
    End If
    
    CLine = CLine + 1
  Loop
Close #FNum
Echo CLine & " lines found."
Echo frmLstFiles.lstLbl.ListCount & " labels found."
Echo "*** Loading " + wDg + " complete. ***" + vbCrLf

'frmLstFiles.Show

' Then we parse the code line by line here...
' So...we can actually say the center of the
' dialog universe is here!
For CLine = 0 To frmLstFiles.lstCode.ListCount - 1
    dLine = frmLstFiles.lstCode.List(CLine)
    DgType = Left$(LTrim$(dLine), 1)
    
    ' For the #Choice purposes. geez.
    If uChoiceMode = True Then
      ' If the next line is no longer a #choice command, display all choice
      If DgType <> "#" Or LCase$(Left$(LTrim$(dLine), 7)) <> "#choice" Then DisplayChoice
    End If
    
    ' We check what type of line it is...
    Select Case DgType
      Case "#"   ' It's a code! Execute it!
        ExecuteLine Trim$(dLine), CLine
      Case ":", "", "*"   ' It's either a label, blank space, or comment
        ' wala lang
      Case Else
        ' this is supposed to be a conversation...!
        Conv dLine
    End Select
Next CLine
Echo "Execution done."

' Code if running a dialog script from another script
If Run2Prog <> "" Then ReadDialog Run2Prog

End Sub

Sub ExecuteLine(wLine As String, wNum As Integer)
' Code Execution begins here

On Error GoTo ErTrap

Dim dCmd As String
Dim dOpeningP As Integer
Dim dClosingP As Integer
Dim dArguments() As String

dOpeningP = InStr(wLine, "(")
dClosingP = InStrRev(wLine, ")")
dOpeningP = dOpeningP + 1

If dOpeningP = 0 Or dClosingP = 0 Then Echo "ERROR: ( ) Expected on line " & wNum, True: Exit Sub

dCmd = Mid$(wLine, 2, dOpeningP - 3)
Echo "Executing '" & dCmd & "' command on line " & wNum + 1 & "..."
dArguments = Split(Mid$(wLine, dOpeningP, dClosingP - dOpeningP), ",")

' Looks for the next #choice command, if any
If uChoiceMode = True And LCase$(dCmd) <> "choice" Then
  DisplayChoice
  If Left$(wLine, 1) <> "#" Then Exit Sub
End If


Select Case UCase$(dCmd)
  Case "ADDINVENTORY"
    ProcessVar dArguments(0)
    
    AddInventory dArguments(0)
    
  Case "CHOICE"
    uChoiceMode = True
    AddChoice Mid$(dArguments(0), 2, Len(dArguments(0)) - 2), dArguments(1)
    
  Case "DISPLAYGFX"
    ReDim Preserve dArguments(2)
    Dim Sw As Long, Sh As Long

    ProcessVar dArguments(0)
    
    If Dir$(WrkDir + dArguments(0)) = "" Then
      Echo "ERROR: '" & dArguments(0) & "' not found!", True
      Exit Sub
    End If
    
    GetImgSize WrkDir + dArguments(0), Sw, Sh
    'MsgBox Sw & " " & Sh
    
    If dArguments(1) = "" Then dArguments(1) = ((640 * 15) / 2) - ((Sw * 15) / 2) Else dArguments(1) = dArguments(1) * 15
    If dArguments(2) = "" Then dArguments(2) = ((370 * 15) / 2) - ((Sh * 15) / 2) Else dArguments(2) = dArguments(2) * 15
  
    frmMain.dScreen.PaintPicture LoadPicture(WrkDir + dArguments(0)), Val(dArguments(1)), Val(dArguments(2))
    
  Case "END"
    CLine = frmLstFiles.lstCode.ListCount - 1
    
  Case "GOTOLINE"
    GoToLine dArguments(0)
    
  Case "GOTOROOM"
    ProcessVar dArguments(0)
  
    GotoRoom dArguments(0)
    
  Case "IF"
    If Evaluate(dArguments(0)) = True Then GoToLine dArguments(1)
    
  Case "PRINT"
  
    ProcessVar dArguments(0)
    
    If GScreen = "game" And frmMain.dNav.Visible = True Then
        frmMain.dMsg = dArguments(0)
    Else
        MsgBox dArguments(0)
    End If
    
  Case "PLAYMUSIC"
    PlayMusic dArguments(0)
    
  Case "RUN"
    ProcessVar dArguments(0)
    
    Run2Prog = dArguments(0)
    
    If Run2Prog <> "" Then CLine = frmLstFiles.lstCode.ListCount - 1
 
  Case "MUSICVOL"
    ProcessVar dArguments(0)
    
    MusicVol CLng(dArguments(0))
     
  Case "SETVAR"
    SetVar dArguments(0), dArguments(1)
     
  Case "STOPMUSIC"
    StopMusic
  
  Case "TRANS"
  
  Case "WAIT"
    ProcessVar dArguments(0)
    If IsNumeric(dArguments(0)) Then Wait dArguments(0)
    
  Case "WIN"
    Unload frmMain
    DeleteTmp
    End
  
End Select

Exit Sub

ErTrap:

Select Case Err.Number
  Case 0
    Resume Next
  Case Else
    Echo "ERROR: Parse Error [" & Err.Description & "]"
    Resume ExtExec
End Select

ExtExec:

End Sub

Sub Conv(dConv As String)
' The conversation thingie is all here!!
' Isn't that kewl!!! w00t!!1
dConv = Replace(dConv, "\n", vbCrLf)
uOK = False
frmMsg.dMsg = dConv
frmMsg.Show 'vbModal
Do Until uOK = True
  DoEvents
Loop
End Sub

Sub AddChoice(wChoice As String, wLbl As String)
With frmChoice
If .dChoice.Count = 1 And .dChoice(0).Caption = "" Then
  .dChoice(0).Caption = wChoice
  .dChoice(0).BackStyle = 1
  .lstChoice.AddItem wLbl
Else
  Load .dChoice(.dChoice.Count)
  .dChoice(.dChoice.UBound).Visible = True
  .dChoice(.dChoice.UBound).Move 240, .dChoice(.dChoice.UBound - 1).Top + .dChoice(.dChoice.UBound - 1).Height + 105
  .dChoice(.dChoice.UBound).Caption = wChoice
  .dChoice(.dChoice.UBound).BackStyle = 0
  .lstChoice.AddItem wLbl
  .shape.Height = .dChoice(.dChoice.UBound).Top + .dChoice(.dChoice.UBound).Height
  .Height = .shape.Height + 240
End If
End With
End Sub

Sub DisplayChoice()
  uChoiceMode = False
  frmChoice.Show 'vbModal
  StayOnTop frmChoice
  frmChoice.xLoop
End Sub

Sub GoToLine(ByRef wLbl As String)
' Go to the line label specified
Dim lblFound As Boolean
    lblFound = False
    For a = 0 To frmLstFiles.lstLbl.ListCount - 1
      If Left$(LTrim$(frmLstFiles.lstLbl.List(a)), 1) <> ":" Then frmLstFiles.lstLbl.List(a) = ":" + LTrim$(frmLstFiles.lstLbl.List(a))
      If Left$(LTrim$(wLbl), 1) <> ":" Then wLbl = ":" + LTrim$(wLbl)
      
      If LCase$(frmLstFiles.lstLbl.List(a)) = LCase$(wLbl) Then
        lblFound = True
        CLine = frmLstFiles.lstLbl.ItemData(a)
        Echo "Skipping to line " & frmLstFiles.lstLbl.ItemData(a) + 1 & "..."
        Exit For
      End If
    Next a
    If lblFound = False Then Echo "ERROR: Label '" + wLbl + "' not found."
End Sub

Sub SetVar(wVar As String, wVal As String)
' Heehee...just a simple variable system that stores all variables in a collection... :)
Dim VExist As Boolean
Dim DaVal As String

VExist = False

' We'll first check if the variable already existed
For a = 0 To frmLstFiles.lstVar.ListCount - 1
  If LCase$(Trim$(wVar)) = LCase$(frmLstFiles.lstVar.List(a)) Then
    VExist = True
    Exit For
  End If
Next a

ProcessVar wVal
DaVal = wVal

If VExist = False Then  ' If the variable doesn't exist....
  DVar.Add DaVal, LCase$(Trim$(wVar))
  frmLstFiles.lstVar.AddItem LCase$(Trim$(wVar))
Else   ' and if it already existed...
  DVar.Remove a + 1
  frmLstFiles.lstVar.RemoveItem a
  
  DVar.Add DaVal, LCase$(Trim$(wVar))
  frmLstFiles.lstVar.AddItem LCase$(Trim$(wVar))
End If


End Sub

Function GetVar(wVar As String) As String
' Get the value of the variable
Dim VExist As Boolean

VExist = False

' We'll first check if the variable is existing
For a = 0 To frmLstFiles.lstVar.ListCount - 1
  If LCase$(Trim$(wVar)) = LCase$(frmLstFiles.lstVar.List(a)) Then
    VExist = True
    Exit For
  End If
Next a

If VExist = False Then  ' If the variable doesn't exist....
  GetVar = ""
Else   ' and if it already existed...
  GetVar = DVar(a + 1)
End If

End Function

Sub PlayMusic(wFile As String)
Dim MscFile As String


ProcessVar wFile
MscFile = wFile

If Trim$(MscFile = "") Then Exit Sub

GMusic = MscFile

Select Case LCase$(Right$(MscFile, 3))
  Case "mid"
    PlayMIDI WrkDir + MscFile
End Select

End Sub

Sub Wait(ByVal TimeToWait As Long)

Dim EndTime As Long
EndTime = GetTickCount + TimeToWait

Do Until GetTickCount > EndTime
  DoEvents
Loop

End Sub

Sub ProcessVar(ByRef wVar As String)

' Is it a literal value or a variable?
If (InStr(wVar, "<") > 0 And InStr(wVar, ">") > 0) Or (Left$(LTrim$(wVar), 1) = Chr(34) And Right$(RTrim$(wVar), 1) = Chr(34) And Len(Trim$(wVar)) > 1) Then
  ' Value is a literal...
  wVar = Mid$(Trim(wVar), 2, Len(wVar) - 2)
  
ElseIf IsNumeric(Trim$(Left$(wVar, 1))) = True Then
  ' Value is numeric
  ' Well, we won't do any string manipulations, then!
  
Else  ' Value is a variable
  wVar = GetVar(wVar)
End If
End Sub

Sub Trans(wTran As Integer, wPic1 As PictureBox, wPic2 As PictureBox)
Dim a As Long

If wTran = 0 Then Exit Sub

If wTran <> 1 Then frmMain.dScreen.AutoRedraw = False

Select Case wTran
  Case 1  ' Fade
    For a = 1 To 255 Step 55
      Blend a, wPic1, wPic2
    Next a
  Case 2  ' Wipe
    Wipe wPic1, wPic2
  Case 3  'Hour Double
    HourDouble wPic1, wPic2
  Case 4  ' Hour Inverse
    HourInverse wPic1, wPic2
  Case 5  ' Circle
    CircleTrans wPic1, wPic2
  Case 6  ' Implode
    Implode wPic1, wPic2
  Case 7  ' Tenda
    Tenda wPic1, wPic2
End Select

frmMain.dScreen.AutoRedraw = True

End Sub

Sub AddInventory(wItem As String)

If Dir$(WrkDir + wItem) = "" Then Exit Sub

Inventory.Add wItem
End Sub

Function Evaluate(wText As String) As Boolean
Dim dLength As Long, dVal1 As String, dVal2 As String
Dim dPart As String, startAt As Long, EqType As String
Dim ReturnVal As Boolean

dLength = Len(wText)
dVal1 = ""
dVal2 = ""

    'Get first variable
    For a = 1 To dLength
        dPart = Mid$(wText, a, 1)
        If dPart = "=" Or dPart = "~" Or dPart = ">" Or dPart = "<" Then
            'Found equality operator
            EqType = dPart
            startAt = a
            a = dLength
        Else
            If dPart <> " " Then dVal1 = dVal1 + dPart
        End If
    Next a
    
    ' Get the equation
    For a = startAt + 1 To dLength
        dPart = Mid$(wText, a, 1)
        
        If dPart <> " " Then
            If dPart = "=" Or dPart = ">" Or dPart = "<" Then
                EqType = EqType + dPart
                startAt = a + 1
                a = dLength
            Else
                startAt = a
                a = dLength
            End If
        End If
    Next a
    
    'Now get the other variable
    For a = startAt To dLength
        dPart = Mid$(wText, a, 1)
        If dPart <> " " Then dVal2 = dVal2 + dPart
    Next a

    
    ProcessVar dVal1
    ProcessVar dVal2
    
    'MsgBox dVal1
    'MsgBox EqType
    'MsgBox dVal2


    If (IsNumeric(dVal1) = False And IsNumeric(dVal1) = True) Or (IsNumeric(dVal1) = True And IsNumeric(dVal1) = False) Then Evaluate = False: Exit Function
    
    If dVal1 = "" And dVal2 = "" Then Evaluate = False: Exit Function
    
    If dVal1 <> "" And dVal2 = "" Then
        If IsNumeric(dVal1) = True Then
            Evaluate = False: Exit Function
        Else
            Evaluate = True: Exit Function
        End If
    End If

    ReturnVal = False
    
    If EqType = "=" Or EqType = "==" Then
        If IsNumeric(dVal1) = True And IsNumeric(dVal2) = True Then
            'numerical
            If Val(dVal1) = Val(dVal2) Then ReturnVal = True
        Else
            If dVal1 = dVal2 Then ReturnVal = True
        End If
    End If
    
    If EqType = "~" Or EqType = "~=" Or EqType = "=~" Then
        If IsNumeric(dVal1) = True And IsNumeric(dVal2) = True Then
            'numerical
            If Val(dVal1) <> Val(dVal2) Then ReturnVal = True
        Else
            If dVal1 <> dVal2 Then ReturnVal = True
        End If
    End If
    
    If EqType = "<=" Or EqType = "=<" Then
        If IsNumeric(dVal1) = True And IsNumeric(dVal2) = True Then
            'numerical
            If Val(dVal1) <= Val(dVal2) Then ReturnVal = True
        Else
            ReturnVal = False
        End If
    End If
    
    If EqType = ">=" Or EqType = "=>" Then
        If IsNumeric(dVal1) = True And IsNumeric(dVal2) = True Then
            'numerical
            If Val(dVal1) >= Val(dVal2) Then ReturnVal = True
        Else
            ReturnVal = False
        End If
    End If
    
    If EqType = ">" Or EqType = ">" Then
        If IsNumeric(dVal1) = True And IsNumeric(dVal2) = True Then
            'numerical
            If Val(dVal1) > Val(dVal2) Then ReturnVal = True
        Else
            ReturnVal = False
        End If
    End If
    If EqType = "<" Or EqType = "<" Then
        If IsNumeric(dVal1) = True And IsNumeric(dVal2) = True Then
            'numerical
            If Val(dVal1) < Val(dVal2) Then ReturnVal = True
        Else
            ReturnVal = False
        End If
    End If
Evaluate = ReturnVal

End Function

Sub MusicVol(wVol As Long)


If GMusic = "" Then Exit Sub

Select Case LCase$(Right$(GMusic, 3))
  Case "mid"
    MIDIVolume wVol
End Select

End Sub
