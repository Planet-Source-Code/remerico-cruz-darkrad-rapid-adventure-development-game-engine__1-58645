Attribute VB_Name = "modSub"

' Current Room
Type QRoom
  RName As String
  RSaved As Boolean
  RBG As String
  RLocName As String
  
  dNorthWest As String
  dNorth As String
  dNorthEast As String
  dWest As String
  dEast As String
  dSouthWest As String
  dSouth As String
  dSouthEast As String
  
  RMusic As String
  RConst(10) As String
  
  ItmCount As Integer
  Itm(50, 5) As String
  Trans As Integer

  '0 = Name
  '1 = Desc
  '2 = Prog
  '3 = X
  '4 = Y
  '5 = Gfx
  
  AreaCount As Integer
  Area(50, 4) As String
  
  '0 = Area Name
  '1 = Desc
  '2 = Prog
  '3 = X
  '4 = Y
  
  DontSave As Integer
  DontShowNav As Integer
  DisableMenu As Integer
  DisableTooltip As Integer
  
End Type
Public DaRoom As QRoom

' Current Item
Type QItem
  itmName As String
  itmDescription As String
  itmGfx As String
  
  IsRmv As Integer
End Type
Public DaItem As QItem

' Game-related Vars
Public GameName As String

' IDE Environment Variables
Public IsSaved As Boolean
Public WrkDir As String
Public VMajor As Integer
Public VMinor As Integer

' These variables serve to pass on info between forms
Public uGfx As String
Public uLst As ListBox
Public uDiag As String
Public uArea As String
Public uDesc As String
Public uEdit As Boolean
Public uItm As String

Sub PackGame(dFName As String, Optional IsNewGame As Boolean = False)
On Error GoTo WErr
' Pack 'em up in one neat pile...:)
Dim PName As String
Dim FCount, FCur
Dim IsProgr As Boolean

IsProgr = True

  frmWait.Show

' We'll....umm....count how many files we have first...
If IsProgr = True Then
  FCount = 0
  PName = Dir(WrkDir + "*.*")
  Do Until PName = ""
    FCount = FCount + 1
    frmWait.msg.Caption = FCount & " files found..."
    frmWait.Refresh
    PName = Dir()
  Loop
End If

  ' Copy da old file as backup in case something bad happens
  If IsNewGame = False Then FileCopy dFName, dFName + ".bak"
  
  ' Now kill (delete) the original file (Appending to it
  ' will make it grow bigger
  If IsNewGame = False Then Kill dFName

  FCur = 0
  PName = Dir(WrkDir + "*.*")
  Do Until PName = ""
    JpkAdd dFName, WrkDir + PName, PName
    If IsProgr = True Then
      FCur = FCur + 1
      Progress frmWait.Progr, GetPercent(FCur, FCount)
      frmWait.msg.Caption = "Saving " & PName & "..."
      frmWait.Refresh
    End If
    PName = Dir()
  Loop
  IsSaved = True
  If IsProgr = True Then
    frmWait.msg.Caption = "Saving done."
    frmWait.Refresh
  End If
  
  If IsNewGame = False Then Kill dFName + ".bak"
  If IsProgr = True Then Unload frmWait
  
Exit Sub
WErr:
Select Case Err.Number
  Case 401
    IsProgr = False
    Resume Next
  Case Else
    MsgBox "An error in saving occurred. " & Err.Description & vbCrLf & vbCrLf & "If '" & GetFName(dFName) & "' becomes corrupted, try opening the backup file '" & GetFName(dFName) & ".bak'.", vbCritical
    Exit Sub
End Select
End Sub

Sub ExtractGame(dFName As String)
' Open it, let's get down and dirty....:P
Dim PName As String

If JpkList(dFName, frmMain.LstFiles) = True Then
  For a = 0 To frmMain.LstFiles.ListCount - 1
    JpkExtract dFName, frmMain.LstFiles.List(a), WrkDir + frmMain.LstFiles.List(a)
  Next a
Else
  MsgBox "Failed to open " & dFName, vbCritical
End If
End Sub

Sub DeleteTmp()
' Deletes Temporary Working Files
Dim PName As String

  PName = Dir(WrkDir + "*.*")
  Do Until PName = ""
    Kill WrkDir + PName
    PName = Dir()
  Loop

End Sub

Sub SaveSetting()
SetInitEntry "Game", "FName", GameName
End Sub

Sub LoadSetting()
GameName = GetInitEntry("Game", "FName")
End Sub

Sub FillGfx(dCmb As ComboBox)
Dim dPic As String

dCmb.Clear

' Fill up TitleScreen options with the pics

' *.bmp;*.gif;*.jpg;*.ico;*.wmf
dPic = Dir(WrkDir + "*.*")
Do Until dPic = ""
  Select Case LCase$(Right$(dPic, 3))
    Case "bmp", "gif", "jpg", "wmf"
      dCmb.AddItem dPic
  End Select
  dPic = Dir()
Loop

End Sub

Sub FillRoom(dCmb As ComboBox)
Dim dRm As String

dCmb.Clear

dRm = Dir(WrkDir + "*.rm")
Do Until dRm = ""
  dCmb.AddItem dRm
  dRm = Dir()
Loop

End Sub

Sub FillDialog(dCmb As ComboBox)
Dim dDg As String

dCmb.Clear

dDg = Dir(WrkDir + "*.dg")
Do Until dDg = ""
  dCmb.AddItem dDg
  dDg = Dir()
Loop

End Sub

Sub FillMusic(dCmb As ComboBox)
Dim dMsc As String

dCmb.Clear

dMsc = Dir(WrkDir + "*.*")
Do Until dMsc = ""
  Select Case LCase$(Right$(dMsc, 3))
    Case "wav", "mp3", "mid"
      dCmb.AddItem dMsc
  End Select
  dMsc = Dir()
Loop

End Sub

Sub FillItem(dCmb As ComboBox)
Dim dItm As String

dCmb.Clear

dItm = Dir(WrkDir + "*.itm")
Do Until dItm = ""
  dCmb.AddItem dItm
  dItm = Dir()
Loop

End Sub

Function GetChar(wKey As Integer)
' Taken directly from our school project game "Visual Revolution"

Select Case wKey
  Case 8
    GetChar = "[BACKSPACE]"
  Case 13
    GetChar = "[ENTER]"
  Case 16
    GetChar = "[SHIFT]"
  Case 17
    GetChar = "[CTRL]"
  Case 18
    GetChar = "[ALT]"
  Case 32
    GetChar = "[SPACE]"
  Case 33
    GetChar = "[PGUP]"
  Case 34
    GetChar = "[PGDOWN]"
  Case 35
    GetChar = "[END]"
  Case 36
    GetChar = "[HOME]"
  Case 37
    GetChar = "[LEFT]"
  Case 38
    GetChar = "[UP]"
  Case 39
    GetChar = "[RIGHT]"
  Case 40
    GetChar = "[DOWN]"
  Case 45
    GetChar = "[INSERT]"
  Case 46
    GetChar = "[DEL]"
  Case 92
    GetChar = "[WINDOWS]"
  Case 93
    GetChar = "[MENU]"
  Case 96
    GetChar = "[NUM 0]"
  Case 97
    GetChar = "[NUM 1]"
  Case 98
    GetChar = "[NUM 2]"
  Case 99
    GetChar = "[NUM 3]"
  Case 100
    GetChar = "[NUM 4]"
  Case 101
    GetChar = "[NUM 5]"
  Case 102
    GetChar = "[NUM 6]"
  Case 103
    GetChar = "[NUM 7]"
  Case 104
    GetChar = "[NUM 8]"
  Case 105
    GetChar = "[NUM 9]"
  Case 106
    GetChar = "*"
  Case 107
    GetChar = "+"
  Case 109
    GetChar = "-"
  Case 110
    GetChar = "."
  Case 111
    GetChar = "/"
  Case 186
    GetChar = ";"
  Case 187
    GetChar = "="
  Case 188
    GetChar = ","
  Case 189
    GetChar = "-"
  Case 190
    GetChar = "."
  Case 191
    GetChar = "/"
  Case 192
    GetChar = "`"
  Case 220
    GetChar = "\"
  Case 222
    GetChar = "'"
  Case 226
    GetChar = "\"
  Case Else
    GetChar = Chr(wKey)
End Select
End Function

Sub DTip(wTitle As String, wInst As String)
If frmMain.tTitle.Caption <> wTitle Then frmMain.tTitle.Caption = wTitle
If frmMain.lblInst.Caption <> wInst Then frmMain.lblInst.Caption = wInst
End Sub

Sub VSwap(X1 As Variant, X2 As Variant)

    Dim T1 As Variant
    T1 = X1
    X1 = X2
    X2 = T1
End Sub
