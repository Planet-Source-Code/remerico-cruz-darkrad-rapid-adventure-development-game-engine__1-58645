Attribute VB_Name = "modGameSubs"
' Routine Engine
' The frequently used routines that happens in the game
' (like going from room to room or checking the save game)
' are here.

Function DisplayTitleScreen() As Boolean
Dim dTScrn
dTScrn = GetInitEntry("Init", "TitleScreen", , WrkDir + "Config.cfg")
If dTScrn <> "" Then
  frmMain.dScreen.Picture = LoadPicture(WrkDir + dTScrn)
  DisplayTitleScreen = True
Else
  DisplayTitleScreen = False
End If

End Function

Function SaveGame(SaveFile) As Boolean
  
  If Dir$(SaveFile) <> "" Then Kill SaveFile
  
  SetInitEntry "Game", "Room", GRoom, SaveFile
  SetInitEntry "Game", "Screen", GScreen, SaveFile
  SetInitEntry "Game", "Music", GMusic, SaveFile
  
  SetInitEntry "Variables", "Count", frmLstFiles.lstVar.ListCount, SaveFile
  For a = 1 To frmLstFiles.lstVar.ListCount
    SetInitEntry "Variables", Str(a), Trim$(frmLstFiles.lstVar.List(a - 1)) + "|" + DVar(frmLstFiles.lstVar.List(a - 1)), SaveFile
  Next a
  
  SetInitEntry "Inventory", "Count", Inventory.Count, SaveFile
  For a = 1 To Inventory.Count
    SetInitEntry "Inventory", Str(a), Inventory(a), SaveFile
  Next a
  
End Function

Function LoadGame(SaveFile) As Boolean
Dim InvCount As Integer, VarCount As Integer
  
  If Dir$(SaveFile) <> "" Then Kill SaveFile
  
  GetInitEntry "Game", "Room", , SaveFile
  GetInitEntry "Game", "Screen", , SaveFile
  GetInitEntry "Game", "Music", , SaveFile
  
  VarCount = GetInitEntry("Variables", "Count", 0, SaveFile)
  
  'TODO: Set the variable thingies
  'For a = 1 To frmLstFiles.lstVar.ListCount
  '  GetInitEntry "Variables", Str(a), Trim$(frmLstFiles.lstVar.List(a - 1)) + "|" + DVar(frmLstFiles.lstVar.List(a - 1)), SaveFile
  'Next a
  
  
  InvCount = GetInitEntry("Inventory", "Count", 0, SaveFile)
  'TODO: Set the Inventory item thingies
  'For a = 1 To Inventory.Count
  '  GetInitEntry "Inventory", Str(a), Inventory(a), SaveFile
  'Next a
  
End Function

Sub GotoRoom(wRoom As String)

If Trim$(wRoom) = "" Then Exit Sub

GScreen = "loadingroom"
GRoom = wRoom

AllAreas.ClearAll

' Clear all previous items
Do Until frmMain.xItm.Count = 1
  Unload frmMain.xItm(frmMain.xItm.UBound)
Loop

  'Does that room exist? In the real world?? Just kidding.. :)
  If Dir(WrkDir + wRoom) = "" Then Echo "ERROR: Room '" & wRoom & "' does not exist!", True: Exit Sub

  ' Now load all info from that room...
  With DaRoom
    .RName = wRoom
    .RBG = GetInitEntry("Room", "BG", , WrkDir + .RName)
    .RMusic = GetInitEntry("Room", "Music", , WrkDir + .RName)
    .Desc = GetInitEntry("Room", "Description", , WrkDir + .RName)
    
    .Trans = GetInitEntry("Room", "Trans", 0, WrkDir + .RName)
    
    .dNorthWest = GetInitEntry("Dir", "NW", , WrkDir + .RName)
    .dNorth = GetInitEntry("Dir", "N", , WrkDir + .RName)

    .dNorthEast = GetInitEntry("Dir", "NE", , WrkDir + .RName)
    .dWest = GetInitEntry("Dir", "W", , WrkDir + .RName)
    .dEast = GetInitEntry("Dir", "E", , WrkDir + .RName)

    .dSouthWest = GetInitEntry("Dir", "SW", , WrkDir + .RName)
    .dSouth = GetInitEntry("Dir", "S", , WrkDir + .RName)
    .dSouthEast = GetInitEntry("Dir", "SE", , WrkDir + .RName)
    
    .RLocName = GetInitEntry("Room", "Location", , WrkDir + .RName)
    .ItmCount = Val(GetInitEntry("Room", "ItemCount", , WrkDir + .RName))
    .AreaCount = Val(GetInitEntry("Room", "AreaCount", , WrkDir + .RName))
    
    .DontSave = GetInitEntry("Settings", "DontSave", vbUnchecked, WrkDir + .RName)
    .DontShowNav = GetInitEntry("Settings", "DontShowNav", vbUnchecked, WrkDir + .RName)
    .DisableMenu = GetInitEntry("Settings", "DisableMenu", vbUnchecked, WrkDir + .RName)
    .DisableTooltip = GetInitEntry("Settings", "DisableTooltip", vbUnchecked, WrkDir + .RName)
    
      ' The items...
    For a = 1 To .ItmCount
      .Itm(a, 0) = GetInitEntry("Item" & Trim$(Str(a)), "Name", , WrkDir + .RName)
      .Itm(a, 1) = GetInitEntry("Item" & Trim$(Str(a)), "Desc", , WrkDir + .RName)
      If .Itm(a, 1) = "" Then .Itm(a, 1) = GetInitEntry("Item", "Description", , WrkDir + .Itm(a, 0))
      .Itm(a, 2) = GetInitEntry("Item" & Trim$(Str(a)), "Dialog", , WrkDir + .RName)
      .Itm(a, 3) = GetInitEntry("Item" & Trim$(Str(a)), "X", , WrkDir + .RName)
      .Itm(a, 4) = GetInitEntry("Item" & Trim$(Str(a)), "Y", , WrkDir + .RName)
      .Itm(a, 5) = GetInitEntry("Item" & Trim$(Str(a)), "Gfx", , WrkDir + .RName)
    Next a
    
    For a = 1 To .AreaCount
      .Area(a, 0) = GetInitEntry("Area" & Trim$(Str(a)), "Name", , WrkDir + .RName)
      .Area(a, 1) = GetInitEntry("Area" & Trim$(Str(a)), "Desc", , WrkDir + .RName)
      .Area(a, 2) = GetInitEntry("Area" & Trim$(Str(a)), "Dialog", , WrkDir + .RName)
      .Area(a, 3) = GetInitEntry("Area" & Trim$(Str(a)), "X", , WrkDir + .RName)
      .Area(a, 4) = GetInitEntry("Area" & Trim$(Str(a)), "Y", , WrkDir + .RName)
      frmMain.LoadRegion a
    Next a
    
    Echo vbCrLf & "<[---------[ " & .RName & " ]---------]>"
    Echo "NW - " & .dNorthWest
    Echo "N - " & .dNorth
    Echo "NE - " & .dNorthEast
    Echo "W - " & .dWest
    Echo "E - " & .dEast
    Echo "SW - " & .dSouthWest
    Echo "S - " & .dSouth
    Echo "SE - " & .dSouthEast
    Echo "<[------------[ End ]------------]>"
  End With
  
  ' Sets the proper arrows that is available for navigation
  SetDArrows
  
  ' Plays the ambience music
  If GMusic <> DaRoom.RMusic Then
    MIDIVolume -900
  End If
  
  ' Load the pic, if it exist, of course...
  If Dir$(WrkDir + DaRoom.RBG) <> "" And Trim$(DaRoom.RBG) <> "" Then
    frmMain.pTemp.Picture = LoadPicture(WrkDir + DaRoom.RBG)
    Trans DaRoom.Trans, frmMain.pTemp, frmMain.dScreen
    frmMain.dScreen.Picture = LoadPicture(WrkDir + DaRoom.RBG)
  Else
    Echo "ERROR: Picture '" & DaRoom.RBG & "' does not exist!", True
    frmMain.dScreen.Picture = LoadPicture()
  End If
  
  For a = 1 To DaRoom.ItmCount
    frmMain.LoadItmImg DaRoom.Itm(a, 0), Val(DaRoom.Itm(a, 3)), Val(DaRoom.Itm(a, 4))
  Next a
  
  If GMusic <> DaRoom.RMusic Then
    StopMIDI
    If DaRoom.RMusic <> "" Then PlayMIDI WrkDir + DaRoom.RMusic
    MIDIVolume 0
    GMusic = DaRoom.RMusic
  End If
  
  frmMain.dMsg.Caption = DaRoom.Desc
  
  If DaRoom.DontShowNav = vbUnchecked Then frmMain.dNav.Visible = True Else frmMain.dNav.Visible = False
  GScreen = "game"
  
End Sub

Sub SetDArrows()
' Sets the arrows that tell you where you can do and
' where you can't.

With DaRoom
  If .dNorthWest = "" Then frmMain.dDir(0).Visible = False Else frmMain.dDir(0).Visible = True
  If .dNorth = "" Then frmMain.dDir(1).Visible = False Else frmMain.dDir(1).Visible = True
  If .dNorthEast = "" Then frmMain.dDir(2).Visible = False Else frmMain.dDir(2).Visible = True
  If .dWest = "" Then frmMain.dDir(3).Visible = False Else frmMain.dDir(3).Visible = True
  If .dEast = "" Then frmMain.dDir(5).Visible = False Else frmMain.dDir(5).Visible = True
  If .dSouthWest = "" Then frmMain.dDir(6).Visible = False Else frmMain.dDir(6).Visible = True
  If .dSouth = "" Then frmMain.dDir(7).Visible = False Else frmMain.dDir(7).Visible = True
  If .dSouthEast = "" Then frmMain.dDir(8).Visible = False Else frmMain.dDir(8).Visible = True
End With
End Sub

Sub ApplyTheme(wTheme As String)
Dim dProp As String
Dim GUIFile As String

If wTheme = "menu" Then

  ' The Menu theme

  GUIFile = "menu.gui"

  dProp = GetInitEntry("Menu Window", "Background", , WrkDir + GUIFile)
  If dProp <> "" Then frmMain.dMenu.Picture = LoadPicture(WrkDir + dProp)

  dProp = GetInitEntry("Menu Window", "ShowBorder", , WrkDir + GUIFile)
  If LCase$(dProp) = "false" Then frmMain.Shape5.Visible = False

  For a = 0 To 5
    dProp = GetInitEntry("Menu Text " & a, "Caption", , WrkDir + GUIFile)
    If dProp <> "" Then frmMain.dIMnu(a).Caption = dProp
  Next a

ElseIf wTheme = "exit" Then

  ' The Exit Dialog

  GUIFile = "exit.gui"

  dProp = GetInitEntry("Exit Window", "Background", , WrkDir + GUIFile)
  If dProp <> "" Then frmExit.Picture = LoadPicture(WrkDir + dProp)

  dProp = GetInitEntry("Exit Window", "ShowBorder", , WrkDir + GUIFile)
  If LCase$(dProp) = "false" Then frmExit.shape.Visible = False

  dProp = GetInitEntry("Exit Prompt", "Caption", , WrkDir + GUIFile)
  If dProp <> "" Then frmExit.dPrompt.Caption = dProp

  dProp = GetInitEntry("'No' Button", "Caption", , WrkDir + GUIFile)
  If dProp <> "" Then frmExit.PNo.Caption = dProp

  dProp = GetInitEntry("'Yes' Button", "Caption", , WrkDir + GUIFile)
  If dProp <> "" Then frmExit.PYes.Caption = dProp
  
ElseIf wTheme = "nav" Then

  GUIFile = "nav.gui"

  dProp = GetInitEntry("Navigation Bar", "Background", , WrkDir + GUIFile)
  If dProp <> "" Then frmMain.dNav.Picture = LoadPicture(WrkDir + dProp)

End If

End Sub

Sub ShowSave(Optional ISave As Boolean = False)
If ISave Then SaveMode = "save" Else SaveMode = "load"
'MsgBox SaveMode
frmSave.Show vbModal
End Sub
