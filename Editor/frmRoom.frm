VERSION 5.00
Begin VB.Form frmRoom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Room Editor"
   ClientHeight    =   6750
   ClientLeft      =   1785
   ClientTop       =   6300
   ClientWidth     =   9600
   Icon            =   "frmRoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox DlgRoom 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   2280
      ScaleHeight     =   3465
      ScaleWidth      =   7185
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   7215
      Begin VB.FileListBox RoomLst 
         Appearance      =   0  'Flat
         Height          =   2175
         Left            =   120
         Pattern         =   "*.rm"
         TabIndex        =   11
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Save/Open"
         Height          =   375
         Left            =   4800
         TabIndex        =   10
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtFName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   6975
      End
      Begin VB.Label PrevDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   4575
      End
      Begin VB.Image imgPrev 
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label RmTitle 
         BackColor       =   &H80000002&
         Caption         =   " Room List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   7215
      End
   End
   Begin VB.PictureBox tDock 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   2
      Top             =   0
      Width           =   9600
      Begin DarkEdit.GurhanButton tlNew 
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Picture         =   "frmRoom.frx":038A
         PictureWidth    =   16
         PictureHeight   =   16
         PictureSize     =   0
         OriginalPicSizeW=   16
         OriginalPicSizeH=   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         SoundOver       =   ""
         SoundClick      =   ""
         MaskColor       =   8421504
      End
      Begin DarkEdit.GurhanButton tlOpen 
         Height          =   375
         Left            =   375
         TabIndex        =   4
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Picture         =   "frmRoom.frx":049C
         PictureWidth    =   16
         PictureHeight   =   16
         PictureSize     =   0
         OriginalPicSizeW=   16
         OriginalPicSizeH=   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         SoundOver       =   ""
         SoundClick      =   ""
         MaskColor       =   8421504
      End
      Begin DarkEdit.GurhanButton tlSave 
         Height          =   375
         Left            =   750
         TabIndex        =   5
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Picture         =   "frmRoom.frx":05AE
         PictureWidth    =   16
         PictureHeight   =   16
         PictureSize     =   0
         OriginalPicSizeW=   16
         OriginalPicSizeH=   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         SoundOver       =   ""
         SoundClick      =   ""
         MaskColor       =   8421504
      End
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Top             =   6120
      Width           =   6135
   End
   Begin VB.PictureBox Container 
      BorderStyle     =   0  'None
      Height          =   5550
      Left            =   0
      ScaleHeight     =   5550
      ScaleWidth      =   9600
      TabIndex        =   14
      Top             =   480
      Width           =   9600
      Begin VB.PictureBox dRoom 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   5550
         Left            =   0
         ScaleHeight     =   5550
         ScaleWidth      =   9600
         TabIndex        =   15
         Top             =   0
         Width           =   9600
         Begin VB.Image xItm 
            Appearance      =   0  'Flat
            Height          =   975
            Index           =   0
            Left            =   -1200
            Top             =   -1040
            Width           =   1335
         End
         Begin VB.Line ALine 
            DrawMode        =   6  'Mask Pen Not
            Index           =   0
            Visible         =   0   'False
            X1              =   120
            X2              =   1800
            Y1              =   120
            Y2              =   120
         End
      End
   End
   Begin VB.Label lblCoord 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   6000
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Room Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   0
      Top             =   6120
      Width           =   1305
   End
   Begin VB.Menu mnuRoom 
      Caption         =   "&File"
      Begin VB.Menu mnuNewRm 
         Caption         =   "&New Room"
      End
      Begin VB.Menu mnuOpenRm 
         Caption         =   "&Open Room"
      End
      Begin VB.Menu mnuSaveRm 
         Caption         =   "&Save Room"
      End
   End
   Begin VB.Menu mnuRmSettings 
      Caption         =   "&Room Settings"
      Begin VB.Menu mnuSetGfx 
         Caption         =   "&Set Room Graphic..."
      End
      Begin VB.Menu mnuRmOpt 
         Caption         =   "&Room Options"
      End
   End
   Begin VB.Menu mnuAddObj 
      Caption         =   "&Add Object"
      Begin VB.Menu mnuAddItm 
         Caption         =   "Add &Item"
      End
      Begin VB.Menu mnuAddDg 
         Caption         =   "Add &Room Area"
      End
   End
   Begin VB.Menu mnuItmContext 
      Caption         =   "&Item Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuItmCaption 
         Caption         =   "Item Menu"
         Enabled         =   0   'False
      End
      Begin VB.Menu sitm 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItmEdit 
         Caption         =   "&Edit this Item...."
      End
      Begin VB.Menu mnuItmDelete 
         Caption         =   "&Delete this Item..."
      End
   End
   Begin VB.Menu mnuAreaContext 
      Caption         =   "&Area Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAreaCaption 
         Caption         =   "Area Menu"
         Enabled         =   0   'False
      End
      Begin VB.Menu sarea 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAreaEdit 
         Caption         =   "&Edit this Area"
      End
      Begin VB.Menu mnuAreaDelete 
         Caption         =   "&Delete this Area..."
      End
   End
End
Attribute VB_Name = "frmRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AllAreas As New Areas
Dim PolyArea() As POINTAPI

Dim AreaNum As Integer
Dim ItmNum As Integer

' Is it in line-dragging mode? (Area Editting)
Dim LMode As Boolean

Dim EditMode As Integer
' 0 - Normal
' 1 - Item Drag
' 2 - Area Edit Mode

' Local Pass-ons
Dim uIndex As Integer
Dim StartX, StartY

Private Sub cmdAction_Click()

If LCase$(Right$(txtFName.Text, 3)) <> ".rm" Then txtFName.Text = txtFName.Text + ".rm"

DaRoom.RName = txtFName.Text
SetCapt DaRoom.RName

Select Case cmdAction.Caption
  Case "Save"
    SaveRoom DaRoom.RName
    PackGame GameName
  Case "Load"
    LoadRoom DaRoom.RName
End Select

imgPrev.Picture = LoadPicture()
DlgRoom.Visible = False

End Sub

Private Sub cmdCancel_Click()
imgPrev.Picture = LoadPicture()
DlgRoom.Visible = False

End Sub

Sub ShowRmList(Optional dSave As Boolean = False)
' This does not save the room....it only shows the room list!!

DlgRoom.Move (Me.ScaleWidth / 2) - (DlgRoom.Width / 2), (Me.ScaleHeight / 2) - (DlgRoom.Height / 2)
DlgRoom.Visible = True
RoomLst.Refresh

If dSave = False Then cmdAction.Caption = "Load": RmTitle = "Open Room" Else cmdAction.Caption = "Save": RmTitle = "Save Room"

End Sub

Sub LoadRoom(dFile As String)

' Clear all areas from previous room
AllAreas.ClearAll

' Clear all previous items
Do Until xItm.Count = 1
  Unload xItm(xItm.UBound)
Loop


' The loading from room files process begins here
With DaRoom

  .RBG = GetInitEntry("Room", "BG", , WrkDir + dFile)
  txtDesc.Text = GetInitEntry("Room", "Description", , WrkDir + dFile)

  If Dir$(WrkDir + .RBG) <> "" And .RBG <> "" Then dRoom.Picture = LoadPicture(WrkDir + DaRoom.RBG) Else: dRoom.Picture = LoadPicture()

  .RSaved = True
  
  .RMusic = GetInitEntry("Room", "Music", , WrkDir + dFile)
  
  .dNorthWest = GetInitEntry("Dir", "NW", , WrkDir + dFile)
  .dNorth = GetInitEntry("Dir", "N", , WrkDir + dFile)
  .dNorthEast = GetInitEntry("Dir", "NE", , WrkDir + dFile)
  .dWest = GetInitEntry("Dir", "W", , WrkDir + dFile)
  .dEast = GetInitEntry("Dir", "E", , WrkDir + dFile)
  .dSouthWest = GetInitEntry("Dir", "SW", , WrkDir + dFile)
  .dSouth = GetInitEntry("Dir", "S", , WrkDir + dFile)
  .dSouthEast = GetInitEntry("Dir", "SE", , WrkDir + dFile)
  
  .RLocName = GetInitEntry("Room", "Location", , WrkDir + dFile)
  .ItmCount = Val(GetInitEntry("Room", "ItemCount", , WrkDir + dFile))
  .AreaCount = Val(GetInitEntry("Room", "AreaCount", , WrkDir + dFile))
  
  .Trans = Val(GetInitEntry("Room", "Trans", , WrkDir + dFile))
  
  .DontSave = GetInitEntry("Settings", "DontSave", vbUnchecked, WrkDir + dFile)
  .DontShowNav = GetInitEntry("Settings", "DontShowNav", vbUnchecked, WrkDir + dFile)
  .DisableMenu = GetInitEntry("Settings", "DisableMenu", vbUnchecked, WrkDir + dFile)
  .DisableTooltip = GetInitEntry("Settings", "DisableTooltip", vbUnchecked, WrkDir + dFile)

  ' The items...
  For a = 1 To .ItmCount
    .Itm(a, 0) = GetInitEntry("Item" & Trim$(Str(a)), "Name", , WrkDir + dFile)
    .Itm(a, 1) = GetInitEntry("Item" & Trim$(Str(a)), "Desc", , WrkDir + dFile)
    .Itm(a, 2) = GetInitEntry("Item" & Trim$(Str(a)), "Dialog", , WrkDir + dFile)
    .Itm(a, 3) = GetInitEntry("Item" & Trim$(Str(a)), "X", , WrkDir + dFile)
    .Itm(a, 4) = GetInitEntry("Item" & Trim$(Str(a)), "Y", , WrkDir + dFile)
    .Itm(a, 5) = GetInitEntry("Item" & Trim$(Str(a)), "Gfx", , WrkDir + dFile)
    LoadItmImg .Itm(a, 0), Val(.Itm(a, 3)), Val(.Itm(a, 4))
  Next a
  
  ' Now were going to load the areas!
  For a = 1 To .AreaCount
    .Area(a, 0) = GetInitEntry("Area" & Trim$(Str(a)), "Name", , WrkDir + dFile)
    .Area(a, 1) = GetInitEntry("Area" & Trim$(Str(a)), "Desc", , WrkDir + dFile)
    .Area(a, 2) = GetInitEntry("Area" & Trim$(Str(a)), "Dialog", , WrkDir + dFile)
    .Area(a, 3) = GetInitEntry("Area" & Trim$(Str(a)), "X", , WrkDir + dFile)
    .Area(a, 4) = GetInitEntry("Area" & Trim$(Str(a)), "Y", , WrkDir + dFile)
    LoadRegion a
  Next a
    
End With

End Sub

Sub LoadRegion(wNum)
Dim yAr() As String


yAr = Split(DaRoom.Area(wNum, 3), ",")
ReDim PolyArea(UBound(yAr))
For a = 0 To UBound(yAr)
  PolyArea(a).x = Val(yAr(a))
Next a

yAr = Split(DaRoom.Area(wNum, 4), ",")
ReDim Preserve PolyArea(UBound(yAr))
For a = 0 To UBound(yAr)
  PolyArea(a).y = Val(yAr(a))
Next a

AddRGNPoly AllAreas, PolyArea, UBound(PolyArea), DaRoom.Area(wNum, 0), "", vbGreen, vbRed, RGN_HS_NOSHADE, RGN_BS_HOLLOW

End Sub

Sub SaveRoom(dFile As String)
' This is where we save the room!

If Dir$(WrkDir + dFile) <> "" And Trim$(dFile) <> "" Then
  ' We'll backup the old file...
  ' If it exists, that is.
  FileCopy WrkDir + dFile, WrkDir + "bak." + dFile

  ' Kill the original file, we're going to have
  ' problems with the old data in it
  Kill WrkDir + dFile
End If


SetInitEntry "Room", "BG", DaRoom.RBG, WrkDir + dFile
SetInitEntry "Room", "Description", txtDesc.Text, WrkDir + dFile

With DaRoom

  SetInitEntry "Dir", "NW", .dNorthWest, WrkDir + dFile
  SetInitEntry "Dir", "N", .dNorth, WrkDir + dFile
  SetInitEntry "Dir", "NE", .dNorthEast, WrkDir + dFile
  SetInitEntry "Dir", "W", .dWest, WrkDir + dFile
  SetInitEntry "Dir", "E", .dEast, WrkDir + dFile
  SetInitEntry "Dir", "SW", .dSouthWest, WrkDir + dFile
  SetInitEntry "Dir", "S", .dSouth, WrkDir + dFile
  SetInitEntry "Dir", "SE", .dSouthEast, WrkDir + dFile
  
  SetInitEntry "Room", "Location", .RLocName, WrkDir + dFile
  SetInitEntry "Room", "Music", .RMusic, WrkDir + dFile
  SetInitEntry "Room", "ItemCount", .ItmCount, WrkDir + dFile
  SetInitEntry "Room", "AreaCount", .AreaCount, WrkDir + dFile
  
  SetInitEntry "Room", "Trans", .Trans, WrkDir + dFile
  
  SetInitEntry "Settings", "DontSave", .DontSave, WrkDir + dFile
  SetInitEntry "Settings", "DontShowNav", .DontShowNav, WrkDir + dFile
  SetInitEntry "Settings", "DisableMenu", .DisableMenu, WrkDir + dFile
  SetInitEntry "Settings", "DisableTooltip", .DisableTooltip, WrkDir + dFile
  
  Dim SiNum As Integer
  SiNum = 0
  For a = 1 To 50
    If Trim$(.Itm(a, 0)) <> "" Then
      SiNum = SiNum + 1
      SetInitEntry "Item" & Trim$(Str(SiNum)), "Name", .Itm(a, 0), WrkDir + dFile
      SetInitEntry "Item" & Trim$(Str(SiNum)), "Desc", .Itm(a, 1), WrkDir + dFile
      SetInitEntry "Item" & Trim$(Str(SiNum)), "Dialog", .Itm(a, 2), WrkDir + dFile
      SetInitEntry "Item" & Trim$(Str(SiNum)), "X", .Itm(a, 3), WrkDir + dFile
      SetInitEntry "Item" & Trim$(Str(SiNum)), "Y", .Itm(a, 4), WrkDir + dFile
      SetInitEntry "Item" & Trim$(Str(SiNum)), "Gfx", .Itm(a, 5), WrkDir + dFile
    End If
  Next a
  
  SiNum = 0
  For a = 1 To 50
    If Trim$(.Area(a, 0)) <> "" Then
      SiNum = SiNum + 1
      SetInitEntry "Area" & Trim$(Str(SiNum)), "Name", .Area(a, 0), WrkDir + dFile
      SetInitEntry "Area" & Trim$(Str(SiNum)), "Desc", .Area(a, 1), WrkDir + dFile
      SetInitEntry "Area" & Trim$(Str(SiNum)), "Dialog", .Area(a, 2), WrkDir + dFile
      SetInitEntry "Area" & Trim$(Str(SiNum)), "X", .Area(a, 3), WrkDir + dFile
      SetInitEntry "Area" & Trim$(Str(SiNum)), "Y", .Area(a, 4), WrkDir + dFile
    End If
  Next a
  
End With

' Kill the backup, it's not needed.
If Dir$(WrkDir + "bak." + dFile) <> "" Then Kill WrkDir + "bak." + dFile

DaRoom.RSaved = True

End Sub

Function SetCapt(dFile As String)
Me.Caption = "Room Editor (" + dFile + ")"
End Function

Private Sub dRoom_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    StartX = x
    StartY = y
End Sub

Private Sub dRoom_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblCoord.Caption = Str(x / 15) & ", " & Str(y / 15)
If EditMode = 1 Then
  xItm(ItmNum).Left = x - (xItm(ItmNum).Width / 2)
  xItm(ItmNum).Top = y - (xItm(ItmNum).Height / 2)
ElseIf EditMode = 2 And LMode = True Then
  ALine(ALine.UBound).X2 = x
  ALine(ALine.UBound).Y2 = y
ElseIf EditMode = 0 And Button < 1 Then
  On Error Resume Next

  Dim i As Integer
  Dim LRGN As Long

  LRGN = IsInRegion(AllAreas, x, y)
  For i = 1 To AllAreas.Count
  If AllAreas(i).AreaNumber = LRGN Then
        
        If AllAreas(i).AreaSelected = False Then
            AllAreas(i).AreaSelected = True
            InvertRgn AllAreas.ParentHDC, AllAreas(i).AreaNumber
        End If
  Else
        If AllAreas(i).AreaSelected = True Then
            AllAreas(i).AreaSelected = False
            InvertRgn AllAreas.ParentHDC, AllAreas(i).AreaNumber
        End If
  End If
  Next i
ElseIf EditMode = 0 And Button = 1 Then
  If (dRoom.Left >= 0 And StartX < x) Or (dRoom.Left <= Container.Width - dRoom.Width And StartX > x) Then
    If dRoom.Left > 0 Then dRoom.Left = 0
    If dRoom.Left < (Container.Width - dRoom.Width) Then dRoom.Left = dRoom.Width - Container.Width
  Else
    dRoom.Left = IIf(x < StartX, dRoom.Left - (StartX - x), dRoom.Left + (x - StartX))
  End If
 
  If (dRoom.Top >= 0 And StartY < y) Or (dRoom.Top <= Container.Height - dRoom.Height And StartY > y) Then
    If dRoom.Top > 0 Then dRoom.Top = 0
    If dRoom.Top < (Container.Height - dRoom.Height) Then dRoom.Top = dRoom.Height - Container.Height
  Else
    dRoom.Top = IIf(y < StartY, dRoom.Top - (StartY - y), dRoom.Top + (y - StartY))
  End If
End If
End Sub

Private Sub dRoom_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If EditMode = 2 And LMode = False And Button = 1 Then
  LMode = True
  AddLine x, y
  
  ReDim Preserve PolyArea(0)
  PolyArea(0).x = x / 15
  PolyArea(0).y = y / 15
ElseIf EditMode = 2 And LMode = True And Button = 1 Then
  AddLine x, y
  
  ReDim Preserve PolyArea(ALine.UBound - 1)
  PolyArea(ALine.UBound - 1).x = x / 15
  PolyArea(ALine.UBound - 1).y = y / 15
ElseIf EditMode = 2 And LMode = True And Button = 2 Then
  LMode = False
  
  ReDim Preserve PolyArea(ALine.UBound - 1)
  PolyArea(ALine.UBound - 1).x = PolyArea(0).x
  PolyArea(ALine.UBound - 1).y = PolyArea(0).y
  
  '0 = Area Name
  '1 = Desc
  '2 = Prog
  '3 = X
  '4 = Y
  
  AddRGNPoly AllAreas, PolyArea, UBound(PolyArea), "Testz", "Dis a testz", vbGreen, vbRed, RGN_HS_DIAGCROSS, RGN_BS_HATCHED
  For a = 1 To ALine.UBound
    Unload ALine(a)
  Next a
  
  Dim dAr As String
  dAr = ""
  For a = LBound(PolyArea) To UBound(PolyArea)
    dAr = dAr + Str(PolyArea(a).x)
    If a < UBound(PolyArea) Then dAr = dAr + ","
  Next a
  DaRoom.Area(DaRoom.AreaCount, 3) = Trim$(dAr)
  
  dAr = ""
  For a = LBound(PolyArea) To UBound(PolyArea)
    dAr = dAr + Str(PolyArea(a).y)
    If a < UBound(PolyArea) Then dAr = dAr + ","
  Next a
  
  DaRoom.Area(DaRoom.AreaCount, 4) = Trim$(dAr)
  
  EditMode = 0
ElseIf EditMode = 0 And Button = 2 Then
  Dim i As Integer
  Dim LRGN As Long
  LRGN = IsInRegion(AllAreas, x, y)
  If LRGN <> 0 Then
    For i = 1 To AllAreas.Count
    If AllAreas(i).AreaNumber = LRGN Then
        uIndex = i
        PopupMenu mnuAreaContext
    End If
    Next i
  End If
  
End If
End Sub

Private Sub dRoom_Paint()
'dRoom.Cls
Dim i As Integer

For i = 1 To AllAreas.Count
If AllAreas(i).AreaState = 0 Then
    PaintARGN AllAreas.ParentHDC, AllAreas(i).AreaNumber, AllAreas(i).AreaPen, AllAreas(i).AreaBrush
Else
    PaintARGN AllAreas.ParentHDC, AllAreas(i).AreaNumber, AllAreas(i).AreaPen, AllAreas(i).AreaAlertBrush
End If
        
Next i
End Sub

Private Sub Form_Load()
RoomLst.Path = WrkDir

Set AllAreas = New Areas
AllAreas.ParentHDC = dRoom.Hdc

If DaRoom.RName <> "" Then LoadRoom DaRoom.RName: SetCapt DaRoom.RName

End Sub

Private Sub Form_Unload(Cancel As Integer)
AllAreas.ClearAll
End Sub

Private Sub mnuAddDg_Click()
uArea = ""
uDiag = ""
frmSetArea.Show vbModal

If uArea <> "" Then
  AddArea uArea, uDesc
End If
End Sub

Private Sub mnuAddItm_Click()
uDiag = ""
frmSetItm.Show vbModal
End Sub

Private Sub mnuAreaDelete_Click()

dcon = MsgBox("This feature is not working properly in this version. Are YOU sure you really want to delete the item '" & DaRoom.Area(uIndex, 0) & "'?" & uIndex, vbYesNo + vbExclamation, "Delete Item")

If dcon = vbNo Then Exit Sub

  ' Consider it deleted if the area name is blank
  DaRoom.Area(uIndex, 0) = ""
  
  If uIndex < 50 Then
    For a = uIndex To DaRoom.AreaCount
      VSwap DaRoom.Area(a, 0), DaRoom.Area(a + 1, 0)
      VSwap DaRoom.Area(a, 1), DaRoom.Area(a + 1, 1)
      VSwap DaRoom.Area(a, 2), DaRoom.Area(a + 1, 2)
      VSwap DaRoom.Area(a, 3), DaRoom.Area(a + 1, 3)
      VSwap DaRoom.Area(a, 4), DaRoom.Area(a + 1, 4)
    Next a
  End If
  
  ' Decrease the amount of the area
  DaRoom.AreaCount = DaRoom.AreaCount - 1
  
  ' Reload the room
  ReloadRoom
End Sub

Private Sub mnuAreaEdit_Click()
        uEdit = True
        uArea = DaRoom.Area(uIndex, 0)
        uDesc = DaRoom.Area(uIndex, 1)
        uDiag = DaRoom.Area(uIndex, 2)
        frmSetArea.Show vbModal
        
        If uArea <> "" Then
            DaRoom.Area(uIndex, 0) = uArea
            DaRoom.Area(uIndex, 1) = uDesc
            DaRoom.Area(uIndex, 2) = uDiag
        End If
End Sub

Private Sub mnuItmDelete_Click()

dcon = MsgBox("Are YOU sure you really want to delete the item '" & DaRoom.Itm(uIndex, 0) & "'?", vbYesNo + vbQuestion, "Delete Item")

If dcon = vbNo Then Exit Sub

  
  ' Consider it deleted if the item name is blank
  DaRoom.Itm(uIndex, 0) = ""
  
  If uIndex < 50 Then
    For a = uIndex To DaRoom.ItmCount
      VSwap DaRoom.Itm(a, 0), DaRoom.Itm(a + 1, 0)
      VSwap DaRoom.Itm(a, 1), DaRoom.Itm(a + 1, 1)
      VSwap DaRoom.Itm(a, 2), DaRoom.Itm(a + 1, 2)
      VSwap DaRoom.Itm(a, 3), DaRoom.Itm(a + 1, 3)
      VSwap DaRoom.Itm(a, 4), DaRoom.Itm(a + 1, 4)
      VSwap DaRoom.Itm(a, 5), DaRoom.Itm(a + 1, 5)
    Next a
  End If
  
  ' Decrease the amount of the item
  DaRoom.ItmCount = DaRoom.ItmCount - 1
  
  ' Reload the room
  ReloadRoom
End Sub

Private Sub mnuItmEdit_Click()
  uItm = DaRoom.Itm(uIndex, 0)
  uDesc = DaRoom.Itm(uIndex, 1)
  uDiag = DaRoom.Itm(uIndex, 2)
  uGfx = DaRoom.Itm(uIndex, 5)
  uEdit = True
  frmSetItm.Show vbModal
  
  uEdit = False
  If Trim$(uItm) <> "" Then
    DaRoom.Itm(uIndex, 0) = uItm
    DaRoom.Itm(uIndex, 1) = uDesc
    DaRoom.Itm(uIndex, 2) = uDiag
    DaRoom.Itm(uIndex, 5) = uGfx
  End If
End Sub

Private Sub mnuOpenRm_Click()
ShowRmList
End Sub

Private Sub mnuRmOpt_Click()
frmRoomOpt.Show vbModal
End Sub

Private Sub mnuSaveRm_Click()
tlSave_Click
End Sub

Private Sub mnuSetGfx_Click()
frmSetGraphic.Show vbModal

If uGfx <> "" Then
  dRoom.Picture = LoadPicture(WrkDir + uGfx)
  DaRoom.RBG = uGfx
End If
End Sub

Private Sub RoomLst_Click()
Dim dDesc As String

txtFName.Text = RoomLst.List(RoomLst.ListIndex)
If GetInitEntry("Room", "BG", , WrkDir + txtFName.Text) <> "" Then
  If Dir$(WrkDir + GetInitEntry("Room", "BG", , WrkDir + txtFName.Text)) <> "" Then imgPrev.Picture = LoadPicture(WrkDir + GetInitEntry("Room", "BG", , WrkDir + txtFName.Text))
  
  dDesc = GetInitEntry("Room", "Description", , WrkDir + txtFName.Text)
  If Len(dDesc) > 50 Then dDesc = Left$(dDesc, 50) + "..."
  PrevDesc.Caption = dDesc
End If
End Sub

Private Sub RoomLst_DblClick()
cmdAction_Click
End Sub

Private Sub tlNew_Click()
dRoom.Picture = LoadPicture()

' Clear all areas from previous room
AllAreas.ClearAll

' Clear all previous items
Do Until xItm.Count = 1
  Unload xItm(xItm.UBound)
Loop

SetCapt ""

With DaRoom
  .RName = ""
  .RBG = ""
  txtDesc.Text = ""
  .RSaved = False
  .RMusic = ""
  .dNorthWest = ""
  .dNorth = ""
  .dNorthEast = ""
  .dWest = ""
  .dEast = ""
  .dSouthWest = ""
  .dSouth = ""
  .dSouthEast = ""
  .RLocName = ""
  .ItmCount = 0
  .AreaCount = 0

  For a = 1 To 50
    .Itm(a, 0) = ""
    .Itm(a, 1) = ""
    .Itm(a, 2) = ""
    .Itm(a, 3) = ""
    .Itm(a, 4) = ""
  Next a
  
  For a = 1 To 50
    .Area(a, 0) = ""
    .Area(a, 1) = ""
    .Area(a, 2) = ""
    .Area(a, 3) = ""
    .Area(a, 4) = ""
  Next a
  
End With

dRoom.Refresh
End Sub

Private Sub tlOpen_Click()
ShowRmList
End Sub

Private Sub tlSave_Click()

If Trim$(DaRoom.RName) = "" Then
  ShowRmList True
Else
  SaveRoom DaRoom.RName
  PackGame GameName
End If
End Sub

Sub AddItm(dItm As String)

  LoadItmImg dItm
  
  DaRoom.ItmCount = DaRoom.ItmCount + 1
  
  EditMode = 1
  DaRoom.Itm(DaRoom.ItmCount, 0) = dItm
  DaRoom.Itm(DaRoom.ItmCount, 2) = uDiag
  
End Sub

Sub AddArea(dArea As String, dDesc As String)
  DaRoom.AreaCount = DaRoom.AreaCount + 1
  
  EditMode = 2
  DaRoom.Area(DaRoom.AreaCount, 0) = dArea
  DaRoom.Area(DaRoom.AreaCount, 1) = dDesc
  DaRoom.Area(DaRoom.AreaCount, 2) = uDiag
  
End Sub

Private Sub xItm_Click(Index As Integer)
If EditMode = 1 Then
  EditMode = 0
  
  '0 = Name
  '1 = Desc
  '2 = Prog
  '3 = X
  '4 = Y
  
  DaRoom.Itm(DaRoom.ItmCount, 3) = xItm(Index).Left
  DaRoom.Itm(DaRoom.ItmCount, 4) = xItm(Index).Top
End If
End Sub

Private Sub xItm_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If EditMode = 1 Then
  xItm(ItmNum).Left = (xItm(ItmNum).Left + x) - xItm(ItmNum).Width / 2
  xItm(ItmNum).Top = (xItm(ItmNum).Top + y) - xItm(ItmNum).Height / 2
End If
End Sub

Sub LoadItmImg(dItm As String, Optional wX As Long, Optional wY As Long)
  Dim dGfx As String
  
  Load xItm(xItm.Count)
  
  dGfx = GetInitEntry("Item", "Graphic", , WrkDir + dItm)
  If Dir$(WrkDir + dGfx) <> "" And Trim$(dGfx) <> "" Then xItm(xItm.UBound).Picture = LoadPicture(WrkDir + dGfx)
  
  xItm(xItm.UBound).Visible = True
  If wX <> 0 Then xItm(xItm.UBound).Left = wX
  If wY <> 0 Then xItm(xItm.UBound).Top = wY
  xItm(xItm.UBound).ToolTipText = Left$(dItm, Len(dItm) - 4) & ": " & GetInitEntry("Item", "Description", , WrkDir + dItm)
  
  ItmNum = xItm.UBound
End Sub

Sub AddLine(wX As Single, wY As Single)
  Load ALine(ALine.Count)
  ALine(ALine.UBound).Visible = True
  ALine(ALine.UBound).X1 = wX
  ALine(ALine.UBound).Y1 = wY

End Sub

Private Sub xItm_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
  uIndex = Index
  PopupMenu mnuItmContext
End If
End Sub

Sub ReloadRoom()
' Reloads the current room's items and areas
' You know, when we're deleting things and stuffs,
' the memory's going to be a little bit messy, so...

' If there's no open room, then there's nothing to reload,
' so if that's the case, then just exit peacefully.
With DaRoom

If .RName = "" Then Exit Sub

' Clear all areas from previous room
AllAreas.ClearAll

' Clear all previous items
Do Until xItm.Count = 1
  Unload xItm(xItm.UBound)
Loop

  ' Then reload thee items...
  For a = 1 To 50
    If .Itm(a, 0) <> "" Then LoadItmImg .Itm(a, 0), Val(.Itm(a, 3)), Val(.Itm(a, 4))
  Next a
  
  ' ...and then the areas!
  For a = 1 To 50
    If DaRoom.Area(a, 0) <> "" Then LoadRegion a
  Next a
 
End With

dRoom.Refresh
  
End Sub
