VERSION 5.00
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Dark Adventure Editor"
   ClientHeight    =   5340
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6600
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox tStat 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   6600
      TabIndex        =   9
      Top             =   4860
      Width           =   6600
      Begin VB.Frame tTitle 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   545
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   6615
         Begin VB.Label lblInst 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   11
            Top             =   210
            Width           =   60
         End
      End
   End
   Begin VB.PictureBox tDock 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4485
      Left            =   0
      ScaleHeight     =   299
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   1
      Top             =   375
      Visible         =   0   'False
      Width           =   1815
      Begin DarkEdit.GurhanButton cmdCmp 
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   12
         Top             =   2250
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Appearance Editor"
         ButtonStyle     =   1
         Picture         =   "frmMain.frx":0000
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
      Begin DarkEdit.GurhanButton cmdCmp 
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   8
         Top             =   1875
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Character Editor"
         ButtonStyle     =   1
         Picture         =   "frmMain.frx":0352
         PictureWidth    =   16
         PictureHeight   =   16
         PictureSize     =   0
         OriginalPicSizeW=   15
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
      Begin DarkEdit.GurhanButton cmdCmp 
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   7
         Top             =   1500
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Item Editor"
         ButtonStyle     =   1
         Picture         =   "frmMain.frx":06A4
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
      Begin DarkEdit.GurhanButton cmdCmp 
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   1125
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Dialog Editor"
         ButtonStyle     =   1
         Picture         =   "frmMain.frx":09F6
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
      Begin DarkEdit.GurhanButton cmdCmp 
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   750
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Room Editor"
         ButtonStyle     =   1
         Picture         =   "frmMain.frx":0D48
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
      Begin VB.ListBox LstFiles 
         Height          =   1815
         Left            =   0
         TabIndex        =   2
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin DarkEdit.GurhanButton cmdCmp 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Game Settings"
         ButtonStyle     =   1
         Picture         =   "frmMain.frx":109A
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
      Begin DarkEdit.GurhanButton cmdCmp 
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   375
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Resource Editor"
         ButtonStyle     =   1
         Picture         =   "frmMain.frx":13EC
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
   Begin VB.PictureBox tool 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6600
      TabIndex        =   0
      Top             =   0
      Width           =   6600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewAdv 
         Caption         =   "&New Adventure..."
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Adventure Game"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &as..."
      End
      Begin VB.Menu mnus1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Dark Adventure Toolkit..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCmp_Click(Index As Integer)
Select Case Index
  Case 0
    frmGOpt.Show 'vbModal
  Case 1
    frmResource.Show 'vbModal
  Case 2
    frmRoom.Show 'vbModal
  Case 3
    frmDialog.Show 'vbModal
  Case 4
    frmItem.Show
  Case 5
    MsgBox "Not done yet!", vbInformation, "Sowee"
  Case 6
    frmAppear.Show
End Select
End Sub

Private Sub cmdCmp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
  Case 0
    DTip "Game Settings", "Allows you to customize your game by setting the options you want."
  Case 1
    DTip "Resource Editor", "Stores the resources for your game, such as the pictures and music that you will use in your game."
  Case 2
    DTip "Room Editor", "Allows you to edit the rooms in your game. Rooms are the places where your player will navigate in your game."
  Case 3
    DTip "Dialog Editor", "Allows you to edit the conversations and dialogs in your game. You can also control the flow of your game here."
  Case 4
    DTip "Item Editor", "Allows you to edit the inventory items in your game. Your player will use these to solve puzzles, etc."
  Case 5
    DTip "Character Editor", "Allows you to edit the characters and other people in your game."
  Case 6
    DTip "Appearance Editor", "Allows you to change visual settings for your game."
End Select
End Sub

Private Sub MDIForm_Load()

WrkDir = App.Path + "\WorkingTmp\"
VMajor = 1
VMinor = 0

MakeTmpDir

LoadSetting

If GameName <> "" Then
  tDock.Visible = True
  ExtractGame GameName
  Me.Caption = SetCaption(GetFName(GameName))
End If

InitSound


Unload frmSplash

End Sub

Private Sub MDIForm_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
tTitle.Width = Me.ScaleWidth + tDock.Width

'tTitle.Width = tStat.Width

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
SaveSetting
DeleteTmp
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuNewAdv_Click()
frmNewGame.Show vbModal
End Sub

Private Sub mnuOpen_Click()
Dim DgFile As String

If GetFileName(DgFile, "Adventure Toolkit Game file|*.adv", "Open Adventure Game") = True Then
  DeleteTmp
  tDock.Visible = True
  GameName = DgFile
  ExtractGame GameName
  Me.Caption = SetCaption(GetFName(GameName))
End If
End Sub

Private Sub mnuSave_Click()
If GameName <> "" Then PackGame GameName
End Sub

Private Sub mnuSaveAs_Click()
Dim GSaveas As String

If GetFileName(GSaveas, "Adventure Toolkit Game file|*.adv", "Save Adventure Game As...", True) Then
  GSaveas = GSaveas + ".adv"
  PackGame GSaveas, True
End If
End Sub

Private Sub tDock_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DTip "", ""
End Sub
