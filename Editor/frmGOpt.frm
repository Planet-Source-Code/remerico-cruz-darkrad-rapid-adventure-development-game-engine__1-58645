VERSION 5.00
Begin VB.Form frmGOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game Settings"
   ClientHeight    =   3360
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7335
   Icon            =   "frmGOpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.PictureBox dTab 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2415
      Index           =   0
      Left            =   120
      ScaleHeight     =   2385
      ScaleWidth      =   7065
      TabIndex        =   2
      Top             =   360
      Width           =   7095
      Begin VB.CheckBox chkDBug 
         Caption         =   "Debug Mode"
         Height          =   255
         Left            =   5520
         TabIndex        =   26
         ToolTipText     =   "When activated, it will display the console window to help you in debugging your game"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtSaveExt 
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   19
         ToolTipText     =   "Lets you modify the file extension of the save game (default: .sav)"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtAuthor 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox txtGTitle 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   5775
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6840
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Save Game Extension:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1220
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Author(s):"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   630
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Game Title:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   285
         Width           =   810
      End
   End
   Begin VB.PictureBox dTab 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2415
      Index           =   2
      Left            =   120
      ScaleHeight     =   2385
      ScaleWidth      =   7065
      TabIndex        =   11
      Top             =   360
      Width           =   7095
      Begin VB.TextBox txtKRun 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   30
         Top             =   1850
         Width           =   975
      End
      Begin VB.ComboBox cmbGOverDg 
         Height          =   315
         Left            =   2280
         TabIndex        =   24
         Top             =   960
         Width           =   3735
      End
      Begin VB.ComboBox cmbRunDg 
         Height          =   315
         Left            =   2280
         TabIndex        =   22
         Top             =   600
         Width           =   3735
      End
      Begin VB.ComboBox cmbInitDg 
         Height          =   315
         Left            =   2280
         TabIndex        =   20
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Runtime Activation Key:"
         Height          =   195
         Left            =   360
         TabIndex        =   29
         Top             =   1880
         Width           =   1695
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   " [Activation Keys] "
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Shape Shape2 
         Height          =   615
         Left            =   120
         Top             =   1680
         Width           =   6855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   " [Programs] "
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   0
         Width           =   840
      End
      Begin VB.Shape Shape1 
         Height          =   1335
         Left            =   120
         Top             =   120
         Width           =   6855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Game Over Dialog: "
         Height          =   195
         Left            =   840
         TabIndex        =   25
         Top             =   1005
         Width           =   1395
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Runtime Dialog: "
         Height          =   195
         Left            =   960
         TabIndex        =   23
         Top             =   645
         Width           =   1170
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Startup Dialog: "
         Height          =   195
         Left            =   1080
         TabIndex        =   21
         Top             =   285
         Width           =   1095
      End
   End
   Begin VB.PictureBox dTab 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2415
      Index           =   1
      Left            =   120
      ScaleHeight     =   2385
      ScaleWidth      =   7065
      TabIndex        =   7
      Top             =   360
      Width           =   7095
      Begin VB.ComboBox cmbInitRoom 
         Height          =   315
         Left            =   2280
         TabIndex        =   17
         Top             =   840
         Width           =   3855
      End
      Begin VB.ComboBox cmbTitleScreen 
         Height          =   315
         Left            =   2280
         TabIndex        =   15
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Initial Room: "
         Height          =   195
         Left            =   1290
         TabIndex        =   14
         Top             =   870
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Game Title Screen:"
         Height          =   195
         Left            =   810
         TabIndex        =   13
         Top             =   510
         Width           =   1365
      End
   End
   Begin VB.PictureBox dTab 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2415
      Index           =   3
      Left            =   120
      ScaleHeight     =   2385
      ScaleWidth      =   7065
      TabIndex        =   12
      Top             =   360
      Width           =   7095
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   481
      X2              =   481
      Y1              =   27
      Y2              =   185
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   11
      X2              =   482
      Y1              =   185
      Y2              =   185
   End
   Begin VB.Label xTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Graphics"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label xTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dialog Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label xTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Startup Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label xTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Project Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmGOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()

FillGfx cmbTitleScreen
FillRoom cmbInitRoom
FillDialog cmbInitDg
FillDialog cmbRunDg
FillDialog cmbGOverDg

GetGameSettings
End Sub

Private Sub OKButton_Click()
SetGameSettings
PackGame GameName
Unload Me
End Sub

Private Sub txtKRun_KeyUp(KeyCode As Integer, Shift As Integer)
txtKRun.Text = GetChar(KeyCode)
End Sub

Private Sub xTab_Click(Index As Integer)
dTab(Index).ZOrder

xTab(Index).BackColor = vbButtonShadow

For a = xTab.LBound To xTab.UBound
  If a <> Index Then xTab(a).BackColor = vbButtonFace
Next a

End Sub

Sub GetGameSettings()
txtGTitle.Text = GetInitEntry("General", "Title", , WrkDir + "Config.cfg")
txtAuthor.Text = GetInitEntry("General", "Author", , WrkDir + "Config.cfg")

txtSaveExt.Text = GetInitEntry("General", "SaveExt", "sav", WrkDir + "Config.cfg")
chkDBug.Value = GetInitEntry("General", "Debug", vbUnchecked, WrkDir + "Config.cfg")

cmbTitleScreen.Text = GetInitEntry("Init", "TitleScreen", , WrkDir + "Config.cfg")
cmbInitRoom.Text = GetInitEntry("Init", "InitRoom", , WrkDir + "Config.cfg")

cmbInitDg.Text = GetInitEntry("Dialog", "Startup", , WrkDir + "Config.cfg")
cmbRunDg.Text = GetInitEntry("Dialog", "Runtime", , WrkDir + "Config.cfg")
cmbGOverDg.Text = GetInitEntry("Dialog", "GameOver", , WrkDir + "Config.cfg")

End Sub

Sub SetGameSettings()
SetInitEntry "General", "Title", txtGTitle.Text, WrkDir + "Config.cfg"
SetInitEntry "General", "Author", txtAuthor.Text, WrkDir + "Config.cfg"

SetInitEntry "General", "VersionMajor", Str(VMajor), WrkDir + "Config.cfg"
SetInitEntry "General", "VersionMinor", Str(VMinor), WrkDir + "Config.cfg"

SetInitEntry "General", "SaveExt", txtSaveExt.Text, WrkDir + "Config.cfg"
SetInitEntry "General", "Debug", chkDBug.Value, WrkDir + "Config.cfg"

SetInitEntry "Init", "TitleScreen", cmbTitleScreen.Text, WrkDir + "Config.cfg"
SetInitEntry "Init", "InitRoom", cmbInitRoom.Text, WrkDir + "Config.cfg"

SetInitEntry "Dialog", "Startup", cmbInitDg.Text, WrkDir + "Config.cfg"
SetInitEntry "Dialog", "Runtime", cmbRunDg.Text, WrkDir + "Config.cfg"
SetInitEntry "Dialog", "GameOver", cmbGOverDg.Text, WrkDir + "Config.cfg"
End Sub
