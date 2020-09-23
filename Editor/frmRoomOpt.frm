VERSION 5.00
Begin VB.Form frmRoomOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Room Options"
   ClientHeight    =   3405
   ClientLeft      =   3120
   ClientTop       =   3795
   ClientWidth     =   7350
   Icon            =   "frmRoomOpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   227
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Default         =   -1  'True
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
      TabIndex        =   10
      Top             =   360
      Width           =   7095
      Begin VB.ComboBox cmbTrans 
         Height          =   315
         ItemData        =   "frmRoomOpt.frx":038A
         Left            =   1800
         List            =   "frmRoomOpt.frx":03A6
         TabIndex        =   19
         Text            =   "cmbTrans"
         Top             =   1300
         Width           =   5175
      End
      Begin VB.ComboBox cmbRMusic 
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Top             =   800
         Width           =   5175
      End
      Begin VB.TextBox txtLocName 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Room Transition:"
         Height          =   195
         Left            =   540
         TabIndex        =   18
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Background Music:"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Room/Location Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   1620
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
      TabIndex        =   3
      Top             =   360
      Width           =   7095
      Begin VB.Frame dRm 
         Appearance      =   0  'Flat
         Caption         =   " Room"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   5295
         Begin VB.ComboBox cmbDRm 
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   4815
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1080
         Left            =   360
         ScaleHeight     =   70
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   70
         TabIndex        =   6
         Top             =   600
         Width           =   1080
         Begin VB.Label lblCD 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   360
            TabIndex        =   9
            Top             =   360
            Width           =   375
         End
         Begin VB.Image dDir 
            Height          =   360
            Index           =   0
            Left            =   0
            Picture         =   "frmRoomOpt.frx":0413
            Top             =   0
            Width           =   360
         End
         Begin VB.Image dDir 
            Height          =   360
            Index           =   1
            Left            =   360
            Picture         =   "frmRoomOpt.frx":04BD
            Top             =   0
            Width           =   360
         End
         Begin VB.Image dDir 
            Height          =   360
            Index           =   2
            Left            =   720
            Picture         =   "frmRoomOpt.frx":0567
            Top             =   0
            Width           =   360
         End
         Begin VB.Image dDir 
            Height          =   360
            Index           =   3
            Left            =   0
            Picture         =   "frmRoomOpt.frx":0611
            Top             =   360
            Width           =   360
         End
         Begin VB.Image dDir 
            Height          =   360
            Index           =   5
            Left            =   720
            Picture         =   "frmRoomOpt.frx":06BB
            Top             =   360
            Width           =   360
         End
         Begin VB.Image dDir 
            Height          =   360
            Index           =   6
            Left            =   0
            Picture         =   "frmRoomOpt.frx":0765
            Top             =   720
            Width           =   360
         End
         Begin VB.Image dDir 
            Height          =   360
            Index           =   7
            Left            =   360
            Picture         =   "frmRoomOpt.frx":080F
            Top             =   720
            Width           =   360
         End
         Begin VB.Image dDir 
            Height          =   360
            Index           =   8
            Left            =   720
            Picture         =   "frmRoomOpt.frx":08B9
            Top             =   720
            Width           =   360
         End
      End
      Begin VB.OptionButton optRmK 
         Caption         =   "Dialog Script"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optRmK 
         Caption         =   "Room"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   16
         Top             =   1440
         Value           =   -1  'True
         Width           =   735
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
      TabIndex        =   5
      Top             =   360
      Width           =   7095
      Begin VB.CheckBox chkDToolTip 
         Caption         =   "Disable Item and Tooltip in Room"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   960
         Width           =   2775
      End
      Begin VB.CheckBox chkDSave 
         Caption         =   "Disable Save Game in this Room"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox chkDNav 
         Caption         =   "Don't Show Navigation Bar in this Room"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   3375
      End
      Begin VB.CheckBox chkDMenu 
         Caption         =   "Disable Menu in this Room"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Label xTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "General Settings"
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
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label xTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Misc. Options"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   10
      X2              =   481
      Y1              =   185
      Y2              =   185
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   481
      X2              =   481
      Y1              =   26
      Y2              =   186
   End
   Begin VB.Label xTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Directional Links"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmRoomOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CDrm

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub cmbDRm_Change()
cmbDRm_Click
End Sub

Private Sub cmbDRm_Click()
If CDrm = "" Then Exit Sub

With DaRoom
Select Case CDrm
  Case "NW"
    .dNorthWest = cmbDRm.Text
  Case "N"
    .dNorth = cmbDRm.Text
  Case "NE"
    .dNorthEast = cmbDRm.Text
  Case "W"
    .dWest = cmbDRm.Text
  Case "E"
    .dEast = cmbDRm.Text
  Case "SW"
    .dSouthWest = cmbDRm.Text
  Case "S"
    .dSouth = cmbDRm.Text
  Case "SE"
    .dSouthEast = cmbDRm.Text
End Select
End With
End Sub

Private Sub dDir_Click(Index As Integer)
With dRm
Select Case Index
  Case 0  ' NW
    .Caption = "Northwest Room"
    CDrm = "NW"
    
    If LCase$(Right$(DaRoom.dNorthWest, 2)) = "rm" Then
      optRmK(0).Value = True
    ElseIf LCase$(Right$(DaRoom.dNorthWest, 2)) = "dg" Then
      optRmK(1).Value = True
    End If
    
    cmbDRm.Text = DaRoom.dNorthWest
  Case 1  ' N
    .Caption = "North Room"
    CDrm = "N"
    
    If LCase$(Right$(DaRoom.dNorth, 2)) = "rm" Then
      optRmK(0).Value = True
    ElseIf LCase$(Right$(DaRoom.dNorth, 2)) = "dg" Then
      optRmK(1).Value = True
    End If
    
    cmbDRm.Text = DaRoom.dNorth
  Case 2  ' NE
    .Caption = "Northeast Room"
    CDrm = "NE"
    
    If LCase$(Right$(DaRoom.dNorthEast, 2)) = "rm" Then
      optRmK(0).Value = True
    ElseIf LCase$(Right$(DaRoom.dNorthEast, 2)) = "dg" Then
      optRmK(1).Value = True
    End If
    
    cmbDRm.Text = DaRoom.dNorthEast
  Case 3  ' W
    .Caption = "West Room"
    CDrm = "W"
    
    If LCase$(Right$(DaRoom.dWest, 2)) = "rm" Then
      optRmK(0).Value = True
    ElseIf LCase$(Right$(DaRoom.dWest, 2)) = "dg" Then
      optRmK(1).Value = True
    End If
    
    cmbDRm.Text = DaRoom.dWest
  Case 5  ' E
    .Caption = "East Room"
    CDrm = "E"
    
    If LCase$(Right$(DaRoom.dEast, 2)) = "rm" Then
      optRmK(0).Value = True
    ElseIf LCase$(Right$(DaRoom.dEast, 2)) = "dg" Then
      optRmK(1).Value = True
    End If
    
    cmbDRm.Text = DaRoom.dEast
  Case 6  ' SW
    .Caption = "Southwest Room"
    CDrm = "SW"
    
    If LCase$(Right$(DaRoom.dSouthWest, 2)) = "rm" Then
      optRmK(0).Value = True
    ElseIf LCase$(Right$(DaRoom.dSouthWest, 2)) = "dg" Then
      optRmK(1).Value = True
    End If
    
    cmbDRm.Text = DaRoom.dSouthWest
  Case 7  ' S
    .Caption = "South Room"
    CDrm = "S"
    
    If LCase$(Right$(DaRoom.dSouth, 2)) = "rm" Then
      optRmK(0).Value = True
    ElseIf LCase$(Right$(DaRoom.dSouth, 2)) = "dg" Then
      optRmK(1).Value = True
    End If
    
    cmbDRm.Text = DaRoom.dSouth
  Case 8  ' SE
    .Caption = "Southeast Room"
    CDrm = "SE"
    
    If LCase$(Right$(DaRoom.dSouthEast, 2)) = "rm" Then
      optRmK(0).Value = True
    ElseIf LCase$(Right$(DaRoom.dSouthEast, 2)) = "dg" Then
      optRmK(1).Value = True
    End If
    
    cmbDRm.Text = DaRoom.dSouthEast
End Select
End With

lblCD.Caption = CDrm
End Sub

Private Sub dDir_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
dDir(Index).Move dDir(Index).Left + 1, dDir(Index).Top + 1
End Sub

Private Sub dDir_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
dDir(Index).Move dDir(Index).Left - 1, dDir(Index).Top - 1
End Sub

Private Sub Form_Load()
FillRoom cmbDRm
FillMusic cmbRMusic

txtLocName.Text = DaRoom.RLocName
cmbRMusic.Text = DaRoom.RMusic

chkDSave.Value = DaRoom.DontSave
chkDNav.Value = DaRoom.DontShowNav
chkDMenu.Value = DaRoom.DisableMenu
chkDToolTip.Value = DaRoom.DisableTooltip

Select Case DaRoom.Trans
  Case 0
    cmbTrans.Text = "0 - None"
  Case 1
    cmbTrans.Text = "1 - Fade"
  Case 2
    cmbTrans.Text = "2 - Wipe"
  Case 3
    cmbTrans.Text = "3 - Hour (Double)"
  Case 4
    cmbTrans.Text = "4 - Hour (Inverse)"
  Case 5
    cmbTrans.Text = "5 - Circle"
  Case 6
    cmbTrans.Text = "6 - Implode"
  Case 7
    cmbTrans.Text = "7 - Tenda"
End Select

End Sub

Private Sub OKButton_Click()

DaRoom.RLocName = txtLocName.Text
DaRoom.RMusic = cmbRMusic.Text
DaRoom.Trans = Val(Left$(cmbTrans.Text, 1))
DaRoom.DontSave = chkDSave.Value
DaRoom.DontShowNav = chkDNav.Value
DaRoom.DisableMenu = chkDMenu.Value
DaRoom.DisableTooltip = chkDToolTip.Value

If DaRoom.RName <> "" Then frmRoom.SaveRoom DaRoom.RName
'PackGame GameName

'MsgBox DaRoom.Trans

Unload Me
End Sub

Private Sub optRmK_Click(Index As Integer)
Select Case Index
  Case 0
   dRm.Caption = "Room"
   FillRoom cmbDRm
  Case 1
   dRm.Caption = "Dialog"
   FillDialog cmbDRm
End Select
End Sub

Private Sub xTab_Click(Index As Integer)
dTab(Index).ZOrder

xTab(Index).BackColor = vbButtonShadow

For a = xTab.LBound To xTab.UBound
  If a <> Index Then xTab(a).BackColor = vbButtonFace
Next a
End Sub
