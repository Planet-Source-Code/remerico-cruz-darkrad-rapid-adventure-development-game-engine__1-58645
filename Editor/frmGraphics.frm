VERSION 5.00
Begin VB.Form frmResource 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resource Editor"
   ClientHeight    =   5010
   ClientLeft      =   225
   ClientTop       =   960
   ClientWidth     =   7275
   Icon            =   "frmGraphics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   4560
      Width           =   975
   End
   Begin VB.Frame PicPrev 
      Caption         =   "Preview"
      Height          =   2775
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   3375
      Begin VB.Image dImg 
         Height          =   2415
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.ListBox lstRcCat 
      Height          =   1620
      ItemData        =   "frmGraphics.frx":038A
      Left            =   0
      List            =   "frmGraphics.frx":039A
      TabIndex        =   4
      Top             =   0
      Width           =   3375
   End
   Begin VB.PictureBox rTab 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   3480
      ScaleHeight     =   4575
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmdRmvRc 
         Caption         =   "Remove Resource"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   0
         Width           =   1815
      End
      Begin VB.FileListBox RcLst 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4125
         Left            =   0
         Pattern         =   "."
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton cmdAddRc 
         Caption         =   "Add Resource"
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.Frame MscPrev 
      Caption         =   "Music Preview"
      Height          =   2775
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   3375
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label cMSt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[Stopped]"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   690
      End
      Begin VB.Label cMsc 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   3135
      End
   End
   Begin VB.Menu mnuRsc 
      Caption         =   "&Resource Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuRnRc 
         Caption         =   "&Rename Resource..."
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrevFile As String

Private Sub cmdAddRc_Click()
Dim dFName As String

If GetFileName(dFName, "Resource Files|*.bmp;*.gif;*.jpg;*.wmf;*.ico;*.wav;*.mp3;*.midi;*.ico;*.cur;*.mid;*.mp3;*.wav", "Select Resource File") = True Then
  FileCopy dFName, WrkDir + GetFName(dFName)
  RcLst.Refresh
End If
End Sub

Private Sub cmdClose_Click()
StopMIDI
Unload Me
End Sub

Private Sub cmdPlay_Click()
  If lstRcCat.List(lstRcCat.ListIndex) <> "Music" Or PrevFile = "" Then Exit Sub

  Select Case LCase$(Right$(PrevFile, 3))
    Case "mid"
      PlayMIDI PrevFile
  End Select
  
  cMsc.Caption = RcLst.List(RcLst.ListIndex)
  cMSt.Caption = "[Playing]"
End Sub

Private Sub cmdRmvRc_Click()
If RcLst.ListIndex < 0 Then Exit Sub
Kill WrkDir + RcLst.List(RcLst.ListIndex)
dImg.Picture = LoadPicture()
RcLst.Refresh
End Sub

Private Sub cmdSave_Click()
PackGame GameName
End Sub

Private Sub cmdStop_Click()
  If lstRcCat.List(lstRcCat.ListIndex) <> "Music" Or PrevFile = "" Then Exit Sub

  Select Case LCase$(Right$(PrevFile, 3))
    Case "mid"
      StopMIDI
  End Select
  
  cMSt.Caption = "[Stopped]"
  cMsc.Caption = ""
End Sub

Private Sub Form_Load()
RcLst.Path = WrkDir
End Sub

Private Sub lstRcCat_Click()
dImg.Picture = LoadPicture()
Select Case lstRcCat.List(lstRcCat.ListIndex)
  Case "Graphics"
    RcLst.Pattern = "*.bmp;*.gif;*.jpg;*.wmf"
  Case "Music"
    RcLst.Pattern = "*.wav;*.mp3;*.mid"
  Case "Icon"
    RcLst.Pattern = "*.ico"
  Case "Cursor"
    RcLst.Pattern = "*.cur"
End Select
End Sub

Private Sub mnuRnRc_Click()
Dim NName As String
Dim OName As String

OName = RcLst.List(RcLst.ListIndex)
NName = InputBox("Enter new name for " + RcLst.List(RcLst.ListIndex), "Rename Resource", OName)

If NName <> "" And OName <> NName Then
  FileCopy WrkDir + OName, WrkDir + NName
  Kill WrkDir + OName
End If

End Sub

Private Sub RcLst_Click()
Select Case lstRcCat.List(lstRcCat.ListIndex)
  Case "Graphics", "Icon", "Cursor"
    PicPrev.ZOrder
    dImg.Picture = LoadPicture(WrkDir + RcLst.List(RcLst.ListIndex))
  Case "Music"
    MscPrev.ZOrder
    PrevFile = WrkDir + RcLst.List(RcLst.ListIndex)
End Select

End Sub

Private Sub RcLst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And RcLst.ListIndex > -1 Then PopupMenu mnuRsc
End Sub
