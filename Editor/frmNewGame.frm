VERSION 5.00
Begin VB.Form frmNewGame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create New Adventure"
   ClientHeight    =   4110
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   2280
      Width           =   4215
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   1560
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      Height          =   3885
      Left            =   120
      Picture         =   "frmNewGame.frx":0000
      ScaleHeight     =   3825
      ScaleWidth      =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1980
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Game Author(s):"
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Game Title:"
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label Label1 
      Caption         =   "You are about to create your own Adventure Game. Please fill up the information required."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "frmNewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCreate_Click()
Dim dFName As String
Dim PName As String

If GetFileName(dFName, "Adventure Game (*.adv)|*.adv", "Save Adventure", True) = True Then
  MakeTmpDir
  If LCase$(Right$(dFName, 4)) <> ".adv" Then dFName = dFName + ".adv"
  GameName = dFName
  
  DeleteTmp

  ' General Configuration File w/ filename "config.cfg"
  PName = WrkDir + "Config.cfg"
  SetInitEntry "General", "Title", txtTitle.Text, PName
  SetInitEntry "General", "Author", txtAuthor.Text, PName
  
  SetInitEntry "General", "VersionMajor", Str(VMajor), PName
  SetInitEntry "General", "VersionMinor", Str(VMinor), PName
  
  
  ' Pack 'em up!
  PackGame GameName, True
  
  frmMain.tDock.Visible = True
  
  frmMain.Caption = SetCaption(GetFName(GameName))

  MsgBox "You may now start creating your game." & _
         vbCrLf + vbCrLf & _
         "Click on the items on the left of the screen to edit different parts of the game.", vbInformation
         
End If
Unload Me
End Sub
