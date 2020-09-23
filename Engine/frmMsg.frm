VERSION 5.00
Begin VB.Form frmMsg 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3150
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label dMsg 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMsg.frx":0000
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
   Begin VB.Shape shape 
      BorderColor     =   &H00FFC0C0&
      Height          =   2895
      Left            =   120
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub dMsg_Click()
uOK = True
frmMain.dScreen.SetFocus
GScreen = "game"
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
uOK = True
frmMain.dScreen.SetFocus
GScreen = "game"
Unload Me
End Sub

Private Sub Form_Load()
GScreen = "conv"
End Sub

Private Sub Form_LostFocus()
uOK = True
frmMain.dScreen.SetFocus
GScreen = "game"
Unload Me
End Sub
