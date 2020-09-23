VERSION 5.00
Begin VB.Form frmChoice 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   615
   ClientTop       =   1200
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChoice 
      Height          =   450
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label dChoice 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
   Begin VB.Shape shape 
      BorderColor     =   &H00FFC0C0&
      Height          =   1575
      Left            =   120
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CSlot As Integer
Dim hChoose As Boolean

Private Sub dChoice_Click(Index As Integer)
dChoice(CSlot).BackStyle = 0
CSlot = Index
dChoice(CSlot).BackStyle = 1
End Sub

Private Sub dChoice_DblClick(Index As Integer)
Form_KeyDown 13, 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
  Case 38   ' UP
    dChoice(CSlot).BackStyle = 0
    CSlot = CSlot - 1
    If CSlot = -1 Then CSlot = dChoice.UBound
    dChoice(CSlot).BackStyle = 1
  Case 40   ' DOWN
    dChoice(CSlot).BackStyle = 0
    CSlot = CSlot + 1
    If CSlot = dChoice.Count Then CSlot = 0
    dChoice(CSlot).BackStyle = 1
  Case 13  ' ENTER
    hChoose = True
    Echo "User selected choice '" & dChoice(CSlot) & "'"
    GoToLine lstChoice.List(CSlot)
    frmMain.dScreen.SetFocus
    GScreen = "game"
    Unload Me
End Select
End Sub

Sub xLoop()
hChoose = False
Do Until hChoose = True
  DoEvents
Loop
End Sub

Private Sub Form_Load()
GScreen = "conv"
CSlot = 0
End Sub
