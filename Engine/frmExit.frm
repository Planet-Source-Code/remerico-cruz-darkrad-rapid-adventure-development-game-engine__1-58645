VERSION 5.00
Begin VB.Form frmExit 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1965
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label PYes 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Very Sure!"
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
      Height          =   240
      Left            =   3840
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label PNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Not sure"
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
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label dPrompt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "How sure are you that you want to exit this very exciting game?"
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
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
   Begin VB.Shape shape 
      BorderColor     =   &H00FFC0C0&
      Height          =   1695
      Left            =   120
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 27      ' [ESC]
    Unload Me
  Case 37, 39  ' [Left], [Right]
    If PNo.BackStyle = 1 Then
      PNo.BackStyle = 0: PYes.BackStyle = 1
    ElseIf PYes.BackStyle = 1 Then
      PYes.BackStyle = 0: PNo.BackStyle = 1
    End If
  Case 13      ' [ENTER]
    If PNo.BackStyle = 1 Then
      Unload Me
    Else
      Unload Me
      Unload frmMain
      DeleteTmp
      'ChangeRes 800, 600
      End
    End If
End Select
End Sub

Private Sub Form_Load()
ApplyTheme "exit"
End Sub

Private Sub PNo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
PYes.BackStyle = 0
PNo.BackStyle = 1
End Sub

Private Sub PNo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Form_KeyDown 13, 0
End Sub

Private Sub PYes_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
PYes.BackStyle = 1
PNo.BackStyle = 0
End Sub

Private Sub PYes_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Form_KeyDown 13, 0
End Sub
