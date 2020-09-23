VERSION 5.00
Begin VB.Form frmConsole 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Quest Console"
   ClientHeight    =   3375
   ClientLeft      =   2985
   ClientTop       =   3960
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtConsole 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move Screen.Width - Me.Width - 100, Screen.Height - Me.Height - 500
End Sub

Private Sub Form_Resize()
txtConsole.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
