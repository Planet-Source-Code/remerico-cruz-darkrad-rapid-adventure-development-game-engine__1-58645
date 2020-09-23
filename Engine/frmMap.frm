VERSION 5.00
Begin VB.Form frmMap 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3840
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox dMap 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3585
      ScaleWidth      =   5745
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.Label dPrompt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Kunwari may mapa d2...."
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
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dMap_KeyDown(KeyCode As Integer, Shift As Integer)
Unload Me
End Sub

Private Sub dPrompt_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Unload Me
End Sub
