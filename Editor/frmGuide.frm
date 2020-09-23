VERSION 5.00
Begin VB.Form frmGuide 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3480
   ClientLeft      =   990
   ClientTop       =   1005
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.Frame tGuide 
      Caption         =   "Add Room Area"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Index           =   2
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   4455
      Begin VB.Label Label7 
         Caption         =   $"frmGuide.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Label6 
         Caption         =   "A Room Area is a portion of the room background that the user can interact with."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label5 
         Caption         =   "What are Room Areas?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.PictureBox cDock 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   4575
      TabIndex        =   3
      Top             =   3090
      Width           =   4575
      Begin VB.CheckBox chkDShw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         Caption         =   "Don't Show these tips again"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   2415
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Default         =   -1  'True
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame tGuide 
      Caption         =   "Open Room"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   4455
      Begin VB.Label Label4 
         Caption         =   "In this window, you can open existing rooms that you have created before."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "You can see a preview of the room you've selected on the right side of the window."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   4095
      End
   End
   Begin VB.Frame tGuide 
      Caption         =   "Welcome to DarkQuest!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Label Label2 
         Caption         =   "If you haven't done so, Please click ""File"" on the menu, and select ""New Adventure"". You will presented with a new window."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "You are about to create your own adventure game shortly."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmGuide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fx, Fy

Private Sub cDock_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Fx = x
Fy = y
End Sub

Private Sub cDock_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    Me.Left = (Me.Left + x) - Fx
    Me.Top = (Me.Top + y) - Fy
  End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub
