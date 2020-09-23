VERSION 5.00
Begin VB.Form frmSetGraphic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Graphic"
   ClientHeight    =   4350
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3255
      Begin VB.Image dImg 
         Height          =   2535
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.ComboBox cmbGfx 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select the Graphic that you want to use:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2925
   End
End
Attribute VB_Name = "frmSetGraphic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub cmbGfx_Click()
dImg.Picture = LoadPicture(WrkDir + cmbGfx.Text)
End Sub

Private Sub Form_Load()
uGfx = ""
FillGfx cmbGfx
End Sub

Private Sub OKButton_Click()
uGfx = cmbGfx.Text
Unload Me
End Sub
