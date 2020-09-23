VERSION 5.00
Begin VB.Form frmSetArea 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Area"
   ClientHeight    =   2085
   ClientLeft      =   2820
   ClientTop       =   3960
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox txtAreaName 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdAdv 
      Caption         =   "Advanced >>"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox cmbCmbDiag 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3240
      Width           =   3135
   End
   Begin VB.ComboBox cmbCmbItm 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2520
      Width           =   3135
   End
   Begin VB.ComboBox cmbSDiag 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Text            =   "cmbSDiag"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Description of Room Area:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1860
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "...execute the following dialog:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   2145
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "When used with the following item...."
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   2595
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Dialog to run when area is activated:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   2595
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Area Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmSetArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
uArea = ""

uEdit = False
Unload Me
End Sub

Private Sub cmdAdv_Click()
If Right$(cmdAdv.Caption, 2) = ">>" Then
  cmdAdv.Caption = "Advanced <<"
  Me.Height = 4230
Else
  cmdAdv.Caption = "Advanced >>"
  Me.Height = 2565
End If
End Sub

Private Sub Form_Load()

If uEdit = False Then
  uArea = ""
  uDiag = ""
  uDesc = ""
End If

FillDialog cmbSDiag
FillItem cmbCmbItm
FillDialog cmbCmbDiag

If uEdit = True Then
  txtAreaName.Text = uArea
  txtDesc.Text = uDesc
  cmbSDiag.Text = uDiag
End If

End Sub

Private Sub OKButton_Click()

uDiag = cmbSDiag.Text

If txtAreaName.Text = "" Then MsgBox "Please enter a name for the room area!", vbExclamation, "Add Area": Exit Sub


uArea = txtAreaName.Text
uDesc = txtDesc.Text

uEdit = False

Unload Me

End Sub
