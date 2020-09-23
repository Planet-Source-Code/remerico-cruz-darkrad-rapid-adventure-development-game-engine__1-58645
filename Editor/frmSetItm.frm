VERSION 5.00
Begin VB.Form frmSetItm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Item"
   ClientHeight    =   2625
   ClientLeft      =   3120
   ClientTop       =   5010
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin DarkEdit.GurhanButton cmdItmGfx 
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   1560
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      Caption         =   "..."
      ButtonStyle     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SoundOver       =   ""
      SoundClick      =   ""
   End
   Begin VB.TextBox txtItmGfx 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtItmDesc 
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton cmdAdv 
      Caption         =   "Advanced >>"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox cmbCmbDiag 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3840
      Width           =   3135
   End
   Begin VB.ComboBox cmbCmbItm 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3120
      Width           =   3135
   End
   Begin VB.ComboBox cmbSItm 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Text            =   "cmbSItm"
      Top             =   360
      Width           =   3135
   End
   Begin VB.ComboBox cmbSDiag 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Text            =   "cmbSDiag"
      Top             =   2160
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
   Begin DarkEdit.GurhanButton cmdGfxClear 
      Height          =   255
      Left            =   3135
      TabIndex        =   16
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "X"
      ButtonStyle     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SoundOver       =   ""
      SoundClick      =   ""
   End
   Begin VB.Label Label6 
      Caption         =   "Room Item Graphic:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Room Item Description:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "...execute the following dialog:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   2145
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "When used with the following item...."
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   2595
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Dialog to run when item is activated:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2565
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Item:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmSetItm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
uItm = ""
uGfx = ""
uDesc = ""
Unload Me
End Sub

Private Sub cmdAdv_Click()
If Right$(cmdAdv.Caption, 2) = ">>" Then
  cmdAdv.Caption = "Advanced <<"
  Me.Height = 4770
Else
  cmdAdv.Caption = "Advanced >>"
  Me.Height = 3000
End If
End Sub

Private Sub cmdGfxClear_Click()
txtItmGfx.Text = ""
End Sub

Private Sub cmdItmGfx_Click()
frmSetGraphic.Show vbModal

If uGfx <> "" Then
  txtItmGfx.Text = uGfx
End If
End Sub

Private Sub Form_Load()

If uEdit = False Then uDiag = ""

FillItem cmbSItm
FillDialog cmbSDiag
FillItem cmbCmbItm
FillDialog cmbCmbDiag

If uEdit = True Then
  cmbSItm.Text = uItm
  cmbSDiag.Text = uDiag
  txtItmDesc.Text = uDesc
  txtItmGfx.Text = uGfx
End If
End Sub

Private Sub OKButton_Click()

If Trim$(cmbSItm.Text) = "" Then MsgBox "Please select an item to add!", vbExclamation, "Add Item": Exit Sub

uDesc = txtItmDesc.Text
uDiag = cmbSDiag.Text
uGfx = txtItmGfx.Text

If uEdit Then uItm = cmbSItm.Text

If Not uEdit Then frmRoom.AddItm cmbSItm.Text

uEdit = False
Unload Me

End Sub
