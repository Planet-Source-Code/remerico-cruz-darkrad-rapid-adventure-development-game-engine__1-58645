VERSION 5.00
Begin VB.Form frmItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Editor"
   ClientHeight    =   4230
   ClientLeft      =   2940
   ClientTop       =   4380
   ClientWidth     =   6930
   Icon            =   "frmItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox lstItm 
      Height          =   3990
      Left            =   120
      Pattern         =   "*.itm"
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.CheckBox chkItmRmv 
      Caption         =   "Remove from inventory when used"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Graphic"
      Height          =   1935
      Left            =   2400
      TabIndex        =   5
      Top             =   1680
      Width           =   4335
      Begin VB.CommandButton cmgChImg 
         Caption         =   "Change Image"
         Height          =   495
         Left            =   3120
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.Image ImgItm 
         Height          =   1575
         Left            =   240
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.TextBox txtItmDesc 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtItmName 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   585
      Width           =   3255
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Close"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   3720
      Width           =   975
   End
   Begin DarkEdit.GurhanButton tlNew 
      Height          =   375
      Left            =   2415
      TabIndex        =   8
      ToolTipText     =   "New Item"
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      Picture         =   "frmItem.frx":038A
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      OriginalPicSizeW=   16
      OriginalPicSizeH=   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   0   'False
      SoundOver       =   ""
      SoundClick      =   ""
      MaskColor       =   8421504
   End
   Begin DarkEdit.GurhanButton tlSave 
      Height          =   375
      Left            =   2790
      TabIndex        =   9
      ToolTipText     =   "Save Item"
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      Picture         =   "frmItem.frx":049C
      PictureWidth    =   16
      PictureHeight   =   16
      PictureSize     =   0
      OriginalPicSizeW=   16
      OriginalPicSizeH=   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   0   'False
      SoundOver       =   ""
      SoundClick      =   ""
      MaskColor       =   8421504
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Item Description:"
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      Top             =   1005
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item Name:"
      Height          =   195
      Left            =   2640
      TabIndex        =   1
      Top             =   615
      Width           =   810
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmgChImg_Click()
frmSetGraphic.Show vbModal

If uGfx <> "" Then
  ImgItm.Picture = LoadPicture(WrkDir + uGfx)
  DaItem.itmGfx = uGfx
End If
End Sub

Private Sub Form_Load()
lstItm.Path = WrkDir
lstItm.Refresh
End Sub

Private Sub lstItm_Click()
LoadItm lstItm.List(lstItm.ListIndex)
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub

Private Sub tlSave_Click()
Dim FNum

If Trim(DaItem.itmName) = "" Then DaItem.itmName = InputBox("Enter name for the item", "Save Item")

If Trim(DaItem.itmName) <> "" Then
  If LCase$(Right(DaItem.itmName, 3)) <> ".itm" Then DaItem.itmName = DaItem.itmName + ".itm"

  SaveItm DaItem.itmName

  PackGame GameName
End If
lstItm.Refresh
End Sub

Sub LoadItm(wFile As String)

  ImgItm.Picture = LoadPicture()

  DaItem.itmName = GetInitEntry("Item", "Name", , WrkDir + wFile)
  DaItem.itmDescription = GetInitEntry("Item", "Description", , WrkDir + wFile)
  DaItem.itmGfx = GetInitEntry("Item", "Graphic", , WrkDir + wFile)
   
  DaItem.IsRmv = GetInitEntry("Property", "RemoveAfterUse", , WrkDir + wFile)
  
  txtItmName.Text = DaItem.itmName
  txtItmDesc.Text = DaItem.itmDescription
  If Dir$(WrkDir + DaItem.itmGfx) <> "" And Trim$(DaItem.itmGfx) <> "" Then ImgItm.Picture = LoadPicture(WrkDir + DaItem.itmGfx)
  chkItmRmv.Value = DaItem.IsRmv
End Sub

Sub SaveItm(wFile As String)
  SetInitEntry "Item", "Name", txtItmName.Text, WrkDir + wFile
  SetInitEntry "Item", "Description", txtItmDesc.Text, WrkDir + wFile
  
  SetInitEntry "Item", "Graphic", DaItem.itmGfx, WrkDir + wFile
  
  SetInitEntry "Property", "RemoveAfterUse", chkItmRmv.Value, WrkDir + wFile
End Sub

