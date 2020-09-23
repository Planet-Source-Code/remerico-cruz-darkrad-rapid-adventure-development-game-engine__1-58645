VERSION 5.00
Begin VB.Form frmDialog 
   Caption         =   "Dialog Editor"
   ClientHeight    =   5745
   ClientLeft      =   2910
   ClientTop       =   4350
   ClientWidth     =   8295
   Icon            =   "frmDialog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin DarkEdit.GurhanButton tlImportDg 
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Caption         =   "Import Dialog"
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin VB.CommandButton AddObj 
      Caption         =   "Add"
      Height          =   330
      Left            =   7560
      TabIndex        =   12
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton DwnDiag 
      Caption         =   "\/"
      Height          =   330
      Left            =   4680
      TabIndex        =   10
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton RmvDiag 
      Caption         =   "Remove Line"
      Height          =   330
      Left            =   7080
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton AddDiag 
      Caption         =   "Insert Line"
      Height          =   330
      Left            =   6120
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtDiag 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   6015
   End
   Begin VB.ListBox ConvLst 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      IntegralHeight  =   0   'False
      Left            =   2160
      TabIndex        =   0
      Top             =   1410
      Width           =   6015
   End
   Begin DarkEdit.GurhanButton tlNew 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "New Dialog"
      Top             =   60
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      Picture         =   "frmDialog.frx":038A
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
      Left            =   2540
      TabIndex        =   2
      ToolTipText     =   "Save Current Dialog"
      Top             =   60
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      Picture         =   "frmDialog.frx":049C
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
   Begin VB.FileListBox lstDiag 
      Height          =   5550
      Left            =   60
      Pattern         =   "*.dg"
      TabIndex        =   7
      Top             =   60
      Width           =   2055
   End
   Begin VB.CommandButton NewDiag 
      Caption         =   "New Line"
      Height          =   330
      Left            =   5160
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton UpDiag 
      Caption         =   "/\"
      Height          =   330
      Left            =   4320
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin DarkEdit.GurhanButton tlExportDg 
      Height          =   255
      Left            =   4335
      TabIndex        =   14
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Caption         =   "Export Dialog"
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Edit Line:"
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "\n = Carriage Return"
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "Add"
      Visible         =   0   'False
      Begin VB.Menu mnuAddCmd 
         Caption         =   "&Command"
      End
      Begin VB.Menu mnuAddLbl 
         Caption         =   "&Line Label"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddBlnk 
         Caption         =   "&Blank Space"
      End
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DiagName As String

Private Sub AddDiag_Click()
ConvLst.AddItem txtDiag.Text
End Sub

Private Sub AddObj_Click()
ConvLst.SetFocus
PopupMenu mnuAdd, , AddObj.Left, AddObj.Top + AddObj.Height
End Sub

Private Sub ConvLst_Click()
txtDiag.Text = ConvLst.List(ConvLst.ListIndex)
End Sub

Private Sub Form_Load()
lstDiag.Path = WrkDir
End Sub

Private Sub Form_Resize()
lstDiag.Height = Me.ScaleHeight - 120

ConvLst.Height = Me.ScaleHeight - 1470
ConvLst.Width = Me.ScaleWidth - 2220
txtDiag.Width = ConvLst.Width

RmvDiag.Left = 2160 + (ConvLst.Width - RmvDiag.Width)
AddDiag.Left = RmvDiag.Left - AddDiag.Width
NewDiag.Left = AddDiag.Left - NewDiag.Width

DwnDiag.Left = NewDiag.Left - (DwnDiag.Width + 120)
UpDiag.Left = DwnDiag.Left - UpDiag.Width

AddObj.Left = 2160 + (txtDiag.Width - AddObj.Width)

End Sub

Private Sub lstDiag_Click()
LoadDiag lstDiag.List(lstDiag.ListIndex)
End Sub

Private Sub mnuAddBlnk_Click()
ConvLst.AddItem "", ConvLst.ListIndex
End Sub

Private Sub mnuAddCmd_Click()
frmCmd.Show vbModal
End Sub

Private Sub mnuAddLbl_Click()
a = InputBox("Enter the name of the line label you want", "Add Line Label")
If a <> "" Then ConvLst.AddItem ":" & Trim(a)
End Sub

Private Sub NewDiag_Click()
ConvLst.ListIndex = -1
txtDiag.Text = ""
End Sub

Private Sub RmvDiag_Click()
If ConvLst.ListIndex > -1 Then ConvLst.RemoveItem ConvLst.ListIndex
End Sub

Private Sub tlExportDg_Click()
Dim DgFile As String

If GetFileName(DgFile, "Adventure Dialog Files (*.dg)|*.dg", "Export Dialog File", True) And Trim$(DiagName) <> "" Then
  If LCase$(Trim$(Right$(DgFile, 3))) <> ".dg" Then DgFile = DgFile + ".dg"
  FileCopy WrkDir + DiagName, DgFile
End If

End Sub

Private Sub tlImportDg_Click()
Dim DgFile As String

If GetFileName(DgFile, "Adventure Dialog Files (*.dg)|*.dg", "Import Dialog File") Then
  FileCopy DgFile, WrkDir + GetFName(DgFile)
  lstDiag.Refresh
End If

End Sub

Private Sub tlNew_Click()
DiagName = ""
txtDiag = ""
ConvLst.Clear
End Sub

Private Sub tlSave_Click()
Dim FNum

If Trim(DiagName) = "" Then DiagName = InputBox("Enter name for the dialog", "Save Dialog")

If Trim(DiagName) <> "" Then
  If LCase$(Right(DiagName, 3)) <> ".dg" Then DiagName = DiagName + ".dg"
  
  FNum = FreeFile
  
  Open WrkDir + DiagName For Output As #FNum
    For a = 0 To ConvLst.ListCount - 1
      Print #FNum, ConvLst.List(a)
    Next a
  Close #FNum
  PackGame GameName
End If
lstDiag.Refresh
End Sub

Private Sub txtDiag_Change()
If ConvLst.ListIndex > -1 Then ConvLst.List(ConvLst.ListIndex) = txtDiag.Text
End Sub

Private Sub txtDiag_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 13
    AddDiag_Click
    txtDiag.Text = ""
End Select
End Sub

Private Sub UpDiag_Click()
Dim X1 As String, X2 As String

If ConvLst.ListIndex > 0 Then
  X1 = ConvLst.List(ConvLst.ListIndex)
  X2 = ConvLst.List(ConvLst.ListIndex - 1)
  VSwap X1, X2
  ConvLst.List(ConvLst.ListIndex) = X1
  ConvLst.List(ConvLst.ListIndex - 1) = X2
  ConvLst.ListIndex = ConvLst.ListIndex - 1
End If
End Sub

Sub VSwap(X1 As Variant, X2 As Variant)

    Dim T1 As Variant
    T1 = X1
    X1 = X2
    X2 = T1

End Sub

Private Sub DwnDiag_Click()
Dim X1 As String, X2 As String

If ConvLst.ListIndex < ConvLst.ListCount - 1 Then
  X1 = ConvLst.List(ConvLst.ListIndex)
  X2 = ConvLst.List(ConvLst.ListIndex + 1)
  VSwap X1, X2
  ConvLst.List(ConvLst.ListIndex) = X1
  ConvLst.List(ConvLst.ListIndex + 1) = X2
  ConvLst.ListIndex = ConvLst.ListIndex + 1
End If
End Sub

Sub LoadDiag(dFName As String)
Dim FNum, xLine As String

txtDiag.Text = ""
ConvLst.Clear

FNum = FreeFile

Open WrkDir + dFName For Input As #FNum
  Do Until EOF(FNum)
    Line Input #FNum, xLine
      ConvLst.AddItem xLine
  Loop
Close #FNum
DiagName = dFName
End Sub
