VERSION 5.00
Begin VB.Form frmAppear 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Appearance Editor"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   Icon            =   "frmAppear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7080
      TabIndex        =   25
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8280
      TabIndex        =   24
      Top             =   6120
      Width           =   1095
   End
   Begin VB.PictureBox dTab 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5655
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   360
      Width           =   9255
      Begin VB.VScrollBar VScroll2 
         Height          =   2775
         Left            =   2760
         TabIndex        =   11
         Top             =   2760
         Width           =   255
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   2655
         TabIndex        =   10
         Top             =   2760
         Width           =   2655
         Begin VB.PictureBox PropCont 
            BorderStyle     =   0  'None
            Height          =   1215
            Left            =   0
            ScaleHeight     =   1215
            ScaleWidth      =   2655
            TabIndex        =   20
            Top             =   0
            Width           =   2655
            Begin VB.TextBox PropVal 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   960
               TabIndex        =   21
               Top             =   -285
               Visible         =   0   'False
               Width           =   1680
            End
            Begin VB.Label PropTitle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Property"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   0
               Left            =   0
               TabIndex        =   22
               Top             =   -285
               Visible         =   0   'False
               Width           =   975
            End
         End
      End
      Begin VB.VScrollBar VScroll 
         Height          =   5175
         Left            =   8880
         TabIndex        =   9
         Top             =   120
         Width           =   255
      End
      Begin VB.HScrollBar HScroll 
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   5280
         Width           =   5775
      End
      Begin VB.ListBox GUIList 
         Appearance      =   0  'Flat
         Height          =   2175
         ItemData        =   "frmAppear.frx":038A
         Left            =   120
         List            =   "frmAppear.frx":039A
         TabIndex        =   7
         Top             =   120
         Width           =   2895
      End
      Begin VB.PictureBox ContPic 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   3120
         ScaleHeight     =   5145
         ScaleWidth      =   5745
         TabIndex        =   6
         Top             =   120
         Width           =   5775
         Begin VB.PictureBox dNav 
            Appearance      =   0  'Flat
            BackColor       =   &H00400000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1650
            Left            =   0
            ScaleHeight     =   110
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   640
            TabIndex        =   12
            Top             =   0
            Visible         =   0   'False
            Width           =   9600
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               ForeColor       =   &H80000008&
               Height          =   1080
               Left            =   8280
               ScaleHeight     =   70
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   70
               TabIndex        =   13
               Top             =   300
               Width           =   1080
               Begin VB.Image dDir 
                  Height          =   360
                  Index           =   0
                  Left            =   0
                  Picture         =   "frmAppear.frx":03CC
                  Top             =   0
                  Width           =   360
               End
               Begin VB.Image dDir 
                  Height          =   360
                  Index           =   1
                  Left            =   360
                  Picture         =   "frmAppear.frx":0476
                  Top             =   0
                  Width           =   360
               End
               Begin VB.Image dDir 
                  Height          =   360
                  Index           =   2
                  Left            =   720
                  Picture         =   "frmAppear.frx":0520
                  Top             =   0
                  Width           =   360
               End
               Begin VB.Image dDir 
                  Height          =   360
                  Index           =   3
                  Left            =   0
                  Picture         =   "frmAppear.frx":05CA
                  Top             =   360
                  Width           =   360
               End
               Begin VB.Image dDir 
                  Height          =   360
                  Index           =   5
                  Left            =   720
                  Picture         =   "frmAppear.frx":0674
                  Top             =   360
                  Width           =   360
               End
               Begin VB.Image dDir 
                  Height          =   360
                  Index           =   6
                  Left            =   0
                  Picture         =   "frmAppear.frx":071E
                  Top             =   720
                  Width           =   360
               End
               Begin VB.Image dDir 
                  Height          =   360
                  Index           =   7
                  Left            =   360
                  Picture         =   "frmAppear.frx":07C8
                  Top             =   720
                  Width           =   360
               End
               Begin VB.Image dDir 
                  Height          =   360
                  Index           =   8
                  Left            =   720
                  Picture         =   "frmAppear.frx":0872
                  Top             =   720
                  Width           =   360
               End
            End
            Begin VB.Label dChoice 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "[O] - Options"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   9
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   120
               Width           =   7935
            End
            Begin VB.Label dMsg 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Navigation text goes in here."
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   9
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   855
               Left            =   120
               TabIndex        =   14
               Top             =   600
               Width           =   7935
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00800000&
               X1              =   0
               X2              =   640
               Y1              =   0
               Y2              =   0
            End
         End
         Begin VB.PictureBox dExit 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   2055
            Left            =   0
            ScaleHeight     =   2055
            ScaleWidth      =   6135
            TabIndex        =   34
            Top             =   0
            Visible         =   0   'False
            Width           =   6135
            Begin VB.Shape shape 
               BorderColor     =   &H00FFC0C0&
               Height          =   1695
               Left            =   120
               Top             =   120
               Width           =   5775
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
               TabIndex        =   37
               Top             =   240
               Width           =   5535
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
               TabIndex        =   36
               Top             =   1320
               Width           =   1455
            End
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
               TabIndex        =   35
               Top             =   1320
               Width           =   1815
            End
         End
         Begin VB.PictureBox xMenu 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Height          =   5535
            Left            =   0
            ScaleHeight     =   5535
            ScaleWidth      =   2895
            TabIndex        =   26
            Top             =   0
            Visible         =   0   'False
            Width           =   2895
            Begin VB.Label MenuTitle 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "[Menu]"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   13.5
                  Charset         =   255
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   270
               Left            =   930
               TabIndex        =   33
               Top             =   120
               Width           =   1035
            End
            Begin VB.Label dIMnu 
               BackColor       =   &H00400000&
               BackStyle       =   0  'Transparent
               Caption         =   " Exit Game"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   13.5
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   270
               Index           =   5
               Left            =   240
               TabIndex        =   32
               Top             =   4680
               Width           =   2340
            End
            Begin VB.Label dIMnu 
               BackColor       =   &H00400000&
               BackStyle       =   0  'Transparent
               Caption         =   " Game Options"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   13.5
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   270
               Index           =   4
               Left            =   240
               TabIndex        =   31
               Top             =   3900
               Width           =   2340
            End
            Begin VB.Label dIMnu 
               BackColor       =   &H00400000&
               BackStyle       =   0  'Transparent
               Caption         =   " Save Game"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   13.5
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   30
               Top             =   3120
               Width           =   2340
            End
            Begin VB.Label dIMnu 
               BackColor       =   &H00400000&
               BackStyle       =   0  'Transparent
               Caption         =   " Load Game"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   13.5
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   270
               Index           =   2
               Left            =   240
               TabIndex        =   29
               Top             =   2280
               Width           =   2340
            End
            Begin VB.Label dIMnu 
               BackColor       =   &H00400000&
               BackStyle       =   0  'Transparent
               Caption         =   " Inventory"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   13.5
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   270
               Index           =   1
               Left            =   240
               TabIndex        =   28
               Top             =   1560
               Width           =   2340
            End
            Begin VB.Label dIMnu 
               BackColor       =   &H00400000&
               Caption         =   " Continue Game"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   13.5
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   270
               Index           =   0
               Left            =   240
               TabIndex        =   27
               Top             =   840
               Width           =   2340
            End
            Begin VB.Shape Shape5 
               BorderColor     =   &H00FFC0C0&
               Height          =   4695
               Left            =   165
               Top             =   600
               Width           =   2535
            End
         End
         Begin VB.PictureBox dInv 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Height          =   3975
            Left            =   0
            ScaleHeight     =   3975
            ScaleWidth      =   5655
            TabIndex        =   16
            Top             =   0
            Visible         =   0   'False
            Width           =   5655
            Begin VB.PictureBox InvDisp 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               ForeColor       =   &H80000008&
               Height          =   3015
               Left            =   120
               ScaleHeight     =   2985
               ScaleWidth      =   5385
               TabIndex        =   17
               Top             =   360
               Width           =   5415
            End
            Begin VB.Label InvTBar 
               BackColor       =   &H00400000&
               Caption         =   "Inventory"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   12
                  Charset         =   255
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   255
               Left            =   0
               TabIndex        =   19
               Top             =   0
               Width           =   5655
            End
            Begin VB.Label InvClose 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Close"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   13.5
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   4440
               TabIndex        =   18
               Top             =   3520
               Width           =   1095
            End
            Begin VB.Shape Shape6 
               BackColor       =   &H00400000&
               BackStyle       =   1  'Opaque
               Height          =   375
               Left            =   4440
               Top             =   3480
               Width           =   1095
            End
         End
      End
      Begin VB.Label PropBar 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "Properties"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   2895
      End
   End
   Begin VB.PictureBox dTab 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5655
      Index           =   2
      Left            =   120
      ScaleHeight     =   5625
      ScaleWidth      =   9225
      TabIndex        =   5
      Top             =   360
      Width           =   9255
   End
   Begin VB.PictureBox dTab 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5655
      Index           =   1
      Left            =   120
      ScaleHeight     =   5625
      ScaleWidth      =   9225
      TabIndex        =   4
      Top             =   360
      Width           =   9255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   10
      X2              =   625
      Y1              =   401
      Y2              =   401
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   625
      X2              =   625
      Y1              =   26
      Y2              =   402
   End
   Begin VB.Label xTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GUI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label xTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Game Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label xTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu mnuProp 
      Caption         =   "Properties"
      Visible         =   0   'False
      Begin VB.Menu mnuWProp 
         Caption         =   "wProp"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmAppear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cGUI As String
Dim dCont As PictureBox
Dim PropLst() As String
Dim ObjSel As String
Dim uIndx As Integer
Dim GUIFile As String
Dim rObj As Object

Private Sub cmdClose_Click()
a = MsgBox("Are you sure you want to close this Editor?", vbYesNo, "Appearance Editor")
If a = vbYes Then Unload Me
End Sub

Private Sub cmdSave_Click()
If GameName <> "" Then PackGame GameName
End Sub

Private Sub dChoice_Click()
PropInit "Navigation Selection", dChoice
AddProp "TextColor"
AddProp "Font"
End Sub

Private Sub dDir_Click(Index As Integer)
PropInit "Navigation Arrow " & Index, dDir(Index)
AddProp "Image", "Gfx"

End Sub

Private Sub dExit_Click()
PropInit "Exit Window", dExit
AddProp "Background", "Gfx"
AddProp "ShowBorder", "Bool"
GetPropVal
End Sub

Private Sub dIMnu_Click(Index As Integer)
PropInit "Menu Text " & Index, dIMnu(Index)
AddProp "BackColor"
AddProp "ForeColor"
AddProp "Font"
AddProp "Caption"
GetPropVal
End Sub

Private Sub dInv_Click()
PropInit "Inventory Window", dInv
AddProp "Background", "Gfx"
GetPropVal
End Sub

Private Sub dMsg_Click()
PropInit "Navigation Text", dMsg
AddProp "TextColor"
AddProp "Font"
GetPropVal
End Sub

Private Sub dNav_Click()
PropInit "Navigation Bar", dNav
AddProp "Background", "Gfx"
GetPropVal
End Sub

Private Sub dPrompt_Click()
PropInit "Exit Prompt", dPrompt
AddProp "BackColor"
AddProp "ForeColor"
AddProp "Font"
AddProp "Caption"
GetPropVal
End Sub

Private Sub GUIList_Click()
If GUIList.Text <> "" Then cGUI = GUIList.Text Else Exit Sub

xMenu.Visible = False
dInv.Visible = False
dNav.Visible = False
dExit.Visible = False

Select Case cGUI
  Case "Menu"
    xMenu.Visible = True
    Set dCont = xMenu
    GUIFile = "menu.gui"
    
    dIMnu_Click 0
    dIMnu_Click 1
    dIMnu_Click 2
    dIMnu_Click 3
    dIMnu_Click 4
    dIMnu_Click 5
    
    xMenu_Click
    
  Case "Inventory"
    dInv.Visible = True
    Set dCont = dInv
    GUIFile = "inv.gui"
    
    InvClose_Click
    InvTBar_Click
    InvDisp_Click
    dInv_Click
  Case "Navigation Bar"
    dNav.Visible = True
    Set dCont = dNav
    GUIFile = "nav.gui"
    
    dMsg_Click
    dChoice_Click
    For a = 0 To 7
      If a <> 4 Then dDir_Click Int(a)
    Next a
    dNav_Click
  Case "Exit Dialog"
    dExit.Visible = True
    Set dCont = dExit
    GUIFile = "exit.gui"
    
    PYes_Click
    PNo_Click
    dPrompt_Click
    dExit_Click
End Select

ContChange

End Sub

Private Sub HScroll_Change()
    dCont.Left = ContPic.Left - HScroll.Value
End Sub

Private Sub HScroll_GotFocus()
ContPic.SetFocus
End Sub

Private Sub HScroll_Scroll()
    dCont.Left = ContPic.Left - HScroll.Value
End Sub

Private Sub InvClose_Click()
PropInit "Close Button", InvClose
AddProp "Caption"
AddProp "Background", "Gfx"
AddProp "BackColor"
AddProp "ForeColor"
GetPropVal
End Sub

Private Sub InvDisp_Click()
PropInit "Inventory Display", InvDisp
AddProp "Background"
AddProp "BackColor"
AddProp "ForeColor"
GetPropVal
End Sub

Private Sub InvTBar_Click()
PropInit "Inventory Titlebar", InvTBar
AddProp "BackColor"
AddProp "ForeColor"
AddProp "Caption"
GetPropVal
End Sub

Private Sub MenuTitle_Click()
PropInit "Menu Title", MenuTitle
AddProp "BackColor"
AddProp "ForeColor"
AddProp "Font"
AddProp "Caption"
AddProp "Visible", "Bool"
GetPropVal
End Sub

Private Sub mnuWProp_Click(Index As Integer)
PropVal(uIndx).Text = mnuWProp(Index).Caption
End Sub

Private Sub PNo_Click()
PropInit "'No' Button", PNo
AddProp "BackColor"
AddProp "ForeColor"
AddProp "Font"
AddProp "Caption"
GetPropVal
End Sub

Private Sub PropTitle_Click(Index As Integer)
PropVal(Index).SetFocus
End Sub

Private Sub PropTitle_DblClick(Index As Integer)
PropVal(Index).SetFocus
End Sub

Private Sub PropVal_Change(Index As Integer)
SetInitEntry Mid$(PropBar.Caption, 2, Len(PropBar.Caption) - 2), PropTitle(Index).Caption, PropVal(Index).Text, WrkDir + GUIFile

Select Case PropTitle(Index).Caption
  Case "Caption"
    rObj.Caption = PropVal(Index).Text
  Case "Image", "Background"
    If Dir$(WrkDir + PropVal(Index).Text) <> "" And PropVal(Index).Text <> "" Then rObj.Picture = LoadPicture(WrkDir + PropVal(Index).Text) 'Else rObj.Picture = LoadPicture()
    
End Select
End Sub

Private Sub PropVal_Click(Index As Integer)
If PropVal(Index).Tag <> "" Then
  uIndx = Index
  FillPropMenu PropVal(Index).Tag
  PopupMenu mnuProp
End If
End Sub

Private Sub PYes_Click()
PropInit "'Yes' Button", PYes
AddProp "BackColor"
AddProp "ForeColor"
AddProp "Font"
AddProp "Caption"
GetPropVal
End Sub

Private Sub VScroll_Change()
    dCont.Top = ContPic.Top - VScroll.Value
End Sub

Private Sub VScroll_GotFocus()
ContPic.SetFocus
End Sub

Private Sub VScroll_Scroll()
    dCont.Top = ContPic.Top - VScroll.Value
End Sub

Private Sub VScroll2_Change()
    PropCont.Top = Picture3.Top - VScroll2.Value
End Sub

Private Sub VScroll2_Scroll()
    PropCont.Top = Picture3.Top - VScroll2.Value
End Sub

Private Sub xMenu_Click()
PropInit "Menu Window", xMenu
AddProp "Background", "Gfx"
AddProp "ShowBorder", "Bool"
GetPropVal
End Sub

Private Sub xTab_Click(Index As Integer)
dTab(Index).ZOrder
End Sub

Sub ContChange()
    VScroll.Min = ContPic.Top
    HScroll.Min = ContPic.Left

    VScroll.Max = dCont.Height - ContPic.Height
    VScroll.Enabled = True
    
    If VScroll.Max < VScroll.Min Then
        VScroll.Max = VScroll.Min
        VScroll.Enabled = False
    End If
    
    HScroll.Max = dCont.Width - ContPic.Width
    HScroll.Enabled = True
    
    If HScroll.Max < HScroll.Min Then
        HScroll.Max = HScroll.Min
        HScroll.Enabled = False
    End If
    
    VScroll.LargeChange = dCont.Height / 10
    HScroll.LargeChange = dCont.Width / 10
End Sub

Sub AddProp(dProp As String, Optional wType As String)
  Load PropTitle(PropTitle.Count)
  Load PropVal(PropVal.Count)
  ReDim Preserve PropLst(PropTitle.UBound)
  
  PropTitle(PropTitle.UBound).Caption = dProp
  PropTitle(PropTitle.UBound).Top = (PropTitle(PropTitle.UBound - 1).Top + 285) - 15
  PropVal(PropVal.UBound).Top = (PropVal(PropVal.UBound - 1).Top + 285) - 15
  PropCont.Height = PropTitle(PropTitle.UBound).Top + 285
  
  PropTitle(PropTitle.UBound).Visible = True
  PropVal(PropVal.UBound).Visible = True
    
    VScroll2.Min = Picture3.Top

    VScroll2.Max = PropCont.Height - Picture3.Height
    VScroll2.Enabled = True
    
    If VScroll2.Max < VScroll2.Min Then
        VScroll2.Max = VScroll2.Min
        VScroll2.Enabled = False
    End If
    
    VScroll2.LargeChange = PropCont.Height / 10
    
    If wType <> "" Then PropVal(PropVal.UBound).Tag = wType
End Sub

Sub ClearProps()
  Do Until PropTitle.Count = 1
    Unload PropTitle(PropTitle.UBound)
    Unload PropVal(PropVal.UBound)
  Loop
  PropCont.Height = 285
  VScroll2.Enabled = False
End Sub

Sub PropInit(dObj As String, xObj As Object)
PropBar.Caption = "[" & dObj & "]"
ObjSel = dObj
ClearProps
Set rObj = xObj
End Sub

Private Sub FillPropMenu(wType As String)
Dim dPic As String

mnuWProp(0).Visible = True
Do Until mnuWProp.Count = 1
  Unload mnuWProp(mnuWProp.UBound)
Loop

If wType = "Gfx" Then
  dPic = Dir(WrkDir + "*.*")
  Do Until dPic = ""
    Select Case LCase$(Right$(dPic, 3))
      Case "bmp", "gif", "jpg", "wmf"
        If wType = "Gfx" Then
          Load mnuWProp(mnuWProp.Count)
          mnuWProp(mnuWProp.UBound).Caption = dPic
        End If
    End Select
    dPic = Dir()
  Loop
ElseIf wType = "Bool" Then
  Load mnuWProp(mnuWProp.Count)
  mnuWProp(mnuWProp.UBound).Caption = "True"
  Load mnuWProp(mnuWProp.Count)
  mnuWProp(mnuWProp.UBound).Caption = "False"
End If

mnuWProp(0).Visible = False

End Sub

Sub GetPropVal()
Dim dDefault As String

For a = 1 To PropVal.UBound
  If PropVal(a).Tag = "Bool" Then
    dDefault = "True"
  Else
    dDefault = ""
  End If
  
 PropVal(a).Text = GetInitEntry(Mid$(PropBar.Caption, 2, Len(PropBar.Caption) - 2), PropTitle(a).Caption, dDefault, WrkDir + GUIFile)
Next a
End Sub
