VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "DarkQuest Engine"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox dScreen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7200
      Left            =   0
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9600
      Begin VB.PictureBox dNav 
         Appearance      =   0  'Flat
         BackColor       =   &H00450000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1650
         Left            =   0
         ScaleHeight     =   110
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   640
         TabIndex        =   2
         Top             =   5520
         Visible         =   0   'False
         Width           =   9600
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1080
            Left            =   8280
            ScaleHeight     =   70
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   70
            TabIndex        =   5
            Top             =   300
            Width           =   1080
            Begin VB.Image dDir 
               Height          =   360
               Index           =   8
               Left            =   720
               Picture         =   "frmMain.frx":0000
               Top             =   720
               Width           =   360
            End
            Begin VB.Image dDir 
               Height          =   360
               Index           =   7
               Left            =   360
               Picture         =   "frmMain.frx":00AA
               Top             =   720
               Width           =   360
            End
            Begin VB.Image dDir 
               Height          =   360
               Index           =   6
               Left            =   0
               Picture         =   "frmMain.frx":0154
               Top             =   720
               Width           =   360
            End
            Begin VB.Image dDir 
               Height          =   360
               Index           =   5
               Left            =   720
               Picture         =   "frmMain.frx":01FE
               Top             =   360
               Width           =   360
            End
            Begin VB.Image dDir 
               Height          =   360
               Index           =   3
               Left            =   0
               Picture         =   "frmMain.frx":02A8
               Top             =   360
               Width           =   360
            End
            Begin VB.Image dDir 
               Height          =   360
               Index           =   2
               Left            =   720
               Picture         =   "frmMain.frx":0352
               Top             =   0
               Width           =   360
            End
            Begin VB.Image dDir 
               Height          =   360
               Index           =   1
               Left            =   360
               Picture         =   "frmMain.frx":03FC
               Top             =   0
               Width           =   360
            End
            Begin VB.Image dDir 
               Height          =   360
               Index           =   0
               Left            =   0
               Picture         =   "frmMain.frx":04A6
               Top             =   0
               Width           =   360
            End
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00800000&
            X1              =   0
            X2              =   640
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Label dMsg 
            BackStyle       =   0  'Transparent
            Caption         =   "Ang Saya Saya."
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
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   7695
         End
         Begin VB.Label dChoice 
            BackStyle       =   0  'Transparent
            Caption         =   "[I] - Inventory | [ENTER] - Menu"
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
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   8055
         End
      End
      Begin VB.PictureBox dInv 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   3720
         ScaleHeight     =   3975
         ScaleWidth      =   5655
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   5655
         Begin VB.PictureBox InvMnu 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   1920
            ScaleHeight     =   735
            ScaleWidth      =   1095
            TabIndex        =   22
            Top             =   1560
            Visible         =   0   'False
            Width           =   1095
            Begin VB.Label imnu 
               BackColor       =   &H00400000&
               BackStyle       =   0  'Transparent
               Caption         =   " » Drop"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   25
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label imnu 
               BackColor       =   &H00400000&
               BackStyle       =   0  'Transparent
               Caption         =   " » Examine"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   24
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label imnu 
               BackColor       =   &H00400000&
               BackStyle       =   0  'Transparent
               Caption         =   " » Use"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox InvCont 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3015
            Left            =   120
            ScaleHeight     =   2985
            ScaleWidth      =   5385
            TabIndex        =   9
            Top             =   360
            Width           =   5415
            Begin VB.VScrollBar InvScroll 
               Height          =   2985
               Left            =   5145
               TabIndex        =   19
               Top             =   0
               Width           =   255
            End
            Begin VB.PictureBox InvScCont 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   2775
               Left            =   0
               ScaleHeight     =   185
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   345
               TabIndex        =   20
               Top             =   0
               Width           =   5175
               Begin VB.Shape shSel 
                  BorderColor     =   &H00FFC0C0&
                  Height          =   1095
                  Left            =   1320
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.Image InvThumb 
                  Height          =   1080
                  Index           =   0
                  Left            =   120
                  Stretch         =   -1  'True
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1080
               End
            End
         End
         Begin VB.Label InvDesc 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   495
            Left            =   120
            TabIndex        =   21
            Top             =   3440
            Width           =   4305
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
            TabIndex        =   10
            Top             =   3520
            Width           =   1095
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
            TabIndex        =   8
            Top             =   0
            Width           =   5655
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
      Begin VB.Image xItm 
         Height          =   495
         Index           =   0
         Left            =   -1060
         MousePointer    =   2  'Cross
         Top             =   -490
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Ito ang screen na walang laman."
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
         Left            =   2040
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   5580
      End
   End
   Begin VB.PictureBox pTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7200
      Left            =   0
      ScaleHeight     =   476
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   636
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.PictureBox dMenu 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   2895
      TabIndex        =   11
      Top             =   0
      Width           =   2895
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartY, StartX
Dim CSlot As Integer

Private Sub dChoice_Click()
dNav.SetFocus
End Sub

Private Sub dDir_Click(Index As Integer)
Dim wRM As String
wRM = ""

Select Case Index
  Case 0  'NW
    wRM = DaRoom.dNorthWest
  Case 1  'N
    wRM = DaRoom.dNorth
  Case 2  'NE
    wRM = DaRoom.dNorthEast
  Case 3  'W
    wRM = DaRoom.dWest
  Case 5  'E
    wRM = DaRoom.dEast
  Case 6  'SW
    wRM = DaRoom.dSouthWest
  Case 7  'S
    wRM = DaRoom.dSouth
  Case 8  'SE
    wRM = DaRoom.dSouthEast
End Select

If GScreen = "game" Then
  If LCase$(Right$(wRM, 2)) = "rm" Then GotoRoom wRM
  If LCase$(Right$(wRM, 2)) = "dg" Then ReadDialog wRM
End If

End Sub

Private Sub dIMnu_Click(Index As Integer)
Select Case Index
  Case 0
    For a = 0 To -dMenu.Width Step -5
      dMenu.Left = a
      DoEvents
    Next a
    dMenu.Left = -dMenu.Width
    dScreen.Enabled = True
    dMenu.Enabled = False
    dScreen.SetFocus
    If LCase$(Right$(GMusic, 3)) = "mid" Then MIDIVolume 0
  Case 3
    ShowSave True
  Case 5
    frmExit.Show vbModal
End Select
End Sub

Private Sub dInventory_Click()
If dInventory.ListIndex < 0 Then Exit Sub
MsgBox GetInitEntry("Item", "Description", , WrkDir + Inventory(dInventory.ListIndex + 1))
End Sub

Private Sub dInventory_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 27
    dInventory.ListIndex = -1
    dInventory.Enabled = False
    dLstMenu.Enabled = True
    dLstMenu.SetFocus
End Select
End Sub

Private Sub dIMnu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
dIMnu(CSlot).BackStyle = 0
CSlot = Index
dIMnu(CSlot).BackStyle = 1
End Sub

Private Sub dMenu_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 38   ' UP
    dIMnu(CSlot).BackStyle = 0
    CSlot = CSlot - 1
    If CSlot = -1 Then CSlot = dIMnu.UBound
    dIMnu(CSlot).BackStyle = 1
  Case 40   ' DOWN
    dIMnu(CSlot).BackStyle = 0
    CSlot = CSlot + 1
    If CSlot = dIMnu.Count Then CSlot = 0
    dIMnu(CSlot).BackStyle = 1
  Case 13  ' ENTER
    dIMnu_Click CSlot
End Select
End Sub

Private Sub dNav_KeyDown(KeyCode As Integer, Shift As Integer)
dScreen_KeyDown KeyCode, Shift
End Sub

Private Sub dScreen_KeyDown(KeyCode As Integer, Shift As Integer)

If GScreen = "title" And KeyCode <> 27 Then
    'CheckGameSave
    
    ' We run the startup dialog, if any.
    If GetInitEntry("Dialog", "Startup", "", WrkDir + "Config.cfg") <> "" Then ReadDialog GetInitEntry("Dialog", "Startup", "", WrkDir + "Config.cfg")

    
    ' After the dialog, we now go to the room!
    GotoRoom GetInitEntry("Init", "InitRoom", , WrkDir + "Config.cfg")

    'dScreen.SetFocus
End If

Select Case KeyCode
  Case 27   ' ESC
    If GScreen = "title" Then
      Unload Me
      End
    End If
  Case 13   ' ENTER [For Menu...or something like that]
    If GScreen = "game" And DaRoom.DisableMenu = vbUnchecked Then
      dScreen.Enabled = False
      dMenu.Enabled = True
      dMenu.Left = -dMenu.Width
      dMenu.ZOrder
      
     '  dInventory.Clear
     '  For a = 1 To Inventory.Count
     '    dInventory.AddItem GetInitEntry("Item", "Name", , WrkDir + Inventory.Item(a))
     '  Next a
     '  Debug.Print Inventory.Count
       
      
      For a = -dMenu.Width To 0 Step 5
        dMenu.Left = a
        DoEvents
      Next a
      
      If LCase$(Right$(GMusic, 3)) = "mid" Then MIDIVolume -900
    End If
  Case 73   ' I
    ListInventory
    dInv.Visible = True
  Case 36   ' NUM7 - NW
    If GScreen = "game" Then
      If LCase$(Right$(DaRoom.dNorthWest, 2)) = "rm" Then
        GotoRoom DaRoom.dNorthWest
      ElseIf LCase$(Right$(DaRoom.dNorthWest, 2)) = "dg" Then
        ReadDialog DaRoom.dNorthWest
      End If
    End If
  Case 38   ' NUM8 - N
    If GScreen = "game" Then
      If LCase$(Right$(DaRoom.dNorth, 2)) = "rm" Then
        GotoRoom DaRoom.dNorth
      ElseIf LCase$(Right$(DaRoom.dNorth, 2)) = "dg" Then
        ReadDialog DaRoom.dNorth
      End If
    End If
  Case 33   ' NUM9 - NE
    If GScreen = "game" Then
      If LCase$(Right$(DaRoom.dNorthEast, 2)) = "rm" Then
        GotoRoom DaRoom.dNorthEast
      ElseIf LCase$(Right$(DaRoom.dNorthEast, 2)) = "dg" Then
        ReadDialog DaRoom.dNorthEast
      End If
    End If
  Case 37   ' NUM4 - W
    If GScreen = "game" Then
      If LCase$(Right$(DaRoom.dWest, 2)) = "rm" Then
        GotoRoom DaRoom.dWest
      ElseIf LCase$(Right$(DaRoom.dWest, 2)) = "dg" Then
        ReadDialog DaRoom.dWest
      End If
    End If
  Case 39   ' NUM6 - E
    If GScreen = "game" Then
      If LCase$(Right$(DaRoom.dEast, 2)) = "rm" Then
        GotoRoom DaRoom.dEast
      ElseIf LCase$(Right$(DaRoom.dEast, 2)) = "dg" Then
        ReadDialog DaRoom.dEast
      End If
    End If
  Case 35   ' NUM1 - SW
    If GScreen = "game" Then
      If LCase$(Right$(DaRoom.dSouthWest, 2)) = "rm" Then
        GotoRoom DaRoom.dSouthWest
      ElseIf LCase$(Right$(DaRoom.dSouthWest, 2)) = "dg" Then
        ReadDialog DaRoom.dSouthWest
      End If
    End If
  Case 40   ' NUM2 - S
    If GScreen = "game" Then
      If LCase$(Right$(DaRoom.dSouth, 2)) = "rm" Then
        GotoRoom DaRoom.dSouth
      ElseIf LCase$(Right$(DaRoom.dSouth, 2)) = "dg" Then
        ReadDialog DaRoom.dSouth
      End If
    End If
  Case 34   ' NUM3 - SE
    If GScreen = "game" Then
      If LCase$(Right$(DaRoom.dSouthEast, 2)) = "rm" Then
        GotoRoom DaRoom.dSouthEast
      ElseIf LCase$(Right$(DaRoom.dSouthEast, 2)) = "dg" Then
        ReadDialog DaRoom.dSouthEast
      End If
    End If
End Select

End Sub

Private Sub dMsg_Click()
dNav.SetFocus
End Sub

Private Sub dScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  
  If GScreen <> "game" Then Exit Sub

  Dim I As Integer
  Dim LRGN As Long
  Dim ISel As Boolean

  LRGN = IsInRegion(AllAreas, x, y)
  ISel = False
  For I = 1 To AllAreas.Count
  If AllAreas(I).AreaNumber = LRGN Then
        ISel = True
        If AllAreas(I).AreaSelected = False Then
            AllAreas(I).AreaSelected = True
            'InvertRgn AllAreas.ParentHDC, AllAreas(I).AreaNumber
           ' dScreen.Cls
            If DaRoom.DontShowNav = vbUnchecked Then dMsg.Caption = DaRoom.Area(I, 1)
            If DaRoom.DisableTooltip = vbUnchecked Then
              dScreen.CurrentX = x
              dScreen.CurrentY = y
              dScreen.Print DaRoom.Area(I, 0)
              Screen.MousePointer = 2
            End If
        End If
  Else
        If AllAreas(I).AreaSelected = True Then
            AllAreas(I).AreaSelected = False
            'InvertRgn AllAreas.ParentHDC, AllAreas(I).AreaNumber
            dScreen.Cls
            dMsg.Caption = ""
            Screen.MousePointer = 0
        End If
  End If
  Next I
  
  If ISel = False Then dMsg.Caption = ""
End Sub

Private Sub dScreen_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim I As Integer
Dim LRGN As Long

If GScreen = "title" Then
    'CheckGameSave
    
    ' We run the startup dialog, if any.
    If GetInitEntry("Dialog", "Startup", "", WrkDir + "Config.cfg") <> "" Then ReadDialog GetInitEntry("Dialog", "Startup", "", WrkDir + "Config.cfg")

    
    ' After the dialog, we now go to the room!
    GotoRoom GetInitEntry("Init", "InitRoom", , WrkDir + "Config.cfg")
    dScreen.SetFocus
End If


If GScreen <> "game" Then Exit Sub

If DaRoom.AreaCount > 0 Then
  LRGN = IsInRegion(AllAreas, x, y)
  If LRGN <> 0 Then
    For I = 1 To AllAreas.Count
    If AllAreas(I).AreaNumber = LRGN Then
        'InvertRgn AllAreas.ParentHDC, AllAreas(I).AreaNumber
        'dMsg.Caption = "Nothing to do with that right now."
        If DaRoom.Area(I, 2) <> "" And Left$(DaRoom.Area(I, 2), 1) <> "#" Then ReadDialog DaRoom.Area(I, 2)
        If Left$(DaRoom.Area(I, 2), 1) = "#" Then ExecuteLine Trim$(DaRoom.Area(I, 2)), 1
        Exit Sub
    End If
    Next I
  End If
End If
End Sub

Private Sub dScreen_Paint()
'dScreen.Cls
Dim I As Integer
For I = 1 To AllAreas.Count
If AllAreas(I).AreaState = 0 Then
    PaintARGN AllAreas.ParentHDC, AllAreas(I).AreaNumber, AllAreas(I).AreaPen, AllAreas(I).AreaBrush
Else
    PaintARGN AllAreas.ParentHDC, AllAreas(I).AreaNumber, AllAreas(I).AreaPen, AllAreas(I).AreaAlertBrush
End If
        
Next I

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
dScreen_KeyDown KeyCode, Shift
End Sub

Private Sub Form_Load()

'ChangeRes 640, 480

InitMIDI

Set AllAreas = New Areas
AllAreas.ParentHDC = dScreen.Hdc

Me.Caption = GetInitEntry("General", "Title", , WrkDir + "Config.cfg")

GScreen = "title"
If DisplayTitleScreen = False Then dScreen_KeyDown 0, 0

ApplyTheme "menu"
ApplyTheme "nav"

End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteTmp
End Sub

Private Sub InvClose_Click()
InvMnu.Visible = False
dInv.Visible = False
End Sub



Private Sub InvMnu_LostFocus()
  InvMnu.Visible = False
End Sub

Private Sub InvScCont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
InvMnu.Visible = False
shSel.Visible = False
InvDesc.Caption = ""
End Sub

Private Sub InvTBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    StartX = x
    StartY = y
End Sub

Private Sub InvTBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
        dInv.Left = IIf(x < StartX, dInv.Left - (StartX - x), dInv.Left + (x - StartX))
        dInv.Top = IIf(y < StartY, dInv.Top - (StartY - y), dInv.Top + (y - StartY))
    End If
End Sub

Private Sub InvTBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        If dInv.Left > dScreen.Width * 15 - dInv.Width Then dInv.Left = (dScreen.Width * 15 - dInv.Width)
        If dInv.Top > dScreen.Height * 15 - dInv.Height Then dInv.Top = (dScreen.Height * 15 - dInv.Height)
        If dInv.Left < 0 Then dInv.Left = 0
        If dInv.Top < 0 Then dInv.Top = 0
End Sub

Private Sub InvThumb_Click(Index As Integer)
'MsgBox Index
InvMnu.Visible = False
End Sub

Private Sub InvThumb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
  InvMnu.Move x + 120, y + 360
  InvMnu.Visible = True
End If
End Sub

Private Sub InvThumb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Index = 0 Then Exit Sub

If shSel.Visible = False Then shSel.Visible = True
shSel.Move InvThumb(Index).Left, InvThumb(Index).Top

InvDesc.Caption = GetInitEntry("Item", "Description", , WrkDir + Inventory(Index))
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
dScreen_KeyDown KeyCode, Shift
End Sub

Sub LoadRegion(wNum)
Dim yAr() As String

yAr = Split(DaRoom.Area(wNum, 3), ",")
ReDim PolyArea(UBound(yAr))
For a = 0 To UBound(yAr)
  PolyArea(a).x = Val(yAr(a))
Next a

yAr = Split(DaRoom.Area(wNum, 4), ",")
ReDim Preserve PolyArea(UBound(yAr))
For a = 0 To UBound(yAr)
  PolyArea(a).y = Val(yAr(a))
Next a

AddRGNPoly AllAreas, PolyArea, UBound(PolyArea), DaRoom.Area(wNum, 0), "", vbGreen, vbRed, RGN_HS_NOSHADE, RGN_BS_HOLLOW

End Sub

Sub LoadItmImg(dItm As String, Optional wX As Long, Optional wY As Long)
  Dim dGfx As String
  
  Load xItm(xItm.Count)
  
  dGfx = GetInitEntry("Item", "Graphic", , WrkDir + dItm)
  If Dir$(WrkDir + dGfx) <> "" And Trim$(dGfx) <> "" Then xItm(xItm.UBound).Picture = LoadPicture(WrkDir + dGfx)
  
  If wX <> 0 Then xItm(xItm.UBound).Left = wX
  If wY <> 0 Then xItm(xItm.UBound).Top = wY
  'xItm(xItm.UBound).ToolTipText = Left$(dItm, Len(dItm) - 4) & ": " & GetInitEntry("Item", "Description", , WrkDir + dItm)
  
  xItm(xItm.UBound).Visible = True
End Sub

Private Sub xItm_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If GScreen <> "game" Then Exit Sub

dMsg.Caption = DaRoom.Itm(Index, 1)
End Sub

Private Sub xItm_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If GScreen <> "game" Then Exit Sub

GScreen = "itemtalk"
  If DaRoom.Itm(Index, 2) <> "" Then ReadDialog DaRoom.Itm(Index, 2)
GScreen = "game"
End Sub

Sub ListInventory()

Do Until InvThumb.Count = 1
  Unload InvThumb(InvThumb.UBound)
Loop

'MsgBox GetInitEntry("Item", "Graphic", , WrkDir + Inventory(1))
'Exit Sub

For a = 1 To Inventory.Count
  Load InvThumb(a)
  InvThumb(a).Picture = LoadPicture(WrkDir + GetInitEntry("Item", "Graphic", , WrkDir + Inventory(a)))
  InvThumb(a).Visible = True
Next a

b = 0
For a = 1 To Inventory.Count Step 4
  InvThumb(a).Left = 8
  InvThumb(a).Top = (b * 80) + 8
  b = b + 1
Next a

b = 0
For a = 2 To Inventory.Count Step 4
  InvThumb(a).Left = 91
  InvThumb(a).Top = (b * 80) + 8
  b = b + 1
Next a

b = 0
For a = 3 To Inventory.Count Step 4
  InvThumb(a).Left = 173
  InvThumb(a).Top = (b * 80) + 8
  b = b + 1
Next a

b = 0
For a = 4 To Inventory.Count Step 4
  InvThumb(a).Left = 256
  InvThumb(a).Top = (b * 80) + 8
  b = b + 1
Next a

End Sub
