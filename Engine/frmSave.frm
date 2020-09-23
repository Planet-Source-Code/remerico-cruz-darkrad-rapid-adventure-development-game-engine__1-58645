VERSION 5.00
Begin VB.Form frmSave 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Dialog Caption"
   ClientHeight    =   5790
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgRm 
      Height          =   1065
      Index           =   3
      Left            =   4755
      Stretch         =   -1  'True
      Top             =   3255
      Width           =   1845
   End
   Begin VB.Image imgRm 
      Height          =   1065
      Index           =   2
      Left            =   4755
      Stretch         =   -1  'True
      Top             =   2055
      Width           =   1845
   End
   Begin VB.Image imgRm 
      Height          =   1065
      Index           =   1
      Left            =   4755
      Stretch         =   -1  'True
      Top             =   855
      Width           =   1845
   End
   Begin VB.Image imgRm 
      Height          =   1065
      Index           =   4
      Left            =   4755
      Stretch         =   -1  'True
      Top             =   4455
      Width           =   1845
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
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
      Height          =   240
      Left            =   300
      TabIndex        =   6
      Top             =   495
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Shape SaveSlot 
      BackColor       =   &H00500000&
      BorderColor     =   &H00FFC0C0&
      Height          =   375
      Index           =   5
      Left            =   240
      Top             =   420
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Height          =   240
      Left            =   5430
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape SaveSlot 
      BackColor       =   &H00500000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      Height          =   375
      Index           =   0
      Left            =   5400
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Save Slot"
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
      Left            =   120
      TabIndex        =   4
      Top             =   75
      Width           =   3240
   End
   Begin VB.Label EmptySL 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Empty Slot"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   3
      Top             =   4560
      Width           =   1365
   End
   Begin VB.Label EmptySL 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Empty Slot"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Width           =   1365
   End
   Begin VB.Label EmptySL 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Empty Slot"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   1365
   End
   Begin VB.Label EmptySL 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Empty Slot"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1365
   End
   Begin VB.Shape SaveSlot 
      BackColor       =   &H00500000&
      BorderColor     =   &H00FFC0C0&
      Height          =   1095
      Index           =   4
      Left            =   240
      Top             =   4440
      Width           =   6375
   End
   Begin VB.Shape SaveSlot 
      BackColor       =   &H00500000&
      BorderColor     =   &H00FFC0C0&
      Height          =   1095
      Index           =   3
      Left            =   240
      Top             =   3240
      Width           =   6375
   End
   Begin VB.Shape SaveSlot 
      BackColor       =   &H00500000&
      BorderColor     =   &H00FFC0C0&
      Height          =   1095
      Index           =   2
      Left            =   240
      Top             =   2040
      Width           =   6375
   End
   Begin VB.Shape SaveSlot 
      BackColor       =   &H00500000&
      BorderColor     =   &H00FFC0C0&
      Height          =   1095
      Index           =   1
      Left            =   240
      Top             =   840
      Width           =   6375
   End
   Begin VB.Shape shape 
      BorderColor     =   &H00FFC0C0&
      Height          =   5295
      Left            =   120
      Top             =   360
      Width           =   6615
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CSlot As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 27   ' ESC
    uSlot = -1
    Unload Me
  Case 38   ' UP
    SaveSlot(CSlot).BackStyle = 0
    CSlot = CSlot - 1
    If CSlot = -1 Then CSlot = 5
    If SaveMode = "save" And CSlot = 5 Then CSlot = 4
    SaveSlot(CSlot).BackStyle = 1
  Case 40   ' DOWN
    SaveSlot(CSlot).BackStyle = 0
    CSlot = CSlot + 1
    If SaveMode = "save" And CSlot = 5 Then CSlot = 0
    If CSlot = 6 Then CSlot = 0
    SaveSlot(CSlot).BackStyle = 1
  Case 13  ' ENTER
    Select Case CSlot
      Case 0
        uSlot = CSlot
        Unload Me
      Case 1, 2, 3, 4
        SaveGame App.Path & "\" & GName & CSlot & "." & SaveExt
        Unload Me
    End Select
End Select
End Sub

Private Sub Form_Load()
CSlot = 0
PrevSave
End Sub

Private Sub Label2_Click()
uSlot = -1
Unload Me
End Sub

Sub PrevSave()
For a = 1 To 4
  If Dir$(App.Path & "\" & GName & a & "." & SaveExt) <> "" Then
    EmptySL(a).Caption = "Slot " & a & " - " & GetInitEntry("Game", "Room", "", App.Path & "\" & GName & a & "." & SaveExt)
    EmptySL(a).Caption = Left$(EmptySL(a).Caption, Len(EmptySL(a).Caption) - 3)
    
    If GetRoomImg(GetInitEntry("Game", "Room", "", App.Path & "\" & GName & a & "." & SaveExt)) <> "" Then
      imgRm(a) = LoadPicture(WrkDir + GetRoomImg(GetInitEntry("Game", "Room", "", App.Path & "\" & GName & a & "." & SaveExt)))
    End If
  End If
Next a
End Sub

Function GetRoomImg(RoomFile) As String
' Gets the image filename of the Room specified for quick viewing
If GetInitEntry("Room", "BG", "", WrkDir + RoomFile) <> "" Then GetRoomImg = GetInitEntry("Room", "BG", "", WrkDir + RoomFile)
End Function
