VERSION 5.00
Begin VB.Form frmCmd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Command"
   ClientHeight    =   3195
   ClientLeft      =   2880
   ClientTop       =   3765
   ClientWidth     =   6030
   Icon            =   "frmCmd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox tBlank 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   22
      Top             =   120
      Width           =   2535
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Please select a command on the left."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   51
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.PictureBox tWin 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   95
      Top             =   120
      Width           =   2535
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Win"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   98
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         Caption         =   "Terminate the game and return to Windows."
         Height          =   495
         Left            =   0
         TabIndex        =   97
         Top             =   600
         Width           =   2535
      End
      Begin VB.Line Line15 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No Parameters in this command."
         Height          =   195
         Left            =   240
         TabIndex        =   96
         Top             =   1560
         Width           =   2280
      End
   End
   Begin VB.PictureBox tEnd 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   59
      Top             =   120
      Width           =   2535
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No Parameters in this command."
         Height          =   195
         Left            =   240
         TabIndex        =   62
         Top             =   1560
         Width           =   2280
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "Ends the execution of the current dialog"
         Height          =   495
         Left            =   0
         TabIndex        =   61
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   60
         Top             =   120
         Width           =   705
      End
   End
   Begin VB.PictureBox tGotoLine 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   37
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtGoToLine 
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "GoToLine"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   42
         Top             =   120
         Width           =   1470
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Directs the flow of the dialog to the specified line label"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label22 
         Caption         =   "Line label where the dialog will continue:"
         Height          =   375
         Left            =   0
         TabIndex        =   40
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label21 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   39
         Top             =   1650
         Width           =   135
      End
   End
   Begin VB.ListBox lstCmd 
      Height          =   2985
      ItemData        =   "frmCmd.frx":038A
      Left            =   120
      List            =   "frmCmd.frx":03BE
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox tTrans 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   52
      Top             =   120
      Width           =   2535
      Begin VB.ComboBox cmbTransGfx 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox cmbTrans 
         Height          =   315
         ItemData        =   "frmCmd.frx":0447
         Left            =   0
         List            =   "frmCmd.frx":0463
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   56
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label32 
         Caption         =   "Transition:"
         Height          =   255
         Left            =   0
         TabIndex        =   55
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Graphic: (blank for fade out effect)"
         Height          =   195
         Left            =   0
         TabIndex        =   54
         Top             =   1200
         Width           =   2430
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Change the graphics displayed with a screen transition."
         Height          =   495
         Left            =   0
         TabIndex        =   53
         Top             =   600
         Width           =   2535
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.PictureBox tWait 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   72
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtWaitMSeconds 
         Height          =   285
         Left            =   0
         TabIndex        =   73
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wait"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   76
         Top             =   120
         Width           =   855
      End
      Begin VB.Line Line12 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "The computer will wait for a given time in milliseconds."
         Height          =   615
         Left            =   120
         TabIndex        =   75
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Milliseconds:"
         Height          =   195
         Left            =   0
         TabIndex        =   74
         Top             =   1320
         Width           =   900
      End
   End
   Begin VB.PictureBox tAddInventory 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   77
      Top             =   120
      Width           =   2535
      Begin VB.ComboBox cmbAddInv 
         Height          =   315
         Left            =   0
         TabIndex        =   81
         Text            =   "cmbAddInv"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   195
         Left            =   0
         TabIndex        =   80
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         Caption         =   "Adds the specified item to the inventory."
         Height          =   615
         Left            =   120
         TabIndex        =   79
         Top             =   600
         Width           =   2415
      End
      Begin VB.Line Line13 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AddInventory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   78
         Top             =   120
         Width           =   2265
      End
   End
   Begin VB.PictureBox tRun 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   63
      Top             =   120
      Width           =   2535
      Begin VB.ComboBox cmbRunDg 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "End the current dialog and run anoher one."
         Height          =   495
         Left            =   0
         TabIndex        =   67
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Dialog:"
         Height          =   195
         Left            =   0
         TabIndex        =   66
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Run"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   65
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox tStopMusic 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   68
      Top             =   120
      Width           =   2535
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "StopMusic"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   71
         Top             =   120
         Width           =   1740
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Stop playing the current music."
         Height          =   255
         Left            =   0
         TabIndex        =   70
         Top             =   600
         Width           =   2535
      End
      Begin VB.Line Line11 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No Parameters in this command."
         Height          =   195
         Left            =   240
         TabIndex        =   69
         Top             =   1560
         Width           =   2280
      End
   End
   Begin VB.PictureBox tChoice 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   29
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtCondLabel 
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtCondition 
         Height          =   285
         Left            =   0
         TabIndex        =   33
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label20 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   35
         Top             =   2480
         Width           =   135
      End
      Begin VB.Label Label19 
         Caption         =   "Line label where the dialog will continue:"
         Height          =   375
         Left            =   0
         TabIndex        =   34
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Choice text:"
         Height          =   195
         Left            =   0
         TabIndex        =   31
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Adds a selection within the dialog and directs the flow of the dialog to the specified line label"
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   2415
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Choice"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   32
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.PictureBox tPrint 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   18
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtPrintMsg 
         Height          =   285
         Left            =   0
         TabIndex        =   23
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   21
         Top             =   120
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Message:"
         Height          =   195
         Left            =   0
         TabIndex        =   20
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Displays the specified message to the navigation bar."
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.PictureBox tGotoRoom 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   13
      Top             =   120
      Width           =   2535
      Begin VB.ComboBox cmbGRoom 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Go to the room specified."
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2415
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Room:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "GotoRoom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   120
         Width           =   1680
      End
   End
   Begin VB.PictureBox tDisplayGfx 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   2040
      ScaleHeight     =   3015
      ScaleWidth      =   2535
      TabIndex        =   3
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtDGfxY 
         Height          =   285
         Left            =   480
         TabIndex        =   12
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtDGfxX 
         Height          =   285
         Left            =   480
         TabIndex        =   10
         Top             =   2240
         Width           =   855
      End
      Begin VB.ComboBox cmbDGfx 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2565
         Width           =   150
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   150
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Optional:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Graphic:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   600
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Displays a graphic at the specified x and y coordinate"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DisplayGfx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   1710
      End
   End
   Begin VB.PictureBox tIf 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   82
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox chkVar2 
         Caption         =   "Variable"
         Height          =   255
         Left            =   1560
         TabIndex        =   94
         Top             =   2190
         Width           =   975
      End
      Begin VB.CheckBox chkVar1 
         Caption         =   "Variable"
         Height          =   255
         Left            =   1560
         TabIndex        =   93
         Top             =   1220
         Width           =   975
      End
      Begin VB.TextBox txtIfVal2 
         Height          =   285
         Left            =   0
         TabIndex        =   92
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtIfLLabel 
         Height          =   285
         Left            =   120
         TabIndex        =   90
         Top             =   2640
         Width           =   2415
      End
      Begin VB.ComboBox cmbIfCond 
         Height          =   315
         ItemData        =   "frmCmd.frx":04D0
         Left            =   0
         List            =   "frmCmd.frx":04E6
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtIfVal1 
         Height          =   285
         Left            =   0
         TabIndex        =   86
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label54 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   91
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "...then goto this line label..."
         Height          =   195
         Left            =   0
         TabIndex        =   89
         Top             =   2400
         Width           =   1890
      End
      Begin VB.Label Label50 
         Caption         =   "...is..."
         Height          =   255
         Left            =   0
         TabIndex        =   87
         Top             =   1320
         Width           =   615
      End
      Begin VB.Line Line14 
         X1              =   0
         X2              =   2520
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Compares two values and evaluates them."
         Height          =   495
         Left            =   0
         TabIndex        =   85
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "If....."
         Height          =   195
         Left            =   0
         TabIndex        =   84
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   83
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox tSetVar 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   43
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox chkSVar 
         Caption         =   "This is a Variable"
         Height          =   195
         Left            =   0
         TabIndex        =   50
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtSVar 
         Height          =   285
         Left            =   0
         TabIndex        =   45
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtSVarVal 
         Height          =   285
         Left            =   0
         TabIndex        =   44
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Sets a value for the variable specified."
         Height          =   495
         Left            =   0
         TabIndex        =   48
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Variable:"
         Height          =   195
         Left            =   0
         TabIndex        =   47
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "Value:"
         Height          =   255
         Left            =   0
         TabIndex        =   46
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SetVar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   49
         Top             =   120
         Width           =   1170
      End
   End
   Begin VB.PictureBox tPlayMusic 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   24
      Top             =   120
      Width           =   2535
      Begin VB.ComboBox cmbPlayMusic 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Plays the specified music in the background."
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Music:"
         Height          =   195
         Left            =   0
         TabIndex        =   26
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "PlayMusic"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   25
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox tMusicVol 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   2040
      ScaleHeight     =   2955
      ScaleWidth      =   2535
      TabIndex        =   99
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtMscVol 
         Height          =   285
         Left            =   0
         TabIndex        =   103
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "Music Volume:"
         Height          =   195
         Left            =   0
         TabIndex        =   101
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         Caption         =   "Plays the specified music in the background."
         Height          =   495
         Left            =   120
         TabIndex        =   100
         Top             =   600
         Width           =   2415
      End
      Begin VB.Line Line16 
         X1              =   0
         X2              =   2520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "MusicVol"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   102
         Top             =   120
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
FillGfx cmbDGfx
FillGfx cmbTransGfx
FillDialog cmbRunDg
FillMusic cmbPlayMusic
FillRoom cmbGRoom
FillItem cmbAddInv
End Sub



Private Sub lstCmd_Click()
Select Case lstCmd.Text
  Case "AddInventory": tAddInventory.ZOrder
  Case "Choice": tChoice.ZOrder
  Case "DisplayGfx": tDisplayGfx.ZOrder
  Case "End": tEnd.ZOrder
  Case "GoToLine": tGotoLine.ZOrder
  Case "GotoRoom": tGotoRoom.ZOrder
  Case "If": tIf.ZOrder
  Case "MusicVol": tMusicVol.ZOrder
  Case "Print": tPrint.ZOrder
  Case "PlayMusic": tPlayMusic.ZOrder
  Case "Run": tRun.ZOrder
  Case "SetVar": tSetVar.ZOrder
  Case "StopMusic": tStopMusic.ZOrder
  Case "Trans": tTrans.ZOrder
  Case "Wait": tWait.ZOrder
  Case "Win": tWin.ZOrder
End Select
End Sub

Private Sub OKButton_Click()
' Generates code here

With frmDialog.ConvLst
Select Case lstCmd.Text
  Case "AddInventory"
    .AddItem "#AddInventory(" + Chr(34) + cmbAddInv.Text + Chr(34) + ")"

  Case "Choice"
    .AddItem "#Choice(" + Chr(34) + txtCondition.Text + Chr(34) + "," + txtCondLabel.Text + ")"
  
  Case "DisplayGfx"
    .AddItem "#DisplayGfx(" + Chr(34) + cmbDGfx.Text + Chr(34) + "," + txtDGfxX + "," + txtDGfxY + ")"
    
  Case "End"
    .AddItem "#End()"
  
  Case "GoToLine"
    .AddItem "#GoToLine(" + txtGoToLine.Text + ")"
  
  Case "GotoRoom"
    .AddItem "#GotoRoom(" + Chr(34) + cmbGRoom.Text + Chr(34) + ")"
    
  Case "If"
    Dim EqType As String
    
    Select Case cmbIfCond.ListIndex
      Case 0
        EqType = "=="
      Case 1
        EqType = "~="
      Case 2
        EqType = "<"
      Case 3
        EqType = ">"
      Case 4
         EqType = "<="
      Case 5
         EqType = ">="
    End Select
    
    If chkVar1.Value = vbUnchecked Then txtIfVal1.Text = Chr(34) + txtIfVal1.Text + Chr(34)
    If chkVar2.Value = vbUnchecked Then txtIfVal2.Text = Chr(34) + txtIfVal2.Text + Chr(34)
      
    .AddItem "#If(" + txtIfVal1.Text + " " + EqType + " " + txtIfVal2.Text + "," + txtIfLLabel.Text + ")"
    
  Case "MusicVol"
    .AddItem "#MusicVol " + txtMscVol.Text
  
  Case "Print"
    .AddItem "#Print(" + Chr(34) + txtPrintMsg.Text + Chr(34) + ")"
  
  Case "PlayMusic"
    .AddItem "#PlayMusic(" + Chr(34) + cmbPlayMusic.Text + Chr(34) + ")"
  
  Case "Run"
    .AddItem "#Run(" + Chr(34) + cmbRunDg.Text + Chr(34) + ")"
  
  Case "SetVar"
    If chkSVar.Value = vbUnchecked Then
      .AddItem "#SetVar(" + txtSVar.Text + "," + Chr(34) + txtSVarVal.Text + Chr(34) + ")"
    Else
      .AddItem "#SetVar(" + txtSVar.Text + "," + txtSVarVal.Text + ")"
    End If
    
  Case "StopMusic"
    .AddItem "#StopMusic()"
  
  Case "Trans"
    .AddItem "#Trans(" + Chr(34) + cmbTransGfx.Text + Chr(34) + "," + Left$(cmbTrans.Text, 1) + ")"
    
  Case "Wait"
    .AddItem "#Wait(" + Chr(34) + txtWaitMSeconds.Text + Chr(34) + ")"
    
  Case "Win"
    .AddItem "#Win()"
End Select
End With

Unload Me
End Sub

Private Sub txtDGfxX_Change()
If IsNumeric(txtDGfxX.Text) = False Then MsgBox "Only numberic values are allowed.": txtDGfxX.Text = ""
End Sub

Private Sub txtDGfxY_Change()
If IsNumeric(txtDGfxY.Text) = False Then MsgBox "Only numberic values are allowed.": txtDGfxY.Text = ""
End Sub
