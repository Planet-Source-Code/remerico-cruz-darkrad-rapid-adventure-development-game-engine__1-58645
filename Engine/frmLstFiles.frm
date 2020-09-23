VERSION 5.00
Begin VB.Form frmLstFiles 
   Caption         =   "Container"
   ClientHeight    =   3600
   ClientLeft      =   -8865
   ClientTop       =   45
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   3315
   Visible         =   0   'False
   Begin VB.ListBox lstVar 
      Height          =   645
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   3135
   End
   Begin VB.ListBox lstCode 
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
   Begin VB.ListBox lstLbl 
      Height          =   645
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.ListBox LstFiles 
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
End
Attribute VB_Name = "frmLstFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstVar_DblClick()
MsgBox DVar(lstVar.List(lstVar.ListIndex))
End Sub
