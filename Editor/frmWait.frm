VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Progr 
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label msg 
      AutoSize        =   -1  'True
      Caption         =   "Saving Game....please wait...."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2115
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
