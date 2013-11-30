VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   480
   ClientLeft      =   8505
   ClientTop       =   885
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   480
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgAddon 
      Height          =   345
      Index           =   9
      Left            =   5040
      Top             =   75
      Width           =   345
   End
   Begin VB.Image imgAddon 
      Height          =   345
      Index           =   8
      Left            =   4560
      Top             =   80
      Width           =   345
   End
   Begin VB.Image imgAddon 
      Height          =   345
      Index           =   7
      Left            =   4080
      Top             =   80
      Width           =   345
   End
   Begin VB.Image imgAddon 
      Height          =   345
      Index           =   6
      Left            =   3600
      Top             =   80
      Width           =   345
   End
   Begin VB.Image imgAddon 
      Height          =   345
      Index           =   5
      Left            =   3120
      Top             =   80
      Width           =   345
   End
   Begin VB.Image imgAddon 
      Height          =   345
      Index           =   4
      Left            =   2640
      Top             =   80
      Width           =   345
   End
   Begin VB.Image imgAddon 
      Height          =   345
      Index           =   3
      Left            =   2160
      Top             =   80
      Width           =   345
   End
   Begin VB.Image imgAddon 
      Height          =   345
      Index           =   2
      Left            =   1680
      Top             =   80
      Width           =   345
   End
   Begin VB.Image imgAddon 
      Height          =   345
      Index           =   1
      Left            =   1200
      Top             =   80
      Width           =   345
   End
   Begin VB.Image imgAddon 
      Height          =   345
      Index           =   0
      Left            =   720
      Top             =   80
      Width           =   345
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
SetParentAcidmax Me
frmDock.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then FormDrag Me
End Sub
