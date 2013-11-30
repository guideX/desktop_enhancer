VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Desktop Enhancer"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   2460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   2460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "This program was developed by |guideX| (Leon Aiossa) for Team Nexgen"
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Desktop Enhancer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "frmAbout.frx":57E2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
On Local Error Resume Next
Unload Me
End Sub
