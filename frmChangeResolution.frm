VERSION 5.00
Begin VB.Form frmChangeResolution 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resolution"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   1920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChangeResolution.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ListBox lstScreenResolutions 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2580
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   4200
      Left            =   4800
      Picture         =   "frmChangeResolution.frx":57E2
      Top             =   1920
      Width           =   300
   End
End
Attribute VB_Name = "frmChangeResolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lResolution As clsScreenResolution

Private Sub cmdSwitch_Click()
On Local Error Resume Next
lResolution.ChangeResolution lstScreenResolutions.ListIndex
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Set lResolution = New clsScreenResolution
lResolution.LoadResolutionInfoIntoListBox lstScreenResolutions
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
Set lResolution = Nothing
End Sub
