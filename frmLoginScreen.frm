VERSION 5.00
Begin VB.Form frmLoginScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picUsers 
      AutoSize        =   -1  'True
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
      Height          =   1890
      Index           =   0
      Left            =   0
      Picture         =   "frmLoginScreen.frx":0000
      ScaleHeight     =   1890
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3225
   End
End
Attribute VB_Name = "frmLoginScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lUserCount As Integer
Private lUserNick(1 To 256) As String
Private lUserPassword(1 To 256) As String

Public Sub LoadUsers()
'On Local Error Resume Next
Dim i As Integer

End Sub

Public Sub AddUser(lNickName As String, lPassword As String)
'On Local Error Resume Next
If Len(lNickName) <> 0 And Len(lPassword) <> 0 Then
    lUserCount = lUserCount + 1
    Load picUsers(lUserCount)
    'Load lblNickname(lUserCount)
    With picUsers(lUserCount)
        .Visible = True
        '.Left = Me.ScaleWidth / 2 + (picUsers(lUserCount).Width)
        Select Case lUserCount
        Case 0
        Case 1
        Case Else
            .Top = ((picUsers(lUserCount).Height * lUserCount - 1) - 1900)
        End Select
    End With
    'With lblNickname(lUserCount)
        '.Left = 0
        '.Top = 0
        '.Caption = lNickName
        '.ForeColor = vbWhite
        '.Visible = True
    'End With
End If
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
AddUser "guidex", "dietpepsi"
'AddUser "guidex", "dietpepsi"
End Sub

Private Sub Label1_Click()

End Sub
