VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmConfigure 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Addons - Configure"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3570
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkUpdateProgram 
      Appearance      =   0  'Flat
      Caption         =   "Update Program"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkLoadNewAddonAfterAddonInstallation 
      Appearance      =   0  'Flat
      Caption         =   "Load New Addon"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CheckBox chkShowAddonsDock 
      Appearance      =   0  'Flat
      Caption         =   "Show Addons Dock"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   1815
   End
   Begin VB.CheckBox chkRestartProgramAfterAddonInstallation 
      Appearance      =   0  'Flat
      Caption         =   "Restart Program"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      Caption         =   "Download new Addons"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Update Addons"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   327682
      Enabled         =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "After new Addon Installation ..."
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "On Startup ..."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblPreformanceText 
      Alignment       =   2  'Center
      Caption         =   "Extremely Low (0%)"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Preformance:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1455
   End
End
Attribute VB_Name = "frmConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check3_Click()

End Sub

Private Sub cmdClose_Click()
On Local Error Resume Next
Unload Me
End Sub

