VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreenResolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const DM_DISPLAYFLAGS = &H200000
Private Const DM_DISPLAYFREQUENCY = &H400000
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpInitData As DEVMODE, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (lpszDeviceName As Any, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const BITSPIXEL = 12
Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H2
Private Const CDS_FULLSCREEN = &H4
Private Const CDS_GLOBAL = &H8
Private Const CDS_SET_PRIMARY = &H10
Private Const CDS_RESET = &H40000000
Private Const CDS_SETRECT = &H20000000
Private Const CDS_NORESET = &H10000000
Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1
Private Const DISP_CHANGE_FAILED = -1
Private Const DISP_CHANGE_BADMODE = -2
Private Const DISP_CHANGE_NOTUPDATED = -3
Private Const DISP_CHANGE_BADFLAGS = -4
Private Const DISP_CHANGE_BADPARAM = -5
Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4
Private D() As DEVMODE, lNumModes As Long

Public Sub LoadResolutionInfoIntoListBox(lListBox As ListBox)
On Local Error GoTo ErrHandler
Dim l As Long, lMaxModes As Long, lBits As Long, lWidth As Long, lHeight As Long, hdc
lBits = GetDeviceCaps(hdc, BITSPIXEL)
lWidth = Screen.Width \ Screen.TwipsPerPixelX
lHeight = Screen.Height \ Screen.TwipsPerPixelY
lMaxModes = 8
ReDim D(0 To lMaxModes) As DEVMODE
lNumModes = 0
l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
Do While l
    lListBox.AddItem D(lNumModes).dmPelsWidth & "x" & D(lNumModes).dmPelsHeight & "x" & D(lNumModes).dmBitsPerPel
    If lBits = D(lNumModes).dmBitsPerPel And lWidth = D(lNumModes).dmPelsWidth And lHeight = D(lNumModes).dmPelsHeight Then lListBox.ListIndex = lListBox.NewIndex
    lNumModes = lNumModes + 1
    If lNumModes > lMaxModes Then
        lMaxModes = lMaxModes + 8
        ReDim Preserve D(0 To lMaxModes) As DEVMODE
    End If
    l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
Loop
lNumModes = lNumModes - 1
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Public Function ChangeResolution(lIndex As Long) As Boolean
On Local Error GoTo ErrHandler
Dim l As Long, Flags As Long, x As Long
x = lIndex
D(x).dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
Flags = CDS_UPDATEREGISTRY
l = ChangeDisplaySettings(D(x), Flags)
Select Case l
Case DISP_CHANGE_RESTART
    l = MsgBox("This screen resolution change will not take effect until you reboot Windows. Would you like to reboot Windows now?", vbYesNo)
    If l = vbYes Then
        Flags = 0
        l = ExitWindowsEx(EWX_REBOOT, Flags)
    End If
Case DISP_CHANGE_SUCCESSFUL
Case Else
    MsgBox "Error changing resolution! " & l
End Select
Exit Function
ErrHandler:
    MsgBox Err.Description
End Function
