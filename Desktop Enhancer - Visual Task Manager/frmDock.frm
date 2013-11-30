VERSION 5.00
Begin VB.Form frmDock 
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmDock.frx":0000
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstNames 
      Height          =   255
      ItemData        =   "frmDock.frx":0CCA
      Left            =   480
      List            =   "frmDock.frx":0CCC
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   120
      Pattern         =   "*.png"
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   120
   End
End
Attribute VB_Name = "frmDock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ULW_ALPHA = &H2
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const HWND_TOPMOST = -1
Private Const GWL_EXSTYLE As Long = -20
Private Const SWP_NOSIZE As Long = &H1
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Type BITMAPINFOHEADER
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    SizeImage As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ClrUsed As Long
    ClrImportant As Long
End Type
Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type
Private Enum GDIPLUS_ALIGNMENT
   StringAlignmentNear = 0
   StringAlignmentCenter = 1
   StringAlignmentFar = 2
End Enum
Private Enum GDIPLUS_COLORS
   AliceBlue = &HFFF0F8FF
   AntiqueWhite = &HFFFAEBD7
   Aqua = &HFF00FFFF
   Aquamarine = &HFF7FFFD4
   Azure = &HFFF0FFFF
   Beige = &HFFF5F5DC
   Bisque = &HFFFFE4C4
   Black = &HFF000000
   BlanchedAlmond = &HFFFFEBCD
   Blue = &HFF0000FF
   BlueViolet = &HFF8A2BE2
   Brown = &HFFA52A2A
   BurlyWood = &HFFDEB887
   CadetBlue = &HFF5F9EA0
   Chartreuse = &HFF7FFF00
   Chocolate = &HFFD2691E
   Coral = &HFFFF7F50
   CornflowerBlue = &HFF6495ED
   Cornsilk = &HFFFFF8DC
   Crimson = &HFFDC143C
   Cyan = &HFF00FFFF
   DarkBlue = &HFF00008B
   DarkBrown = &HFF804040
   DarkCyan = &HFF008B8B
   DarkGoldenrod = &HFFB8860B
   DarkGray = &HFFA9A9A9
   DarkGreen = &HFF006400
   DarkKhaki = &HFFBDB76B
   DarkMagenta = &HFF8B008B
   DarkOliveGreen = &HFF556B2F
   DarkOrange = &HFFFF8C00
   DarkOrchid = &HFF9932CC
   DarkRed = &HFF8B0000
   DarkSalmon = &HFFE9967A
   DarkSeaGreen = &HFF8FBC8B
   DarkSlateBlue = &HFF483D8B
   DarkSlateGray = &HFF2F4F4F
   DarkTurquoise = &HFF00CED1
   DarkViolet = &HFF9400D3
   DeepPink = &HFFFF1493
   DeepSkyBlue = &HFF00BFFF
   DimGray = &HFF696969
   DodgerBlue = &HFF1E90FF
   Firebrick = &HFFB22222
   FloralWhite = &HFFFFFAF0
   ForestGreen = &HFF228B22
   Fuchsia = &HFFFF00FF
   Gainsboro = &HFFDCDCDC
   GhostWhite = &HFFF8F8FF
   Gold = &HFFFFD700
   Goldenrod = &HFFDAA520
   Gray = &HFF808080
   Green = &HFF008000
   GreenYellow = &HFFADFF2F
   Honeydew = &HFFF0FFF0
   HotPink = &HFFFF69B4
   IndianRed = &HFFCD5C5C
   Indigo = &HFF4B0082
   Ivory = &HFFFFFFF0
   Khaki = &HFFF0E68C
   Lavender = &HFFE6E6FA
   LavenderBlush = &HFFFFF0F5
   LawnGreen = &HFF7CFC00
   LemonChiffon = &HFFFFFACD
   LightBlue = &HFFADD8E6
   LightCoral = &HFFF08080
   LightCyan = &HFFE0FFFF
   LightGoldenrodYellow = &HFFFAFAD2
   LightGray = &HFFD3D3D3
   LightGreen = &HFF90EE90
   LightPink = &HFFFFB6C1
   LightSalmon = &HFFFFA07A
   LightSeaGreen = &HFF20B2AA
   LightSkyBlue = &HFF87CEFA
   LightSlateGray = &HFF778899
   LightSteelBlue = &HFFB0C4DE
   LightYellow = &HFFFFFFE0
   Lime = &HFF00FF00
   LimeGreen = &HFF32CD32
   Linen = &HFFFAF0E6
   Magenta = &HFFFF00FF
   Maroon = &HFF800000
   MediumAquamarine = &HFF66CDAA
   MediumBlue = &HFF0000CD
   MediumOrchid = &HFFBA55D3
   MediumPurple = &HFF9370DB
   MediumSeaGreen = &HFF3CB371
   MediumSlateBlue = &HFF7B68EE
   MediumSpringGreen = &HFF00FA9A
   MediumTurquoise = &HFF48D1CC
   MediumVioletRed = &HFFC71585
   MidnightBlue = &HFF191970
   MintCream = &HFFF5FFFA
   MistyRose = &HFFFFE4E1
   Moccasin = &HFFFFE4B5
   NavajoWhite = &HFFFFDEAD
   Navy = &HFF000080
   OldLace = &HFFFDF5E6
   Olive = &HFF808000
   OliveDrab = &HFF6B8E23
   Orange = &HFFFFA500
   OrangeRed = &HFFFF4500
   Orchid = &HFFDA70D6
   PaleGoldenrod = &HFFEEE8AA
   PaleGreen = &HFF98FB98
   PaleTurquoise = &HFFAFEEEE
   PaleVioletRed = &HFFDB7093
   PapayaWhip = &HFFFFEFD5
   PeachPuff = &HFFFFDAB9
   Peru = &HFFCD853F
   Pink = &HFFFFC0CB
   Plum = &HFFDDA0DD
   PowderBlue = &HFFB0E0E6
   Purple = &HFF800080
   Red = &HFFFF0000
   RosyBrown = &HFFBC8F8F
   RoyalBlue = &HFF4169E1
   SaddleBrown = &HFF8B4513
   Salmon = &HFFFA8072
   SandyBrown = &HFFF4A460
   SeaGreen = &HFF2E8B57
   SeaShell = &HFFFFF5EE
   Sienna = &HFFA0522D
   Silver = &HFFC0C0C0
   SkyBlue = &HFF87CEEB
   SlateBlue = &HFF6A5ACD
   SlateGray = &HFF708090
   Snow = &HFFFFFAFA
   SpringGreen = &HFF00FF7F
   SteelBlue = &HFF4682B4
   Tan = &HFFD2B48C
   Teal = &HFF008080
   Thistle = &HFFD8BFD8
   Tomato = &HFFFF6347
   Transparent = &HFFFFFF
   Turquoise = &HFF40E0D0
   Violet = &HFFEE82EE
   Wheat = &HFFF5DEB3
   White = &HFFFFFFFF
   WhiteSmoke = &HFFF5F5F5
   XPBlue = &HFF003CC7
   XPGradient = &HFFC6C5D7
   XPGoldDark = &HFFB08218
   XPGoldLight = &HFFFCF9C3
   Yellow = &HFFFFFF00
   YellowGreen = &HFF9ACD32
End Enum
Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum
Private Type GDIPLUS_STARTINPUT
    GDIPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Enum GDIPLUS_UNIT
    UnitWorld
    UnitDisplay
    UnitPixel
    UnitPoint
    UnitInch
    UnitDocument
    UnitMillimeter
End Enum
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECTF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Enum TASKBAR_POSITION
    vbBottom
    vbLeft
    vbRight
    vbTop
End Enum
Private Type BITMAPINFO
    bmpHeader As BITMAPINFOHEADER
    bmpColors As RGBQUAD
End Type
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As Long, ByVal emSize As Single, ByVal style As GDIPLUS_FONTSTYLE, ByVal UNIT As GDIPLUS_UNIT, createdfont As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal name As String, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hdc As Long, GpGraphics As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As GDIPLUS_COLORS, brush As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal graphics As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal Img As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal brush As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GDIPLUS_STARTINPUT, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipReleaseDC Lib "gdiplus.dll" (ByVal graphics As Long, ByVal hdc As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As GDIPLUS_ALIGNMENT) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As GDIPLUS_ALIGNMENT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Const bytMaxSize As Byte = 128
Private Const bytMinSize As Byte = 64
Private Const gdiBicubic = 7
Dim funcBlend32bpp As BLENDFUNCTION
Dim bmpInfo As BITMAPINFO
Dim dcMemory As Long, bmpMemory As Long
Dim lngHeight As Long, lngWidth As Long, lngBitmap As Long, lngImage As Long, lngGDI As Long, lngReturn As Long, lngCursor As Long
Dim sngIndex As Single, sngUBound As Single, sngStep As Single, sngStartTop As Single, sngStartLeft As Single
Dim lngFont As Long, lngBrush As Long, lngFontFamily As Long, lngCurrentFont As Long, lngFormat As Long
Dim sngHeight As Single, sngWidth As Single, sngLeft As Single, sngTop As Single, sngFrom() As Single
Dim apiWindow As POINTAPI, apiPoint As POINTAPI, apiMouse As POINTAPI
Dim bDrawn As Boolean, bHandle As Boolean
Dim gdiPosition As TASKBAR_POSITION
Dim gdipInit As GDIPLUS_STARTINPUT
Dim gdipColors As GDIPLUS_COLORS
Dim rctText As RECTF

Private Sub Form_Click()
''On Local Error Resume Next
'MsgBox File1.List(sngIndex)
'Dim i As Integer, c As Integer, f As Integer, mbox As VbMsgBoxResult, msg As String, msg2 As String, msg3 As String, msg4 As String, msg5 As String
Dim msg As String, msg2 As String, l As Long
Select Case sngIndex
Case 0
    EndProgram
Case Else
    msg = File1.List(sngIndex)
    l = CLng(Parse(msg, "(", ")"))
End Select
End Sub

Private Sub Form_Initialize()
'On Local Error Resume Next
Dim a As Integer, mbox As VbMsgBoxResult
'lSettings.sDockVisible = True
gdipInit.GDIPlusVersion = 1
If GdiplusStartup(lngGDI, gdipInit, ByVal 0&) <> 0 Then
    MsgBox "Error loading GDI+!", vbCritical
    Unload Me
End If
gdiPosition = vbTop
Me.Height = Screen.Height
Me.Width = Screen.Width
If Right(App.Path, 1) = "\" Then
    File1.Path = App.Path & "Icons\"
Else
    File1.Path = App.Path & "\Icons\"
End If
ReDim Preserve sngFrom(File1.ListCount)
lstNames.Clear
sngUBound = File1.ListCount - 1
If gdiPosition = vbBottom Then sngStartTop = (Me.Height / Screen.TwipsPerPixelX) - bytMaxSize
If gdiPosition = vbBottom Or gdiPosition = vbTop Then sngStartLeft = ((Me.Width / Screen.TwipsPerPixelX) - (File1.ListCount * bytMinSize)) / 2
If gdiPosition = vbLeft Or gdiPosition = vbRight Then sngStartTop = ((Me.Height / Screen.TwipsPerPixelY) - (File1.ListCount * bytMinSize)) / 2
If gdiPosition = vbRight Then sngStartLeft = (Me.Width / Screen.TwipsPerPixelX) - bytMaxSize
bmpInfo.bmpHeader.Size = Len(bmpInfo.bmpHeader)
bmpInfo.bmpHeader.BitCount = 32
bmpInfo.bmpHeader.Height = Me.ScaleHeight
bmpInfo.bmpHeader.Width = Me.ScaleWidth
bmpInfo.bmpHeader.Planes = 1
bmpInfo.bmpHeader.SizeImage = bmpInfo.bmpHeader.Width * bmpInfo.bmpHeader.Height * (bmpInfo.bmpHeader.BitCount / 8)
dcMemory = CreateCompatibleDC(Me.hdc)
bmpMemory = CreateDIBSection(dcMemory, bmpInfo, DIB_RGB_COLORS, ByVal 0, 0, 0)
SelectObject dcMemory, bmpMemory
GdipCreateFromHDC dcMemory, lngImage
If gdiPosition = vbBottom Or gdiPosition = vbTop Then sngLeft = sngStartLeft
If gdiPosition = vbLeft Or gdiPosition = vbRight Then sngTop = sngStartTop
For a = 0 To File1.ListCount - 1
    If IsNumeric(Left(File1.List(a), 2)) = True Then
        lstNames.AddItem Right(File1.List(a), Len(File1.List(a)) - 3)
    Else
        lstNames.AddItem File1.List(a)
    End If
    sngHeight = bytMinSize
    sngWidth = bytMinSize
    If gdiPosition = vbBottom Then sngTop = sngStartTop + bytMaxSize - bytMinSize
    If gdiPosition = vbRight Then sngLeft = sngStartLeft + bytMaxSize - bytMinSize
    LoadPictureGDIPlus File1.Path & "\" & File1.List(a), CLng(sngLeft), CLng(sngTop), CLng(sngWidth), CLng(sngHeight)
    If gdiPosition = vbBottom Or gdiPosition = vbTop Then
        sngFrom(a) = sngLeft * Screen.TwipsPerPixelX
        sngLeft = sngLeft + sngWidth
    ElseIf gdiPosition = vbLeft Or gdiPosition = vbRight Then
        sngFrom(a) = sngTop * Screen.TwipsPerPixelY
        sngTop = sngTop + sngHeight
    End If
Next a
If gdiPosition = vbBottom Or gdiPosition = vbTop Then sngFrom(UBound(sngFrom)) = sngLeft * Screen.TwipsPerPixelX
If gdiPosition = vbLeft Or gdiPosition = vbRight Then sngFrom(UBound(sngFrom)) = sngTop * Screen.TwipsPerPixelY
UpdateGDIPlus
'If lAddons.aCount = 0 Then
'    mbox = MsgBox("You have no Advanced Addons, would you like to search for new Addons now?", vbYesNo + vbQuestion)
'    Me.Visible = True
'End If
'If mbox = vbYes Then CheckForNewAddons
End Sub

Private Sub DrawText(strText As String, X As Single, Y As Single, Optional strFont As String = "Tahoma", Optional bytFontSize As Byte = 22, Optional bytBorderSize As Byte = 3)
'On Local Error Resume Next
GdipCreateFromHDC dcMemory, lngFont
GdipCreateSolidFill Black, lngBrush
GdipCreateFontFamilyFromName StrConv(strFont, vbUnicode), 0, lngFontFamily
GdipCreateFont lngFontFamily, bytFontSize, FontStyleBold, UnitPoint, lngCurrentFont
GdipCreateStringFormat 0, 0, lngFormat
GdipSetStringFormatAlign lngFormat, StringAlignmentCenter
GdipSetStringFormatLineAlign lngFormat, StringAlignmentNear
rctText.Left = Y - bytBorderSize
rctText.Top = X
rctText.Right = Me.ScaleWidth
rctText.Bottom = 36
GdipDrawString lngFont, StrConv(strText, vbUnicode), -1, lngCurrentFont, rctText, lngFormat, lngBrush
rctText.Left = Y + bytBorderSize
rctText.Top = X
rctText.Right = Me.ScaleWidth
rctText.Bottom = 36
GdipDrawString lngFont, StrConv(strText, vbUnicode), -1, lngCurrentFont, rctText, lngFormat, lngBrush
rctText.Left = Y
rctText.Top = X - bytBorderSize
rctText.Right = Me.ScaleWidth
rctText.Bottom = 36
GdipDrawString lngFont, StrConv(strText, vbUnicode), -1, lngCurrentFont, rctText, lngFormat, lngBrush
rctText.Left = Y
rctText.Top = X + bytBorderSize
rctText.Right = Me.ScaleWidth
rctText.Bottom = 36
GdipDrawString lngFont, StrConv(strText, vbUnicode), -1, lngCurrentFont, rctText, lngFormat, lngBrush
GdipCreateSolidFill White, lngBrush
rctText.Left = Y
rctText.Top = X
rctText.Right = Me.ScaleWidth
rctText.Bottom = 36
GdipDrawString lngFont, StrConv(strText, vbUnicode), -1, lngCurrentFont, rctText, lngFormat, lngBrush
GdipDeleteStringFormat lngFormat
GdipDeleteFont lngCurrentFont
GdipDeleteFontFamily lngFontFamily
GdipDeleteBrush lngBrush
GdipDeleteGraphics lngFont
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
If lngImage Then
    GdipReleaseDC lngImage, dcMemory
    GdipDeleteGraphics lngImage
End If
If lngBitmap Then GdipDisposeImage lngBitmap
If lngGDI Then GdiplusShutdown lngGDI
End Sub

Private Function RestoreGDIPlus()
'On Local Error Resume Next
DeleteObject bmpMemory
bmpMemory = CreateDIBSection(dcMemory, bmpInfo, DIB_RGB_COLORS, ByVal 0, 0, 0)
SelectObject dcMemory, bmpMemory
GdipCreateFromHDC dcMemory, lngImage
End Function

Private Function UpdateGDIPlus()
'On Local Error Resume Next
lngReturn = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
SetWindowLong Me.hWnd, GWL_EXSTYLE, lngReturn Or WS_EX_LAYERED
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE
apiPoint.X = 0
apiPoint.Y = 0
apiWindow.X = Screen.Width / Screen.TwipsPerPixelX
apiWindow.Y = Screen.Height / Screen.TwipsPerPixelY
funcBlend32bpp.AlphaFormat = AC_SRC_ALPHA
funcBlend32bpp.BlendFlags = 0
funcBlend32bpp.BlendOp = AC_SRC_OVER
funcBlend32bpp.SourceConstantAlpha = 255
GdipDisposeImage lngBitmap
GdipDeleteGraphics lngImage
UpdateLayeredWindow Me.hWnd, Me.hdc, ByVal 0&, apiWindow, dcMemory, apiPoint, 0, funcBlend32bpp, ULW_ALPHA
End Function

Private Function LoadPictureGDIPlus(strFilename As String, Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean
'On Local Error Resume Next
GdipLoadImageFromFile StrPtr(strFilename), lngBitmap
If Width = -1 Or Height = -1 Then
    GdipGetImageHeight lngBitmap, Height
    GdipGetImageWidth lngBitmap, Width
End If
GdipDrawImageRectI lngImage, lngBitmap, Left, Top, Width, Height
GdipDisposeImage lngBitmap
End Function

Private Sub Timer1_Timer()
'On Local Error Resume Next
Dim a As Integer
lngReturn = GetCursorPos(apiMouse)
If gdiPosition = vbBottom Then bHandle = apiMouse.Y < (Me.Height / Screen.TwipsPerPixelY) - bytMaxSize Or apiMouse.X < sngStartLeft Or apiMouse.X > sngFrom(UBound(sngFrom)) / Screen.TwipsPerPixelX
If gdiPosition = vbTop Then bHandle = apiMouse.Y > bytMaxSize Or apiMouse.X < sngStartLeft Or apiMouse.X > sngFrom(UBound(sngFrom)) / Screen.TwipsPerPixelX
If gdiPosition = vbRight Then bHandle = apiMouse.X < (Me.Width / Screen.TwipsPerPixelY) - bytMaxSize Or apiMouse.Y < sngStartTop Or apiMouse.Y > sngFrom(UBound(sngFrom)) / Screen.TwipsPerPixelX
If gdiPosition = vbLeft Then bHandle = apiMouse.X > bytMaxSize Or apiMouse.Y < sngStartTop Or apiMouse.Y > sngFrom(UBound(sngFrom)) / Screen.TwipsPerPixelX
If bHandle Then
    If bDrawn = False Then
        If gdiPosition = vbBottom Or gdiPosition = vbTop Then sngLeft = sngStartLeft
        If gdiPosition = vbLeft Or gdiPosition = vbRight Then sngTop = sngStartTop
        RestoreGDIPlus
        For a = 0 To File1.ListCount - 1
            sngHeight = bytMinSize
            sngWidth = bytMinSize
            If gdiPosition = vbBottom Then sngTop = sngStartTop + bytMaxSize - bytMinSize
            If gdiPosition = vbRight Then sngLeft = sngStartLeft + bytMaxSize - bytMinSize
            LoadPictureGDIPlus File1.Path & "\" & File1.List(a), CLng(sngLeft), CLng(sngTop), CLng(sngWidth), CLng(sngHeight)
            If gdiPosition = vbBottom Or gdiPosition = vbTop Then
                sngFrom(a) = sngLeft * Screen.TwipsPerPixelX
                sngLeft = sngLeft + sngWidth
            ElseIf gdiPosition = vbLeft Or gdiPosition = vbRight Then
                sngFrom(a) = sngTop * Screen.TwipsPerPixelY
                sngTop = sngTop + sngHeight
            End If
        Next a
        If gdiPosition = vbBottom Or gdiPosition = vbTop Then sngFrom(UBound(sngFrom)) = sngLeft * Screen.TwipsPerPixelX
        If gdiPosition = vbLeft Or gdiPosition = vbRight Then sngFrom(UBound(sngFrom)) = sngTop * Screen.TwipsPerPixelY
        UpdateGDIPlus
        bDrawn = True
    End If
    Exit Sub
End If
bDrawn = False
lngCursor = LoadCursor(0, 32512&)
If (lngCursor > 0) Then SetCursor lngCursor
For a = 0 To sngUBound
    If gdiPosition = vbBottom Or gdiPosition = vbTop Then bHandle = apiMouse.X >= sngFrom(a) / Screen.TwipsPerPixelX And apiMouse.X <= sngFrom(a + 1) / Screen.TwipsPerPixelX
    If gdiPosition = vbLeft Or gdiPosition = vbRight Then bHandle = apiMouse.Y >= sngFrom(a) / Screen.TwipsPerPixelY And apiMouse.Y <= sngFrom(a + 1) / Screen.TwipsPerPixelY
    If bHandle Then
        sngIndex = a
        Exit For
    End If
Next a
If gdiPosition = vbBottom Or gdiPosition = vbTop Then
    sngLeft = sngStartLeft
    sngStep = ((apiMouse.X * Screen.TwipsPerPixelX) - sngFrom(sngIndex)) / (2 * Screen.TwipsPerPixelX)
ElseIf gdiPosition = vbLeft Or gdiPosition = vbRight Then
    sngTop = sngStartTop
    sngStep = ((apiMouse.Y * Screen.TwipsPerPixelX) - sngFrom(sngIndex)) / (2 * Screen.TwipsPerPixelY)
End If
RestoreGDIPlus
For a = 0 To sngUBound
    If a <> sngIndex And (a <> sngIndex - 1 And a <> sngIndex + 1) Then
        sngHeight = bytMinSize
        sngWidth = bytMinSize
        If gdiPosition = vbBottom Then sngTop = sngStartTop + bytMaxSize - bytMinSize
        If gdiPosition = vbRight Then sngLeft = sngStartLeft + bytMaxSize - bytMinSize
    ElseIf a = sngIndex - 1 Then
        sngHeight = bytMaxSize - sngStep
        sngWidth = bytMaxSize - sngStep
        If gdiPosition = vbBottom Then sngTop = sngStartTop + bytMaxSize - (bytMaxSize - sngStep)
        If gdiPosition = vbRight Then sngLeft = sngStartLeft + bytMaxSize - (bytMaxSize - sngStep)
    ElseIf a = sngIndex Then
        sngHeight = bytMaxSize
        sngWidth = bytMaxSize
        If gdiPosition = vbBottom Then sngTop = sngStartTop + bytMaxSize - bytMaxSize
        If gdiPosition = vbRight Then sngLeft = sngStartLeft + bytMaxSize - bytMaxSize
    ElseIf a = sngIndex + 1 Then
        sngHeight = bytMinSize + sngStep
        sngWidth = bytMinSize + sngStep
        If gdiPosition = vbBottom Then sngTop = sngStartTop + bytMaxSize - (bytMinSize + sngStep)
        If gdiPosition = vbRight Then sngLeft = sngStartLeft + bytMaxSize - (bytMinSize + sngStep)
    End If
    DrawText Left(lstNames.List(sngIndex), Len(lstNames.List(sngIndex)) - 4), sngTop + bytMaxSize, 0, "Arial", 24, 1
    LoadPictureGDIPlus File1.Path & "\" & File1.List(a), CLng(sngLeft), CLng(sngTop), CLng(sngWidth), CLng(sngHeight)
    If gdiPosition = vbBottom Or gdiPosition = vbTop Then
        sngFrom(a) = sngLeft * Screen.TwipsPerPixelX
        sngLeft = sngLeft + sngWidth
    ElseIf gdiPosition = vbLeft Or gdiPosition = vbRight Then
        sngFrom(a) = sngTop * Screen.TwipsPerPixelY
        sngTop = sngTop + sngHeight
    End If
Next a
If gdiPosition = vbBottom Or gdiPosition = vbTop Then sngFrom(UBound(sngFrom)) = sngLeft * Screen.TwipsPerPixelX
If gdiPosition = vbLeft Or gdiPosition = vbRight Then sngFrom(UBound(sngFrom)) = sngTop * Screen.TwipsPerPixelY
UpdateGDIPlus
End Sub
