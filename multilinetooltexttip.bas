Attribute VB_Name = "multilinetooltexttip"
Option Explicit

'************************************************************
' Constants
'************************************************************

Private Const GWL_WNDPROC = -4
Private Const GWL_STYLE = (-16)

Private Const WS_BORDER = &H800000

Private Const FW_NORMAL = 400
Private Const FW_HEAVY = 900
Private Const FW_SEMIBOLD = 600
Private Const FW_BLACK = FW_HEAVY
Private Const FW_BOLD = 700
Private Const FW_DEMIBOLD = FW_SEMIBOLD
Private Const FW_DONTCARE = 0
Private Const FW_EXTRABOLD = 800
Private Const FW_EXTRALIGHT = 200
Private Const FW_LIGHT = 300
Private Const FW_MEDIUM = 500
Private Const FW_REGULAR = FW_NORMAL
Private Const FW_THIN = 100
Private Const FW_ULTRABOLD = FW_EXTRABOLD
Private Const FW_ULTRALIGHT = FW_EXTRALIGHT

Private Const SW_SHOWNA = 8
Private Const TRANSPARENT = 1
Private Const ALTERNATE = 1
Private Const BLACK_BRUSH = 4
Private Const DKGRAY_BRUSH = 3

Private Const DT_SINGLELINE = &H20
Private Const DT_NOCLIP = &H100
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400

Private Const CW_USEDEFAULT = &H80000000

Private Const TTS_ALWAYSTIP = 1

Private Const TTF_IDISHWND = 1
Private Const TTF_CENTERTIP = 2
Private Const TTF_RTLREADING = 4
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK = &H20
Private Const TTF_ABSOLUTE = &H80
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_DI_SETITEM = &H8000

Private Const WM_USER = &H400
Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317

Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_SETDELAYTIME = WM_USER + 3
Private Const TTM_ADDTOOL = WM_USER + 4
Private Const TTM_DELTOOL = WM_USER + 5
Private Const TTM_NEWTOOLRECT = WM_USER + 6
Private Const TTM_RELAYEVENT = WM_USER + 7

Private Const LF_FACESIZE = 32
Private Const COLOR_INFOTEXT = 23
Private Const COLOR_INFOBK = 24
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_3DLIGHT = 22

Private Const RGN_OR = 2

'************************************************************
' API Functions
'************************************************************

Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" ( _
    ByVal dwExStyle As Long, ByVal lpClassName As String, _
    ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, _
    lpParam As Any) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, _
    lpRect As RECT) As Long

Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Long

Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function BeginPaint Lib "user32.dll" (ByVal hwnd As Long, _
    lpPaint As PAINTSTRUCT) As Long

Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" ( _
    ByVal hwnd As Long) As Long

Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" ( _
    ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function EndPaint Lib "user32.dll" (ByVal hwnd As Long, _
    lpPaint As PAINTSTRUCT) As Long

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, _
    lpRect As RECT) As Long

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" ( _
    ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, _
    lpRect As RECT, ByVal wFormat As Long) As Long

Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long

Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, _
    ByVal nBkMode As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" ( _
    lpLogFont As LOGFONT) As Long

Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, _
    ByVal crColor As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
    ByVal Y3 As Long) As Long

Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (lpPoint As POINTAPI, _
    ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long

Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, _
    ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, _
    ByVal nCombineMode As Long) As Long

Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, _
    ByVal hRgn As Long, ByVal hBrush As Long) As Long

Private Declare Function GetSysColorBrush Lib "user32.dll" ( _
    ByVal nIndex As Long) As Long

Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hdc As Long, _
    ByVal hRgn As Long, ByVal hBrush As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long

Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, _
    ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" ( _
    ByVal hObject As Long) As Long

'************************************************************
' Types
'************************************************************

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hwnd As Long
    uId As Long
    r As RECT
    hinst As Long
    lpszText As String
End Type

Private Type PAINTSTRUCT
    hdc As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(32) As Byte
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type
'************************************************************
' Variables and Constants
'************************************************************

Private Type TOldWndProc
    hwnd As Long
    lPrevWndProc As Long
End Type

Private WndProc() As TOldWndProc
Private NumTips As Long
Const iOffset = 8
Const FontType = "Tahoma" & vbNullChar
Const FontSize = 13

'*************************************************************
' Custom WindowProc for the ToolTip
'*************************************************************
Private Function CustomTipProc(ByVal hwnd As Long, ByVal uiMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim ps As PAINTSTRUCT
    Dim lpszText As String
    Dim iTextLen As Integer
    Dim rc As RECT
    Dim hFont As Long
    Dim hFontOld As Long
    Dim lf As LOGFONT
    Dim i As Integer
    Dim CurPos As POINTAPI

    Select Case uiMsg
    Case WM_PRINT
        PostMessage hwnd, WM_PAINT, 0, 0
        CustomTipProc = 1
    Case WM_PAINT
        ' Get the Current Window Rect
        GetWindowRect hwnd, rc
        GetCursorPos CurPos
        rc.Right = CurPos.x - iOffset + 6 + rc.Right - rc.Left
        rc.Bottom = CurPos.y + 20 + rc.Bottom - rc.Top
        rc.Left = CurPos.x - iOffset + 6
        rc.Top = CurPos.y + 20
        MoveWindow hwnd, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, False
        ' Get the Window Text (the ToolTip Text)
        iTextLen = GetWindowTextLength(hwnd) + 1
        lpszText = Space(iTextLen)
        GetWindowText hwnd, lpszText, iTextLen
        lpszText = Left(lpszText, Len(lpszText) - 1)
        ' prepare the DC for drawing
        BeginPaint hwnd, ps
        ' create and select the font to be used
        lf.lfHeight = FontSize
        lf.lfWeight = FW_NORMAL
        For i = 1 To Len(FontType)
            lf.lfFaceName(i) = Asc(Mid(FontType, i, 1))
        Next
        hFont = CreateFontIndirect(lf)
        hFontOld = SelectObject(ps.hdc, hFont)
        ' enlarge the window to exactly fit the size of the tooltip text

        ' using DT_CALCRECT the function extends the base of the
        ' rectangle to bound the last line of text but does not draw the text.
        DrawText ps.hdc, lpszText, Len(lpszText), rc, DT_VCENTER + DT_NOCLIP + DT_CALCRECT
        rc.Right = rc.Right + 2 * iOffset
        rc.Bottom = rc.Bottom + 3 * iOffset
        ' show the window before changing its size
        ' (work around the WM_PRINT problem/feature)
        ShowWindow hwnd, SW_SHOWNA
        ' apply new size
        MoveWindow hwnd, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, True
        SetBkMode ps.hdc, TRANSPARENT
        ' draw the balloon
        ToolTip_DrawBalloon hwnd, ps.hdc, lpszText
        ' Restore the Old Font
        SelectObject ps.hdc, hFontOld
        DeleteObject hFont
        ' End Paint
        EndPaint hwnd, ps
        CustomTipProc = 0
    Case Else
        ' Sends message to previous procedure
        For i = 0 To NumTips - 1
            If WndProc(i).hwnd = hwnd Then
                CustomTipProc = CallWindowProc(WndProc(i).lPrevWndProc, hwnd, uiMsg, _
                    wParam, lParam)
                Exit For
            End If
        Next
    End Select
End Function

Private Sub ToolTip_DrawBalloon(hwndTip As Long, hdc As Long, lpszText As String)
    Dim rc As RECT
    Dim hRgn, hrgn1, hrgn2 As Long
    Dim pts(0 To 2) As POINTAPI

    GetClientRect hwndTip, rc
    pts(0).x = rc.Left + iOffset
    pts(0).y = rc.Top
    pts(1).x = pts(0).x
    pts(1).y = pts(0).y + iOffset
    pts(2).x = pts(1).x + iOffset
    pts(2).y = pts(1).y
    hRgn = CreateRectRgn(0, 0, 0, 0)
    ' Create the rounded box
    hrgn1 = CreateRoundRectRgn(rc.Left, rc.Top + iOffset, rc.Right, rc.Bottom, 15, 15)
    ' Create the arrow
    hrgn2 = CreatePolygonRgn(pts(0), 3, ALTERNATE)
    ' combine the two regions
    CombineRgn hRgn, hrgn1, hrgn2, RGN_OR
    ' Fill the Region with the Standard BackColor of the ToolTip Window
    FillRgn hdc, hRgn, GetSysColorBrush(COLOR_INFOBK)
    ' Draw the Frame Region
    FrameRgn hdc, hRgn, GetStockObject(DKGRAY_BRUSH), 1, 1
    rc.Top = rc.Top + iOffset * 2
    rc.Bottom = rc.Bottom - iOffset
    rc.Left = rc.Left + iOffset
    rc.Right = rc.Right - iOffset
    ' Draw the Shadow Text
    SetTextColor hdc, GetSysColor(COLOR_3DLIGHT)
    DrawText hdc, lpszText, Len(lpszText), rc, DT_VCENTER + DT_NOCLIP
    rc.Left = rc.Left - 1
    rc.Top = rc.Top - 1
    ' Draw the Text
    SetTextColor hdc, GetSysColor(COLOR_INFOTEXT)
    DrawText hdc, lpszText, Len(lpszText), rc, DT_VCENTER + DT_NOCLIP
End Sub

' Add the Custom ToolTip to the specified object
Public Sub AddCustomToolTip(x As Object, ToolTipText As String, FormOwner As Form)
    Dim ti As TOOLINFO
    Dim dwStyle As Long
    Dim hTip As Long

    ' A tooltip control with the TTS_ALWAYSTIP style appears when the cursor is
    ' on a tool, regardless of whether the tooltip control's owner window is active
    ' or inactive. Without this style, the tooltip control appears when the tool's
    ' owner window is active, but not when it is inactive.
    hTip = CreateWindowEx(0&, "tooltips_class32", "", TTS_ALWAYSTIP, _
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, _
        FormOwner.hwnd, 0&, App.hInstance, 0&)
    ti.cbSize = Len(ti)
    ti.uFlags = TTF_IDISHWND + TTF_SUBCLASS
    ti.hwnd = x.hwnd
    ti.uId = x.hwnd
    ti.lpszText = ToolTipText
    SendMessage hTip, TTM_ADDTOOL, 0&, ti
    ' SubClass the tooltip window
    ReDim Preserve WndProc(NumTips)
    WndProc(NumTips).lPrevWndProc = SetWindowLong(hTip, GWL_WNDPROC, AddressOf CustomTipProc)
    WndProc(NumTips).hwnd = hTip
    NumTips = NumTips + 1
    ' Remove Border from ToolTip
    dwStyle = GetWindowLong(hTip, GWL_STYLE)
    dwStyle = dwStyle And (Not WS_BORDER)
    SetWindowLong hTip, GWL_STYLE, dwStyle
End Sub


