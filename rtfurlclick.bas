Attribute VB_Name = "rtfurlclick"

Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
'// notification structures
Public Type NMHDR_RICHEDIT
    hwndFrom As Long
    wPad1 As Integer
    idfrom As Integer
    code As Integer
    wPad2 As Integer
End Type

Public Type ENLINK
    NMHDR As NMHDR_RICHEDIT
    msg As Integer
    wPad1 As Integer
    wParam As Integer
    wPad2 As Integer
    lParam As Integer
    chrg As CHARRANGE
End Type
'// events and messages
Public Const ENM_LINK = &H4000000
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEMOVE = &H200

Public Enum ERECLinkEventTypeCOnstants
   ercLButtonDblClick = WM_LBUTTONDBLCLK
   ercLButtonDown = WM_LBUTTONDOWN
   ercLButtonUp = WM_LBUTTONUP
   ercMouseMove = WM_MOUSEMOVE
   ercRButtonDblClick = WM_RBUTTONDBLCLK
   ercRButtonDown = WM_RBUTTONDOWN
   ercRBUttonUp = WM_RBUTTONUP
   ercSetCursor = WM_SETCURSOR
End Enum

Public Const WM_USER = &H400
Public Const EM_SETEVENTMASK = (WM_USER + 69)

Public Const WM_NOTIFY = &H4E
Public Const EN_LINK = &H70B&

'// Event Masks
Public Const ENM_NONE = &H0
Public Const ENM_CHANGE = &H1
Public Const ENM_UPDATE = &H2
Public Const ENM_SCROLL = &H4
Public Const ENM_KEYEVENTS = &H10000
Public Const ENM_MOUSEEVENTS = &H20000
Public Const ENM_REQUESTRESIZE = &H40000
Public Const ENM_SELCHANGE = &H80000
Public Const ENM_DROPFILES = &H100000
Public Const ENM_PROTECTED = &H200000
Public Const ENM_CORRECTTEXT = &H400000               ' /* PenWin specific */
Public Const ENM_SCROLLEVENTS = &H8
Public Const ENM_DRAGDROPDONE = &H10

Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const EM_SETTEXTMODE = (WM_USER + 89)

Public Const EM_AUTOURLDETECT = (WM_USER + 91)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub AttachMessages()
Dim dwMask As Long
    AttachMessage Me, hwnd, WM_NOTIFY
    '// we need to detect the link over messages
    '// by setting enm_link, however, this then
    '// cancels any other messages (such as the
    '// change event, so we need to specify
    '// these too.
    ' Key And Mouse Events
    dwMask = ENM_KEYEVENTS Or ENM_MOUSEEVENTS
    ' Selection change
    dwMask = dwMask Or ENM_SELCHANGE
    ' Update
    dwMask = dwMask Or ENM_DROPFILES
    ' Scrolling
    dwMask = dwMask Or ENM_SCROLL
    ' Update:
    dwMask = dwMask Or ENM_UPDATE
    ' Change:
    dwMask = dwMask Or ENM_CHANGE
    dwMask = dwMask Or ENM_LINK
    SendMessageLong rtfText.hwnd, EM_SETEVENTMASK, 0, dwMask
End Sub
Private Sub DetachMessages()
    DetachMessage Me, hwnd, WM_NOTIFY
End Sub
Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tNMH As NMHDR_RICHEDIT
Dim tEN As ENLINK
   Select Case iMsg

   Case WM_NOTIFY
      CopyMemory tNMH, ByVal lParam, Len(tNMH)
      If (tNMH.hwndFrom = rtfText.hwnd) Then
         Select Case tNMH.code
         Case EN_LINK
            CopyMemory tEN, ByVal lParam, Len(tEN)
            LinkOver tEN.msg, tEN.chrg.cpMin, tEN.chrg.cpMax - tEN.chrg.cpMin
         End Select
      End If
   End Select
End Function
Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer.EMsgResponse)
    '// this sub has to exist whether you like it or not
End Property
Private Property Get ISubclass_MsgResponse() As SSubTimer.EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Property
'///////////////////////////////////////////////////////
'// URL detection
Public Property Let AutoURLDetect(ByVal bState As Boolean)
    m_bAutoURLDetect = bState
    SendMessageLong rtfText.hwnd, EM_AUTOURLDETECT, Abs(bState), 0
End Property
Public Property Get AutoURLDetect() As Boolean
   AutoURLDetect = m_bAutoURLDetect
End Property

'// occurs when the mouse is moved over a link, or it is clicked
Public Sub LinkOver(ByVal iType As ERECLinkEventTypeCOnstants, ByVal lStart As Long, ByVal lLength As Long)
    Dim strText As String
    strText = Mid$(rtfText.Text, lStart + 1, lLength + 1)
    If (iType = ercLButtonUp) Then
        If ShellExecute(hwnd, vbNullString, strText, vbNullString, vbNullString, vbNormalFocus) = 2 Then
            MsgBox "Link Failed", vbExclamation
        End If
    Else
        'lblStatus = "LinkOver: " & strText
    End If
End Sub


