VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form frmJoinGame 
   Caption         =   "XvT Join Game Window"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   Icon            =   "frmJoinGame.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   6120
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   800
      ScreenWidth     =   1280
      ScreenHeightDT  =   800
      ScreenWidthDT   =   1280
      FormHeightDT    =   6240
      FormWidthDT     =   9675
      FormScaleHeightDT=   5730
      FormScaleWidthDT=   9555
   End
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6480
      Top             =   1200
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.CommandButton btnIPIN 
      Caption         =   "IP In"
      Height          =   615
      Left            =   7560
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton btnBack 
      Caption         =   "Back"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7560
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton btnLeave 
      Caption         =   "Leave"
      Height          =   615
      Left            =   7560
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.Timer timerconnection 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3480
      Top             =   2520
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00151515&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      MaxLength       =   300
      TabIndex        =   2
      Top             =   3720
      Width           =   8175
   End
   Begin VB.ListBox lstReply 
      BackColor       =   &H00151515&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1500
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   9375
   End
   Begin VB.ListBox lstPlayers 
      BackColor       =   &H00151515&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3420
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5895
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10215
      ExtentX         =   18018
      ExtentY         =   10398
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmJoinGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private lstReplyarray(100) As String
Private lstArray(100) As String
Private lstArrayCount As Integer
Private remPort As Long
Private rport(3) As String
Private launchedGame As String

Private killed As Boolean

Private backFlag As Boolean

Private Filter() As String



Private Sub btnBack_Click()
    
    On Error GoTo Failed
    
    btnLeave.Enabled = True
    btnBack.Enabled = False
    'btnIPIN.Enabled = False
    btnSend.Enabled = True
    backFlag = True
    
    frmClient.playing = False
    
    frmClient.Socket1.SendLen = Len("^^^^" + "(G)" + frmClient.txtuid.Text + Chr$(3))
    frmClient.Socket1.SendData = "^^^^" + "(G)" + frmClient.txtuid.Text + Chr$(3)
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnBack_click - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next

End Sub

Private Sub btnIPIN_Click()
    
    On Error GoTo Failed
    
    Dim path As String
    backFlag = False
    btnBack.Enabled = True
    frmClient.playing = True



             If launchedGame = "XWA" Then
                
                    
                    path = Chr$(34) + frmClient.XWAGamePath + Chr$(34) + " isclient /" + Chr$(34) + "a=" + Socket1.HostAddress + Chr$(34) + " /" + Chr$(34) + "n=client" + Chr$(34) + " /" + Chr$(34) + "skipintro" + Chr$(34)

                    
                    Shell (path)
                
                ElseIf launchedGame = "BOP" Then
                
                    
                    Open CurDir + "\joingamepath.bat" For Output As #1
                    
                        Print #1, Chr$(34) + frmClient.BOPGamePath + Chr$(34) + " isclient /" + Chr$(34) + "a=" + Socket1.HostAddress + Chr$(34) + " /" + Chr$(34) + "n=client" + Chr$(34)
                        
                    Close #1
        
                    Shell CurDir + "\joinGamePath.bat"
                                
                Else
                    
                    Open CurDir + "\joingamepath.bat" For Output As #1
                    
                        Print #1, Chr$(34) + frmClient.XVTGamePath + Chr$(34) + " isclient /" + Chr$(34) + "a=" + Socket1.HostAddress + Chr$(34) + " /" + Chr$(34) + "n=client" + Chr$(34)
                        
                    Close #1
        
                    Shell CurDir + "\joinGamePath.bat"
            
                End If
        Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnIPIN_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next

End Sub

Private Sub btnLeave_Click()
    On Error GoTo Failed
    Unload Me
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnLeave_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next
    
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    On Error GoTo Failed
    Dim strBuffer As String
    Dim first4 As String
    Dim parsed As String
    Dim cnt As Integer
    Dim lstarraytemp() As String
    
    Dim wrapLines(10) As String
    
        
    Dim B As Integer
    Dim numMsgs As Integer
    Dim msgs(100) As String
    Dim nextMsgStart As Integer
    Dim w As Integer
    
    Dim path As String
    
    Dim i As Integer
    
    Dim tempPassword As String
    tempPassword = " "
    
    nextMsgStart = 1
    numMsgs = 0
    
    Socket1.Read strBuffer, DataLength
    
    For B = 1 To Len(strBuffer)
    
        If Mid(strBuffer, B, 1) = Chr$(3) Then
            msgs(numMsgs) = Mid(strBuffer, nextMsgStart, B - nextMsgStart)
            numMsgs = numMsgs + 1
            nextMsgStart = B + 1
        End If
    Next B
    
    For B = 0 To numMsgs - 1
    
        strBuffer = msgs(B)
    
        first4 = Mid(strBuffer, 1, 4)
    
        parsed = Mid(strBuffer, 5)
    
    
        Select Case first4
         
        Case "%&&%":
            
            killed = True
            Unload Me
            
         
        Case "$$$$":
            
            
            
            If frmClient.cboMute.Value <> 1 And btnBack.Enabled = False Then frmClient.PlayWave App.path + "/joingameroom.wav"
        
        
            lstPlayers.AddItem parsed
            lstArray(lstArrayCount) = parsed
            lstArrayCount = lstArrayCount + 1
        
            Me.Caption = Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName
            
        
        
        Case "&&&&"
                If frmClient.cboMute.Value <> 1 And btnBack.Enabled = False Then frmClient.PlayWave App.path + "/leavegameroom.wav"
    
                lstPlayers.Clear
                cnt = 0
                ReDim lstarraytemp(lstArrayCount)
                
    
                For i = 0 To lstArrayCount - 1
                
                    
                    If parsed <> lstArray(i) Then
                    
                        lstPlayers.AddItem lstArray(i)
                        lstarraytemp(cnt) = lstArray(i)
                        cnt = cnt + 1
                    
                    End If
                
                
                Next i
                
                lstArrayCount = cnt
                For i = 0 To lstArrayCount - 1
                
                    lstArray(i) = lstarraytemp(i)
                
                Next i
    
        Case "$&&$":
    
                backFlag = True
                
                Unload Me
                
                MsgBox "The Host has removed you from the Gameroom", vbOKOnly, "Bye"
    
    
                
        Case "#@+%":
                
                PlayASound (App.path + "\" + "launch.wav")
                frmClient.playing = True
                
                launchedGame = Mid(parsed, 1, 3)
                
                If Mid(parsed, 1, 3) = "XWA" Then
                
                    
                    path = Chr$(34) + frmClient.XWAGamePath + Chr$(34) + " isclient /" + Chr$(34) + "a=" + Socket1.HostAddress + Chr$(34) + " /" + Chr$(34) + "n=client" + Chr$(34) + " /" + Chr$(34) + "skipintro" + Chr$(34)

                    
                    Shell path
                
                ElseIf Mid(parsed, 1, 3) = "BOP" Then
                
                    
                    Open CurDir + "\joingamepath.bat" For Output As #1
                    
                        Print #1, Chr$(34) + frmClient.BOPGamePath + Chr$(34) + " isclient /" + Chr$(34) + "a=" + Socket1.HostAddress + Chr$(34) + " /" + Chr$(34) + "n=client" + Chr$(34) + Chr(13) + "exit"
                        
                    Close #1
        
                    Shell CurDir + "\joinGamePath.bat"
                                
                Else
                    
                    Open CurDir + "\joingamepath.bat" For Output As #1
                    
                        Print #1, Chr$(34) + frmClient.XVTGamePath + Chr$(34) + " isclient /" + Chr$(34) + "a=" + Socket1.HostAddress + Chr$(34) + " /" + Chr$(34) + "n=client" + Chr$(34) + Chr$(13) + "exit"
                        
                    Close #1
        
                    Shell CurDir + "\joinGamePath.bat"
            
                End If
        
                    'btnIPIN.Enabled = True
                    
                    backFlag = False
                    
                    
                    frmClient.Socket1.SendLen = Len("^^^^" + "(P)" + frmClient.txtuid.Text + Chr$(3))
                    frmClient.Socket1.SendData = "^^^^" + "(P)" + frmClient.txtuid.Text + Chr$(3)
                    
                    btnLeave.Enabled = False
                    btnBack.Enabled = True
                    btnSend.Enabled = False
        
        Case "MMMM":
        
            
            For w = 0 To Int(Len(parsed) / 60)
            
                wrapLines(w) = Mid(parsed, (w * 60) + 1, 60)
            
            Next w
                        
            For w = Int(Len(parsed) / 60) To 0 Step -1
                newMessage wrapLines(w)
            Next w
            
        Case "GRPW":
        
            If parsed <> "" Then
            
                Do While tempPassword = " "
                    tempPassword = InputBox("The Host requires a password to join this gameroom", "Enter Password", " ")
                Loop
                
                If tempPassword <> parsed Then
                    MsgBox "Bad Password", vbOKOnly + vbCritical, "Incorrect Password"
                    Call btnLeave_Click
                    Exit Sub
                End If
            End If
            Timer1.Enabled = True
            
            
            
        End Select
    Next B
    
    Me.Caption = Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "socket1_Read - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next

End Sub

Private Sub newMessage(msg As String)
    On Error GoTo Failed
    Dim i As Integer
    
    
    For i = 0 To frmClient.FilterCount

        If frmClient.cboFilterOff.Value <> 1 Then
        
            msg = Replace(msg, Filter(i), "Smurf", , , vbTextCompare)
        
        End If

    Next i

    lstReply.Clear
    
    For i = 99 To 0 Step -1
    
       lstReplyarray(i + 1) = lstReplyarray(i)
    
    Next i
    
    lstReplyarray(0) = msg
    For i = 0 To 100
       lstReply.AddItem lstReplyarray(i)
    Next i

    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "newmessage - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next



End Sub


Private Sub timerconnection_Timer()
    On Error GoTo Failed
    If Not Socket1.Connected Then
        MsgBox "The Connection to the host has been terminated", vbExclamation + vbOKOnly, "Lost Connection"
        Unload Me
        
    End If
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "timerconnection_Timer - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next
    
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    On Error GoTo Failed
  If KeyAscii = Asc(Chr$(3)) Then KeyAscii = 0
  Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "txtSend_Keypress - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next
    
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo Failed
        If backFlag = False Then Cancel = 1
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "Form_QueryUnload - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next
End Sub
Private Sub Form_Load()
    
    On Error GoTo Failed
    Filter = frmClient.getFilter()
    
    WebBrowser1.Navigate2 CurDir + "\gameroom.htm"

    rport(0) = "1001"
    rport(1) = "2303"
    rport(2) = "2304"
    
    backFlag = True
    Dim i As Integer
    
    For i = 0 To 100
        lstReplyarray(i) = ""
        lstArray(i) = ""
    Next i
    lstArrayCount = 0
                       
    remPort = 1
    
    Socket1.AddressFamily = AF_INET
    Socket1.Protocol = IPPROTO_IP
    Socket1.SocketType = SOCK_STREAM
    Socket1.Binary = False
    Socket1.Blocking = False
    Socket1.BufferSize = 1024
    lstArrayCount = 0
    'focusFlag = True
    killed = False
    
    
    Dim strng As String
    
    lstArrayCount = 0
        
    strng = Trim(frmClient.txtuid)
    
   ' On Error GoTo Failed
    Socket1.HostName = Trim$(frmClient.GameHost)
    
    
    Socket1.RemotePort = 1001
    Socket1.Action = SOCKET_CONNECT
    
    'GoSleepEx 0.1
    
    
    frmClient.Socket1.SendLen = Len("^^^^" + "(G)" + frmClient.txtuid.Text + Chr$(3))
    frmClient.Socket1.SendData = "^^^^" + "(G)" + frmClient.txtuid.Text + Chr$(3)
    
    
        
    'GoSleepEx 2
    
    Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "form_Load - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next
    

End Sub
Private Sub btnSend_Click()

    On Error GoTo Failed
        Dim parse As String
        
        
        If txtSend.Text <> "" Then
        
            parse = "MMMM" + "- " + frmClient.txtuid.Text + " - " + txtSend.Text
            
            
            Socket1.SendLen = Len(parse + Chr$(3))
            Socket1.SendData = parse + Chr$(3)
            txtSend.Text = ""
        End If
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnSend_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next

End Sub

Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Failed
    If btnBack.Enabled = True Then Call btnBack_Click
    
    On Error GoTo failsafe
    If Socket1.Connected Then
        frmClient.Socket1.SendLen = Len("^^^^" + "(L)" + frmClient.txtuid.Text + Chr$(3))
        frmClient.Socket1.SendData = "^^^^" + "(L)" + frmClient.txtuid.Text + Chr$(3)
        
        'GoSleepEx 0.1
        
        If killed = False Then
        
            Socket1.SendLen = Len("&&&&" + frmClient.txtuid.Text + Chr$(3))
            Socket1.SendData = "&&&&" + frmClient.txtuid.Text + Chr$(3)
            Socket1.Action = SOCKET_CLOSE
            
        End If

    End If
    
    
    frmClient.btnHost.Enabled = True
    frmClient.btnAway.Enabled = True
    frmClient.lstGameRooms.Enabled = True
        
    Exit Sub
    
    
    
failsafe:
    frmClient.btnHost.Enabled = True
    frmClient.lstGameRooms.Enabled = True

    MsgBox "Disconnected", vbOKOnly, "Game Room lost connection to Host or never got connected"

    Unload Me
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "Form_UnLoad - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next

End Sub

Private Sub Timer1_Timer()

    On Error GoTo Failed
    On Error GoTo cannotconnect
    
    If Socket1.Connected Then
    
    
        Socket1.SendLen = Len("$$$$" + frmClient.txtuid.Text + Chr$(3))
        Socket1.SendData = "$$$$" + frmClient.txtuid.Text + Chr$(3)
        Timer1.Enabled = False
        timerconnection.Enabled = True
        
    Else
        
        Socket1.RemotePort = rport(remPort)
        Socket1.Action = SOCKET_CONNECT
        remPort = remPort + 1
        If remPort = 3 Then remPort = 0
    End If

    Exit Sub
    
cannotconnect:
    MsgBox "The Host cannot be connected to, they may be behind a firewall, or router that is blocking access", vbOKOnly + vbCritical, "UN Reachable host"
    
    frmClient.Socket1.SendLen = Len("^^^^" + "(L)" + frmClient.txtuid.Text + Chr$(3))
    frmClient.Socket1.SendData = "^^^^" + "(L)" + frmClient.txtuid.Text + Chr$(3)
     
    Unload Me
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "timer1 - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next

End Sub

Function Shell(Program As String, Optional ShowCmd As Long = _
vbNormalNoFocus, Optional ByVal WorkDir As Variant) As Long

    On Error GoTo Failed
    Dim FirstSpace As Integer, Slash As Integer

    If Left(Program, 1) = """" Then
        FirstSpace = InStr(2, Program, """")


        If FirstSpace <> 0 Then
            Program = Mid(Program, 2, FirstSpace - 2) & _
              Mid(Program, FirstSpace + 1)
            FirstSpace = FirstSpace - 1
        End If

    Else
        FirstSpace = InStr(Program, " ")
    End If

    If FirstSpace = 0 Then FirstSpace = Len(Program) + 1

    If IsMissing(WorkDir) Then

        For Slash = FirstSpace - 1 To 1 Step -1
            If Mid(Program, Slash, 1) = "\" Then Exit For
        Next

        If Slash = 0 Then
            WorkDir = CurDir
        ElseIf Slash = 1 Or Mid(Program, Slash - 1, 1) = ":" Then
            WorkDir = Left(Program, Slash)
        Else
            WorkDir = Left(Program, Slash - 1)
        End If

    End If

    Shell = ShellExecute(0, vbNullString, _
    Left(Program, FirstSpace - 1), LTrim(Mid(Program, _
    FirstSpace)), WorkDir, ShowCmd)
    If Shell < 32 Then VBA.Shell Program, ShowCmd 'To raise Error
    
    Exit Function
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, " - " + Str(Err.Number) + " - " + Err.Description + " - frmJoinGame"
    Close #4
    Resume Next


End Function


