VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form frmHostGame 
   Caption         =   "XvT Hosting Window"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   Icon            =   "frmHostGame.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   3960
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   800
      ScreenWidth     =   1280
      ScreenHeightDT  =   800
      ScreenWidthDT   =   1280
      FormHeightDT    =   6255
      FormWidthDT     =   9660
      FormScaleHeightDT=   5745
      FormScaleWidthDT=   9540
   End
   Begin SocketWrenchCtrl.Socket Socket1 
      Index           =   0
      Left            =   6120
      Top             =   360
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
   Begin SocketWrenchCtrl.Socket Socket2 
      Index           =   0
      Left            =   6120
      Top             =   840
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
   Begin VB.TextBox txtpwd 
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
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   300
      PasswordChar    =   "*"
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CheckBox cboPwd 
      BackColor       =   &H00000000&
      Caption         =   "Use A password?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Timer keepalive 
      Interval        =   3000
      Left            =   5640
      Top             =   2160
   End
   Begin VB.CommandButton btnBOP 
      Caption         =   "Launch Balance of Power"
      Height          =   615
      Left            =   6960
      TabIndex        =   13
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton btnBoot 
      Caption         =   "Remove Player"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00151515&
      Caption         =   "Game to Play"
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
      Height          =   1215
      Left            =   6960
      TabIndex        =   8
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton optBop 
         BackColor       =   &H00151515&
         Caption         =   "Balance Of Power"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF33&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optXWA 
         BackColor       =   &H00151515&
         Caption         =   "X-Wing Alliance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF33&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optXVT 
         BackColor       =   &H00151515&
         Caption         =   "X-Wing Vs Tie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF33&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton btnBack 
      Caption         =   "Back"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6960
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5640
      Top             =   2640
   End
   Begin VB.CommandButton btnLeave 
      Caption         =   "Leave"
      Height          =   615
      Left            =   6960
      TabIndex        =   6
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton txtLaunchXWA 
      Caption         =   "Launch XWA"
      Height          =   615
      Left            =   6960
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton txtLaunchXVT 
      Caption         =   "Launch XVT"
      Height          =   615
      Left            =   6960
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
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
      Height          =   2940
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5895
      Left            =   -120
      TabIndex        =   14
      Top             =   0
      Width           =   9975
      ExtentX         =   17595
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
Attribute VB_Name = "frmHostGame"
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
Private lstArraySocket(100) As Integer
Private lstArrayLat(100) As String

Private lstArrayCount As Integer

Private selectedSocket As Integer

Private lastsocket As Integer

Private numsockets As Integer

Private backFlag As Boolean

Private Filter() As String






Private Sub btnBOP_Click()

    On Error GoTo Failed
    
    Dim path As String
    Dim i As Integer
    
        
    backFlag = False
    
    path = Chr$(34) + frmClient.BOPGamePath + Chr$(34) + " ishost /" + Chr$(34) + "a=" + frmClient.exposedIP + Chr$(34) + " /" + Chr$(34) + "n=host" + Chr$(34) + " /" + Chr$(34) + "skipintro" + Chr$(34)
    
        
    Shell path
    
        
    For i = 0 To lastsocket
                       
        If Socket2(i).Connected Then
            Socket2(i).SendLen = Len("#@+%BOP" + Chr$(3))
            Socket2(i).SendData = "#@+%BOP" + Chr$(3)
        End If
    Next i
    
    GoSleepEX 1
     
    frmClient.Socket1.SendLen = Len("^^^^" + "(P)" + frmClient.txtuid.Text + Chr$(3))
    frmClient.Socket1.SendData = "^^^^" + "(P)" + frmClient.txtuid.Text + Chr$(3)
    
    txtLaunchXVT.Enabled = False
    btnBOP.Enabled = False
    txtLaunchXWA.Enabled = False
    btnLeave.Enabled = False
    'btnLeave.Enabled = True
    btnBack.Enabled = True
    btnSend.Enabled = False
    frmClient.playing = True
     Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnBOP_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next


End Sub


Private Sub cboPwd_Click()
    On Error GoTo Failed
    If cboPwd.Value = 1 Then txtpwd.Visible = True Else txtpwd.Visible = False
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "cboPwd_click - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo Failed
        If backFlag = False Then Cancel = 1
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "form_QueryUnload - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next


End Sub

Private Sub keepalive_Timer()
    On Error GoTo Failed
    frmClient.Socket1.SendLen = Len("KALV" + frmClient.gameName + Chr$(3))
    frmClient.Socket1.SendData = "KALV" + frmClient.gameName + Chr$(3)
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "keepalive_timer - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next


End Sub

Private Sub Timer1_Timer()
    On Error GoTo Failed

    Dim cnt As Integer
    Dim lstarraytemp() As String
    Dim lstarraysockettemp() As Integer
    Dim j As Integer
    Dim parsed As String
    Dim i As Integer
    ' loop through the list of players in room
    For j = 1 To lstArrayCount - 1
    'check to see that the player is still connected
        If Socket2(lstArraySocket(j)).Connected = False Then
        
                'if not connected then get the name of player
                parsed = lstArray(j)
                ' clear the list of players
                lstPlayers.Clear
                ' set the count to 0
                cnt = 0
                ' redimention the temp arrays
                ReDim lstarraytemp(lstArrayCount)
                ReDim lstarraysockettemp(lstArrayCount)
                
                ' loop through the list and re add everyone except the disconnected player
                For i = 0 To lstArrayCount - 1
                
                    'if the disconnected player name does not equal the current name
                    If parsed <> lstArray(i) Then
                        'add the name back into the list
                        lstPlayers.AddItem lstArray(i)
                        'add the player info back to the array
                        lstarraytemp(cnt) = lstArray(i)
                        lstarraysockettemp(cnt) = lstArraySocket(i)
                        'increment the counter
                        cnt = cnt + 1
                    
                    End If
                Next i
                
                ' set the array to equal the count (removed players not connected)
                lstArrayCount = cnt
                
                 ' decrement the number of sockets checked
                lastsocket = lastsocket - 1
                ' loop through the list of elements
                For i = 0 To lstArrayCount - 1
                    'copy from the temp array back into the arrays
                    lstArray(i) = lstarraytemp(i)
                    lstArraySocket(i) = lstarraysockettemp(i)
                
                Next i
                ' update the game stats
                frmClient.Socket1.SendLen = Len("^++^" + Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName + Chr$(3))
                frmClient.Socket1.SendData = "^++^" + Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName + Chr$(3)
        
            
        
        
        End If
    
    Next j
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "timer1_timer - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next

    

End Sub



Private Sub btnBoot_Click()
    On Error GoTo Failed
    Dim ok2boot As Integer
    
    ok2boot = MsgBox("Are you sure you want to Boot player " + lstPlayers.Text + "?", vbYesNo + vbQuestion, "Really boot Player?")

    If ok2boot = 6 Then
    
        Socket2(selectedSocket).SendLen = Len("$&&$" + Chr$(3))
        Socket2(selectedSocket).SendData = "$&&$" + Chr$(3)
        
        btnBoot.Enabled = True
        
    End If
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnboot_click - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next



End Sub

Private Sub lstPlayers_Click()
    
    On Error GoTo Failed
    If lstPlayers.Text <> "" And lstPlayers.Text <> frmClient.txtuid.Text Then
        selectedSocket = lstArraySocket(lstPlayers.ListIndex)
        btnBoot.Enabled = True
    End If
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "lstPlayers_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next


End Sub


Private Sub btnBack_Click()
    On Error GoTo Failed
    backFlag = True
    
    frmClient.playing = False
    
    
    txtLaunchXVT.Enabled = True
    btnBOP.Enabled = True
    txtLaunchXWA.Enabled = True
    
    btnLeave.Enabled = True
    btnBack.Enabled = False
    
    frmClient.Socket1.SendLen = Len("^^^^" + "(H)" + frmClient.txtuid.Text + Chr$(3))
    frmClient.Socket1.SendData = "^^^^" + "(H)" + frmClient.txtuid.Text + Chr$(3)
    btnSend.Enabled = True
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnBack_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next

   
End Sub

Private Sub btnLeave_Click()
    On Error GoTo Failed
    Unload Me
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnLeave_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next


End Sub

Private Sub btnSend_Click()
    On Error GoTo Failed

            Dim i As Integer
            Dim w As Integer
            Dim parsed As String
            Dim wrapLines(10) As String
            
                        
            If txtSend.Text <> "" Then
            
                For i = 0 To lastsocket
                       
                    If Socket2(i).Connected Then
                        Socket2(i).SendLen = Len("MMMM" + "- " + frmClient.txtuid.Text + " - " + txtSend.Text + Chr$(3))
                        Socket2(i).SendData = "MMMM" + "- " + frmClient.txtuid.Text + " - " + txtSend.Text + Chr$(3)
                        
                        
                    End If
                Next i
                
                parsed = "- " + frmClient.txtuid.Text + " - " + txtSend.Text
                
                For w = 0 To Int(Len(parsed) / 60)
                
                    wrapLines(w) = Mid(parsed, (w * 60) + 1, 60)
                
                Next w
                            
                For w = Int(Len(parsed) / 60) To 0 Step -1
                    newMessage wrapLines(w)
                Next w
    
                txtSend.Text = ""
            End If
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnSend - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next


End Sub

Private Sub Form_Load()
            
    On Error GoTo Failed
    Filter = frmClient.getFilter()
        
    WebBrowser1.Navigate2 CurDir + "\gameroom.htm"
    
    Dim i As Integer
     backFlag = True
    For i = 0 To 100
        lstReplyarray(i) = ""
        lstArray(i) = ""
    Next i
    lstArrayCount = 0
    lastsocket = 0
    On Error GoTo errorhandler
    Me.Caption = Trim(Str(lstArrayCount + 1)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName
    numsockets = 1
    For i = 1 To numsockets
        Load Socket1(i)
        Socket1(i).AddressFamily = AF_INET
        Socket1(i).Protocol = IPPROTO_IP
        Socket1(i).SocketType = SOCK_STREAM
        Socket1(i).Blocking = False
        Socket1(i).LocalPort = 2303 + i
        Socket1(i).Action = SOCKET_LISTEN
        lastsocket = 0

    Next i
    Load Socket1(i)
    Socket1(i).AddressFamily = AF_INET
    Socket1(i).Protocol = IPPROTO_IP
    Socket1(i).SocketType = SOCK_STREAM
    Socket1(i).Blocking = False
    Socket1(i).LocalPort = 1001
    Socket1(i).Action = SOCKET_LISTEN

    frmClient.Socket1.SendLen = Len("^^^^" + "(H)" + frmClient.txtuid.Text + Chr$(3))
    frmClient.Socket1.SendData = "^^^^" + "(H)" + frmClient.txtuid.Text + Chr$(3)

    lstArrayCount = 1
    lstArray(0) = frmClient.txtuid.Text
    lstPlayers.AddItem frmClient.txtuid.Text
    If frmClient.XVTGamePath = "" Then
        optXVT.Visible = False
    End If
    
    If frmClient.BOPGamePath = "" Then
        optBop.Visible = False
    End If
    
    If frmClient.XWAGamePath = "" Then
        optXWA.Visible = False
    End If
        
    
    
    If frmClient.game = "XVT" Then
        optXVT.Value = True
    ElseIf frmClient.game = "BOP" Then
        optBop.Value = True
    Else
        optXWA.Value = True
    End If
    
    'does this in the gameroom configuration window
    'frmClient.Socket1.SendLen = Len("GRPL" + frmClient.gameName + chr$(30) + "Host: " + frmClient.txtuid.Text + chr$(3))
    'frmClient.Socket1.SendData = "GRPL" + frmClient.gameName + chr$(30) + "Host: " + frmClient.txtuid.Text + chr$(3)
    Exit Sub
errorhandler:
    If Str(Err.Number) = 24048 Then
        MsgBox "Port " + Str(2303 + i) + " is already being used... Ports 2303 and 2304 are the two ports used for listening, at least one needs to be open", vbExclamation + vbOKOnly, "Warning"
        Resume Next
    End If
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "Form_Load - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next

    
End Sub

Private Sub optXVT_Click()
        
    On Error GoTo Failed
        txtLaunchXWA.Visible = False
        txtLaunchXVT.Visible = True
        btnBOP.Visible = False

        frmClient.game = "XVT"
        
        frmClient.Socket1.SendLen = Len("^++^" + Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName + Chr$(3))
        frmClient.Socket1.SendData = "^++^" + Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName + Chr$(3)
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "optXVT_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next

        
End Sub

Private Sub optXWA_Click()
    On Error GoTo Failed
        txtLaunchXWA.Visible = True
        txtLaunchXVT.Visible = False
        btnBOP.Visible = False

        frmClient.game = "XWA"
        
        frmClient.Socket1.SendLen = Len("^++^" + Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName + Chr$(3))
        frmClient.Socket1.SendData = "^++^" + Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName + Chr$(3)
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "optXWA_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next


End Sub
Private Sub optBOP_Click()
    On Error GoTo Failed
        txtLaunchXWA.Visible = False
        txtLaunchXVT.Visible = False
        btnBOP.Visible = True
        
        frmClient.game = "BOP"
        
        frmClient.Socket1.SendLen = Len("^++^" + Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName + Chr$(3))
        frmClient.Socket1.SendData = "^++^" + Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName + Chr$(3)
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "optBop_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next


End Sub



Private Sub txtLaunchXWA_Click()

    On Error GoTo Failed
    Dim path As String
    Dim i As Integer
    
    backFlag = False
    
    
    path = Chr$(34) + frmClient.XWAGamePath + Chr$(34) + " ishost /" + Chr$(34) + "a=" + frmClient.exposedIP + Chr$(34) + " /" + Chr$(34) + "n=host" + Chr$(34) + " /" + Chr$(34) + "skipintro" + Chr$(34)
        
    Shell path
    
        
    For i = 0 To lastsocket
                       
        If Socket2(i).Connected Then
            Socket2(i).SendLen = Len("#@+%XWA" + Chr$(3))
            Socket2(i).SendData = "#@+%XWA" + Chr$(3)
        End If
    Next i
    
    GoSleepEX 1
     
    frmClient.Socket1.SendLen = Len("^^^^" + "(P)" + frmClient.txtuid.Text + Chr$(3))
    frmClient.Socket1.SendData = "^^^^" + "(P)" + frmClient.txtuid.Text + Chr$(3)
    
    txtLaunchXVT.Enabled = False
    txtLaunchXWA.Enabled = False
    btnLeave.Enabled = False
    'btnLeave.Enabled = True
    btnBack.Enabled = True
    btnSend.Enabled = False
    frmClient.playing = True
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "txtLaunchXWA - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next


End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    On Error GoTo Failed
  If KeyAscii = Asc(Chr$(3)) Then KeyAscii = 0
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "txtSend_Keypress - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next

    
End Sub
Private Sub Socket1_Accept(index As Integer, SocketId As Integer)

    On Error GoTo Failed

        Dim i As Integer
        For i = 0 To lastsocket
            If Not Socket2(i).Connected Then Exit For
        Next i
        If i > lastsocket Then
            lastsocket = lastsocket + 1: i = lastsocket
            Load Socket2(i)
        End If
        
        Socket2(i).AddressFamily = AF_INET
        Socket2(i).Protocol = IPPROTO_IP
        Socket2(i).SocketType = SOCK_STREAM
        Socket2(i).Binary = True
        Socket2(i).BufferSize = 1024
        Socket2(i).Blocking = False
        Socket2(i).Accept = SocketId
            
        If cboPwd.Value = 1 Then
            Socket2(i).SendLen = Len("GRPW" + txtpwd.Text + Chr$(3))
            Socket2(i).SendData = "GRPW" + txtpwd.Text + Chr$(3)
        Else
            Socket2(i).SendLen = Len("GRPW" + Chr$(3))
            Socket2(i).SendData = "GRPW" + Chr$(3)
        End If
        
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "socket1.accept - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next


    
End Sub
Private Sub Socket2_Read(index As Integer, DataLength As Integer, IsUrgent As Integer)
    
    On Error GoTo Failed
    
    Dim strData As String
    Dim first4 As String
    Dim parsed As String
    Dim cnt As Integer
    Dim wrapLines(10) As String
    
    Dim i As Integer
    Dim w As Integer
    Dim B As Integer
    Dim numMsgs As Integer
    Dim msgs(100) As String
    Dim nextMsgStart As Integer
    
    Dim lstarraytemp() As String
    Dim lstarraysockettemp() As Integer
    
    Dim GRPLString As String
    
    
    
    nextMsgStart = 1
    numMsgs = 0
    
    Socket2(index).Read strData, DataLength
    
    
    For B = 1 To Len(strData)
    
        If Mid(strData, B, 1) = Chr$(3) Then
            msgs(numMsgs) = Mid(strData, nextMsgStart, B - nextMsgStart)
            numMsgs = numMsgs + 1
            nextMsgStart = B + 1
        End If
    Next B
    
    For B = 0 To numMsgs - 1
    
        strData = msgs(B)
        
        first4 = Mid(strData, 1, 4)
        
        parsed = Mid(strData, 5)
        
        Select Case first4
        
        Case "$$$$"
                If frmClient.cboMute.Value <> 1 And btnBack.Enabled = False Then frmClient.PlayWave App.path + "/joingameroom.wav"

                parsed = Mid(strData, 5)
                
                lstPlayers.AddItem parsed
                lstArray(lstArrayCount) = parsed
                lstArraySocket(lstArrayCount) = index
                lstArrayCount = lstArrayCount + 1
                
                For i = 0 To lstArrayCount - 1
                
                    Socket2(index).SendLen = Len("$$$$" + lstArray(i) + Chr$(3))
                    Socket2(index).SendData = "$$$$" + lstArray(i) + Chr$(3)
                    
                    
                    'GoSleepEx 0.1
                Next i
                
                For i = 0 To lastsocket
                       
                    If Socket2(i).Connected And i <> index Then
                        Socket2(i).SendLen = Len(strData + Chr$(3))
                        Socket2(i).SendData = strData + Chr$(3)
                        
                    End If
                Next i
        
                frmClient.Socket1.SendLen = Len("^++^" + Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName + Chr$(3))
                frmClient.Socket1.SendData = "^++^" + Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName + Chr$(3)
                
                GRPLString = "Host:" + " " + frmClient.txtuid.Text + ", "
                
                For i = 1 To lstArrayCount - 1
                    GRPLString = GRPLString + ", " + lstArray(i)
                Next i
                
                
                frmClient.Socket1.SendLen = Len("GRPL" + frmClient.gameName + Chr$(30) + GRPLString + Chr$(3))
                frmClient.Socket1.SendData = "GRPL" + frmClient.gameName + Chr$(30) + GRPLString + Chr$(3)
                
                
                
        Case "&&&&"
                If frmClient.cboMute.Value <> 1 And btnBack.Enabled = False Then frmClient.PlayWave App.path + "/leavegameroom.wav"
    
                For i = 0 To lastsocket
                                   
                    If Socket2(i).Connected Then
                        Socket2(i).SendLen = Len("&&&&" + parsed + Chr$(3))
                        Socket2(i).SendData = "&&&&" + parsed + Chr$(3)
                    End If
                Next i
                
                lstPlayers.Clear
                cnt = 0
                ReDim lstarraytemp(lstArrayCount)
                ReDim lstarraysockettemp(lstArrayCount)
                
    
                For i = 0 To lstArrayCount - 1
                
                    
                    If parsed <> lstArray(i) Then
                    
                        lstPlayers.AddItem lstArray(i)
                        lstarraytemp(cnt) = lstArray(i)
                        lstarraysockettemp(cnt) = lstArraySocket(i)
                        cnt = cnt + 1
                    
                    End If
                
                
                Next i
                
                lstArrayCount = cnt
                For i = 0 To lstArrayCount - 1
                
                    lstArray(i) = lstarraytemp(i)
                    lstArraySocket(i) = lstarraysockettemp(i)
                
                Next i
        
        
                'lstArrayCount = lstArrayCount - 1
                                
                frmClient.Socket1.SendLen = Len("^++^" + Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName + Chr$(3))
                frmClient.Socket1.SendData = "^++^" + Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName + Chr$(3)

                GRPLString = "Host:" + " " + frmClient.txtuid.Text + ", "
                
                For i = 1 To lstArrayCount - 1
                    GRPLString = GRPLString + ", " + lstArray(i)
                Next i
                
                
                frmClient.Socket1.SendLen = Len("GRPL" + frmClient.gameName + Chr$(30) + GRPLString + Chr$(3))
                frmClient.Socket1.SendData = "GRPL" + frmClient.gameName + Chr$(30) + GRPLString + Chr$(3)
        
       ' Case "#@+%":
        
            
            
            
            
            
        Case "MMMM":
        
        
        
            For i = 0 To lastsocket
                       
                If Socket2(i).Connected Then
                    Socket2(i).SendLen = Len(strData + Chr$(3))
                    Socket2(i).SendData = strData + Chr$(3)
                        
                End If
            Next i
        
        
            For w = 0 To Int(Len(parsed) / 60)
            
                wrapLines(w) = Mid(parsed, (w * 60) + 1, 60)
            
            Next w
                        
            For w = Int(Len(parsed) / 60) To 0 Step -1
                newMessage wrapLines(w)
            Next w
        
        End Select
        
    Next B
    
    Me.Caption = Trim(Str(lstArrayCount)) + "/" + Trim(Str(frmClient.numberOfPlayers)) + " " + frmClient.game + " " + frmClient.gameName
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "socket2_Read - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
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

    'If Not focusFlag Then Beep
    
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "NewMessage - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next



End Sub
Private Sub Socket2_Disconnect(index As Integer)
    On Error GoTo Failed
    Socket2(index).Action = SOCKET_CLOSE
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "Socket2_Disconnect - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next


End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Failed
    Dim i As Integer
    Dim temp As Integer
    
    
        If btnBack.Enabled Then Call btnBack_Click
        
        frmClient.Socket1.SendLen = Len("%%%%" + frmClient.gameName + Chr$(3))
        frmClient.Socket1.SendData = "%%%%" + frmClient.gameName + Chr$(3)
     
        'GoSleepEx 0.3
        
        For i = 0 To lastsocket
                       
            If Socket2(i).Connected Then
                Socket2(i).SendLen = Len("%&&%" + Chr$(3))
                Socket2(i).SendData = "%&&%" + Chr$(3)
                        
            End If
        Next i
        
        GoSleep 0.3
        
        For i = 0 To numsockets + 1
            If Socket1(0).Listening Then Socket1(0).Action = SOCKET_CLOSE
        Next i
        
        For i = 0 To lastsocket
            If Socket2(i).Connected Then Socket2(i).Action = SOCKET_CLOSE
        Next i
        
        
    
    
    
        
        frmClient.btnHost.Enabled = True
        frmClient.btnAway.Enabled = True
        frmClient.lstGameRooms.Enabled = True
        
        'GoSleepEx 0.3
        
        frmClient.Socket1.SendLen = Len("^^^^" + "(L)" + frmClient.txtuid.Text + Chr$(3))
        frmClient.Socket1.SendData = "^^^^" + "(L)" + frmClient.txtuid.Text + Chr$(3)
    
    
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "Foram_UnLoad - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next

End Sub

Private Sub txtLaunchXVT_Click()
    
    On Error GoTo Failed
    Dim i As Integer
    
    backFlag = False
    
    
    
    Open CurDir + "\hostgamepath.bat" For Output As #1
        Print #1, Chr$(34) + frmClient.XVTGamePath + Chr$(34) + " ishost /" + Chr$(34) + "a=" + frmClient.exposedIP + Chr$(34) + " /" + Chr$(34) + "n=host" + Chr$(34) + Chr(13) + "exit"
    Close #1
    
    Shell CurDir + "\hostGamePath.bat"
        
    For i = 0 To lastsocket
                       
        If Socket2(i).Connected Then
            Socket2(i).SendLen = Len("#@+%XVT" + Chr$(3))
            Socket2(i).SendData = "#@+%XVT" + Chr$(3)
        End If
    Next i
 
    frmClient.Socket1.SendLen = Len("^^^^" + "(P)" + frmClient.txtuid.Text + Chr$(3))
    frmClient.Socket1.SendData = "^^^^" + "(P)" + frmClient.txtuid.Text + Chr$(3)
    
    txtLaunchXVT.Enabled = False
    txtLaunchXWA.Enabled = False
    btnLeave.Enabled = False
    btnBack.Enabled = True
    btnSend.Enabled = False
    frmClient.playing = True
    
    
        
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "txtLaunchXvT_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
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
    Print #4, "shell - " + Str(Err.Number) + " - " + Err.Description + " - frmHostGame"
    Close #4
    Resume Next

End Function

