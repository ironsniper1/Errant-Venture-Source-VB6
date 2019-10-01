VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form frmClient 
   BackColor       =   &H00000000&
   Caption         =   "XvT Client"
   ClientHeight    =   8145
   ClientLeft      =   1860
   ClientTop       =   2445
   ClientWidth     =   11970
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   543
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   7440
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
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   7560
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   800
      ScreenWidth     =   1280
      ScreenHeightDT  =   800
      ScreenWidthDT   =   1280
      ApplicationName =   "Active Resize Control"
      FormHeightDT    =   8655
      FormWidthDT     =   12090
      FormScaleHeightDT=   543
      FormScaleWidthDT=   798
   End
   Begin VB.ComboBox txtHost 
      BackColor       =   &H00000000&
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
      Height          =   360
      Left            =   720
      TabIndex        =   33
      Top             =   120
      Width           =   2175
   End
   Begin VB.CheckBox cboFilterOff 
      BackColor       =   &H00000000&
      Caption         =   "Disable Language Filter?"
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
      Height          =   255
      Left            =   4560
      TabIndex        =   32
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton cmdunBan 
      Caption         =   "UnBan Player"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   31
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   30
      Text            =   "Players"
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   8160
      TabIndex        =   29
      Text            =   "Game Rooms"
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Text            =   "Chat"
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   27
      Text            =   "Host"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   7200
      TabIndex        =   26
      Text            =   "Pasword"
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3960
      TabIndex        =   25
      Text            =   "User ID"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdBan 
      Caption         =   "Ban Player"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdNotSuper 
      Caption         =   "Remove Super Admin Access"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSuper 
      Caption         =   "Make User Super Admin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdNotAdmin 
      Caption         =   "Remove Admin Access"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdmin 
      Caption         =   "Make User Admin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdNotModerator 
      Caption         =   "Remove     Mod Access"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdModerator 
      Caption         =   "Make User A Moderator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdPunt 
      Caption         =   "Punt Player"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   4440
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin SHDocVwCtl.WebBrowser wbReply 
      Height          =   2295
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Visible         =   0   'False
      Width           =   11775
      ExtentX         =   20770
      ExtentY         =   4048
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
   Begin VB.CheckBox cboMute 
      BackColor       =   &H00000000&
      Caption         =   "Mute Sounds?"
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
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Timer pingtimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7680
      Top             =   1920
   End
   Begin VB.CommandButton btnIM 
      Caption         =   "Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7200
      Top             =   1920
   End
   Begin VB.TextBox txtPWD 
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
      ForeColor       =   &H0000FF00&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   8160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox cboTalk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Talk?"
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
      Height          =   255
      Left            =   10800
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton btnAway 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Away"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00151515&
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton btnJoin 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Join Game"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      MaskColor       =   &H00151515&
      TabIndex        =   9
      ToolTipText     =   "Join a Game Room"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton btnHost 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Host Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00151515&
      TabIndex        =   8
      ToolTipText     =   "Create and Host a new Game Room"
      Top             =   4320
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox lstGameRooms 
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
      ForeColor       =   &H0011FF23&
      Height          =   3420
      Left            =   8160
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   3615
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
      ForeColor       =   &H0011FF23&
      Height          =   3660
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "List of Players Currently online"
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton btnUID 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Login/Register"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      MaskColor       =   &H00151515&
      TabIndex        =   3
      ToolTipText     =   "Log onto the Server"
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton btnSend 
      Appearance      =   0  'Flat
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      MaskColor       =   &H00151515&
      TabIndex        =   5
      ToolTipText     =   "Press this to send messages"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtUID 
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
      ForeColor       =   &H0011FF23&
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      ToolTipText     =   "enter Your User Name"
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton btnGo 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Connect"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MaskColor       =   &H00151515&
      TabIndex        =   0
      Top             =   120
      Width           =   855
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
      ForeColor       =   &H0011FF23&
      Height          =   405
      Left            =   120
      MaxLength       =   300
      TabIndex        =   4
      ToolTipText     =   "Type your Message here"
      Top             =   7680
      Visible         =   0   'False
      Width           =   10575
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8655
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   12495
      ExtentX         =   22040
      ExtentY         =   15266
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
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GameHost As String

Public MacAddy As String

Public game As String

Private FlagFocus As Boolean

Private lstReplyarray(100) As String

Public gameName As String
Public numberOfPlayers As Integer

Private lstArray(1000) As String
Private lstArrayIp(1000) As String
Private lstArrayLat(1000) As String
Public lstArrayCount As Integer


Private lstGameArrayIP(1000) As String
Private lstGameArrayPlayers(1000) As String
Private lstGameArrayNames(1000) As String
Private lstGameArrayGame(1000) As String
Private lstGameArrayGRPL(1000) As String

Public lstGameArrayCount As Integer

Public exposedIP As String

Public XVTGamePath As String
Public BOPGamePath As String
Public XWAGamePath As String

Public curVersion As String

Private IMUID As String

Private pingIndex As Integer

Private pingAll As Boolean

Private rtfReply As String

Private talk As Boolean

Public playing As Boolean

Private booted As Boolean

Private BannedUser(10000) As String
Public BannedUserCounter As Integer

Private Filter(10000) As String
Public FilterCount As Integer

Public Function getBannedUsers() As String()
    On Error GoTo Failed

    getBannedUsers = BannedUser
    Exit Function
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "getbanneduser " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

End Function

Public Function getFilter() As String()
    
    On Error GoTo Failed
    
    getFilter = Filter
    Exit Function
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "getfilter " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

End Function

Public Function does_exist(name As String) As Boolean
    
    On Error GoTo Failed
    
       does_exist = False
       For i = 0 To lstGameArrayCount - 1
            If name = lstGameArrayNames(i) Then does_exist = True
       Next i
       Exit Function
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "doesexist " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

End Function


Private Sub cmdAdmin_Click()
    On Error GoTo Failed
    If IMUID <> "" Then
        If MsgBox("Make " + IMUID + " an Administrator?", vbQuestion + vbYesNo, "Errant Venture Main Computer") = vbYes Then
            Socket1.SendLen = Len("MKAM" + IMUID + Chr$(3))
            Socket1.SendData = "MKAM" + IMUID + Chr$(3)
        End If
    End If
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "cmdadmin_click " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

End Sub

Private Sub cmdBan_Click()
    On Error GoTo Failed
    Dim reason As String

    If IMUID <> "" Then
        Do
            reason = InputBox("Reason For Banning" + IMUID + "?", "Errant Venture Main Computer", " ")
        Loop Until reason <> " "
        
        If reason <> "" Then
        
            Socket1.SendLen = Len("BNMC" + IMUID + Chr$(30) + reason + Chr$(3))
            Socket1.SendData = "BNMC" + IMUID + Chr$(30) + reason + Chr$(3)
        
            Call cmdPunt_Click
        End If
        
    End If
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "cmdban_click " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

End Sub



Private Sub cmdModerator_Click()
    On Error GoTo Failed
    If IMUID <> "" Then
        If MsgBox("Make " + IMUID + " a Moderator?", vbQuestion + vbYesNo, "Errant Venture Main Computer") = vbYes Then
            Socket1.SendLen = Len("MKMD" + IMUID + Chr$(3))
            Socket1.SendData = "MKMD" + IMUID + Chr$(3)
        End If
    End If
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "cmdModerator_Click " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

End Sub

Private Sub cmdNotAdmin_Click()
    On Error GoTo Failed
    If IMUID <> "" Then
        If MsgBox("remove " + IMUID + "'s Administrator abilities?", vbQuestion + vbYesNo, "Errant Venture Main Computer") = vbYes Then
            Socket1.SendLen = Len("RMAM" + IMUID + Chr$(3))
            Socket1.SendData = "RMAM" + IMUID + Chr$(3)
            Call cmdPunt_Click
        End If
    End If

    Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "cmdNotAdmin_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

End Sub

Private Sub cmdNotModerator_Click()
    On Error GoTo Failed
    If IMUID <> "" Then
        If MsgBox("Remove " + IMUID + "'s Moderator abilities?", vbQuestion + vbYesNo, "Errant Venture Main Computer") = vbYes Then
            Socket1.SendLen = Len("RMMD" + IMUID + Chr$(3))
            Socket1.SendData = "RMMD" + IMUID + Chr$(3)
            Call cmdPunt_Click
        End If
    End If
    Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "cmdNotModerator_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

End Sub

Private Sub cmdNotSuper_Click()
    On Error GoTo Failed
    If IMUID <> "" Then
        If MsgBox("Remove " + IMUID + "'s Super Administrator abilities?", vbQuestion + vbYesNo, "Errant Venture Main Computer") = vbYes Then
            Socket1.SendLen = Len("RMSU" + IMUID + Chr$(3))
            Socket1.SendData = "RMSU" + IMUID + Chr$(3)
            Call cmdPunt_Click
        End If
    End If
    Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "cmdNotSuper_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

End Sub

Private Sub cmdPunt_Click()
    On Error GoTo Failed
    If IMUID <> "" Then
        If MsgBox("Are you Sure you want to punt " + IMUID + "?", vbQuestion + vbYesNo, "Errant Venture Main Computer") = vbYes Then
            Socket1.SendLen = Len("BPWD" + IMUID + Chr$(3))
            Socket1.SendData = "BPWD" + IMUID + Chr$(3)
        End If
    End If
        Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "cmdPunt_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

End Sub

Private Sub cmdSuper_Click()
    On Error GoTo Failed
    If IMUID <> "" Then
        If MsgBox("Make " + IMUID + " a Super Administrator?", vbQuestion + vbYesNo, "Errant Venture Main Computer") = vbYes Then
            Socket1.SendLen = Len("MKSU" + IMUID + Chr$(3))
            Socket1.SendData = "MKSU" + IMUID + Chr$(3)
        End If
    End If
    Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "cmdSuper_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

End Sub

Private Sub cmdunBan_Click()
    On Error GoTo Failed
    Socket1.SendLen = Len("UNBN" + Chr$(3))
    Socket1.SendData = "UNBN" + Chr$(3)
        Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "cmdUnBan_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

End Sub

Private Sub Form_Load()
       
    Dim readhost As String
       
    On Error GoTo Failed
    booted = False
    
    ' Set properties needed by MCI to open.
    MMControl1.Notify = False
    MMControl1.Wait = False
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
   
    WebBrowser1.Navigate2 CurDir + "\connect.htm"

    wbReply.Navigate2 "about:blank"
    newMessage " "
    
    talk = False
    
    
    pingAll = True
    
    curVersion = "10.93"
    
    Dim uId As String
    Dim pwd As String
    
    Open CurDir + "\ipconfig.bat" For Append As #3
    Close #3
    Open CurDir + "\ipconfig.bat" For Output As #3
        Print #3, "cd \"
        Print #3, "ipconfig /all > " + Chr$(34) + CurDir + "\ipconfig.txt" + Chr$(34)
    Close #3
    
    
    
    
    Shell "ipconfig.bat", vbHide
    
    
    Open CurDir + "\xvtpwdlog.dat" For Append As #1
    Close #1
    
    
    Open CurDir + "\xvtpwdlog.dat" For Append As #1
    Close #1
    Open CurDir + "\xvtpwdlog.dat" For Input As #1
    If EOF(1) Then
        Close #1
        GoTo filenotfound
    End If
    Input #1, uId
    Input #1, pwd
    Close #1
    
    
    
    
    txtUID.Text = uId
    txtPWD.Text = pwd
    
    

filenotfound:
    
    Open CurDir + "\hosts.dat" For Append As #1
    Close #1
    Open CurDir + "\hosts.dat" For Input As #1
    
    Do Until EOF(1)
        Line Input #1, readhost
        txtHost.AddItem readhost
        
        If txtHost.Text = "" Then txtHost.Text = readhost

        
    Loop
       
    Close #1
    
    
    
    Me.Caption = "The NRSD Errant Venture - Version " + curVersion + "m"
    
    XVTGamePath = regQuery_A_Key(HKEY_LOCAL_MACHINE, "SOFTWARE\LucasArts Entertainment Company\X-Wing vs. TIE Fighter\1.0\", "Executable")
    BOPGamePath = regQuery_A_Key(HKEY_LOCAL_MACHINE, "SOFTWARE\LucasArts Entertainment Company\X-Wing vs. TIE Fighter\2.0\", "Executable")
    XWAGamePath = regQuery_A_Key(HKEY_LOCAL_MACHINE, "SOFTWARE\LucasArts Entertainment Company LLC\X-Wing Alliance\v1.0", "Executable")
    
    Socket1.AddressFamily = AF_INET
    Socket1.Protocol = IPPROTO_IP
    Socket1.SocketType = SOCK_STREAM
    Socket1.Binary = False
    Socket1.Blocking = False
    Socket1.BufferSize = 1024
    lstArrayCount = 0
    FlagFocus = True
    
    
   Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "from_load - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    
End Sub

Private Sub pingtimer_Timer()
    On Error GoTo Failed
    Dim ECHO As ICMP_ECHO_REPLY
    Dim lat As String

    ' on load this global variable will be true
    If pingAll = True And playing = False Then
        'set the timer to ping a new address every 30 seconds
        pingtimer.Interval = 30000
        
        lat = ""
        ' loop through all players
        For i = 0 To lstArrayCount - 1
            ' if the current player is not the player this client services then
            If Mid(UCase(lstArray(i)), 4) <> UCase(txtUID.Text) Then
               ' ping the player
               Call Ping(lstArrayIp(i), ECHO)
               ' get returned latancy time if any
               If ECHO.status = 0 Then
                    lat = " [" + Str(ECHO.RoundTripTime) & " ms]"
               Else
                    lat = " [N/A]"
               End If
                    
               'set the lat
               lstArrayLat(i) = lat
                    
            End If
        Next i
        
        ' set flag to refresh all to false
        pingAll = False
        
    
    End If
    

    ' otherwise set i to the global variable that tracks which ip it's pinging

    i = pingIndex
    ' initialize the lat
    lat = ""
    
            
    ' if current player is not the one client is servicing
    If Mid(UCase(lstArray(i)), 4) <> UCase(txtUID.Text) Then
        ' ping the player
       Call Ping(lstArrayIp(i), ECHO)
       ' get the results if there is any
       If ECHO.status = 0 Then
            lat = " [" + Str(ECHO.RoundTripTime) & " ms]"
        Else
            lat = " [N/A]"
        End If
        ' add it to the array
        lstArrayLat(i) = lat
            
    End If
    ' increment the global variable that keeps track of which player it's updating
    pingIndex = i + 1
    ' if it's greater than the number of players, reset it
    If pingIndex > lstArrayCount - 1 Then pingIndex = 0
    
    ' empty the listbox that the players live in
    lstPlayers.Clear
    'refill the list box
    For i = 0 To lstArrayCount - 1
        lstPlayers.AddItem lstArray(i) + lstArrayLat(i)
    Next i
    
   Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "ping timer - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    
    
End Sub

Private Sub btnAway_Click()
    On Error GoTo Failed
    If btnAway.Caption = "Away" Then
        btnAway.Caption = "Back"
        Socket1.SendLen = Len("^^^^" + "(A)" + txtUID.Text + Chr$(3))
        Socket1.SendData = "^^^^" + "(A)" + txtUID.Text + Chr$(3)
        btnHost.Enabled = False
        btnJoin.Enabled = False
            
    Else
        btnAway.Caption = "Away"
        Socket1.SendLen = Len("^^^^" + "(L)" + txtUID.Text + Chr$(3))
        Socket1.SendData = "^^^^" + "(L)" + txtUID.Text + Chr$(3)
        btnHost.Enabled = True
    
    End If
       Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnaway_click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub

Private Sub cboTalk_Click()
    
    On Error GoTo Failed
    If cboTalk.Value = 0 Then
        
        Unload frmTalk
    Else
        frmTalk.Show
        Me.SetFocus
    End If
       Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "cboTalk_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub





Private Sub btnGo_Click()
    On Error GoTo Failed
    Dim strng As String
    Dim i As Integer
    Dim tempstring As String
    Dim remPort As String
    
       
    On Error GoTo Failed
    
    frmClient.Caption = frmClient.Caption + " Connected to: " + txtHost.Text
    
    
    strng = Trim(txtUID)
    
    Open CurDir + "\port.dat" For Append As #1
    Close #1
    
    
    Open CurDir + "\port.dat" For Input As #1
        If Not EOF(1) Then Line Input #1, remPort
    Close #1
    
    If remPort = "" Then remPort = "2020"
    
    
    'On Error GoTo Failed
    Socket1.HostName = Trim$(txtHost.Text)
    Socket1.RemotePort = Val(remPort)
    Socket1.Action = SOCKET_CONNECT
    
    txtHost.Enabled = False
    btnGo.Enabled = False
    txtUID.Visible = True
    btnUID.Visible = True
    txtPWD.Visible = True
    btnUID.Default = True
    
    txtUID.SetFocus
    
    Text2.Visible = True
    Text3.Visible = True
    
    
    
   
    
    Open CurDir + "\ipconfig.txt" For Input As #3
    
    Do Until EOF(3)
    
        Line Input #3, tempstring
        If tempstring <> "" Then
            If Mid(tempstring, Len(tempstring) - 2, 1) = "-" Then
                MacAddy = Mid(tempstring, Len(tempstring) - 16)
            End If
        End If
    Loop
    
    Close #3
    
    'MacAddy = Socket1.PhysicalAddress
    

    
   Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnGo_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    


End Sub

Private Sub btnHost_Click()
    On Error GoTo Failed
    numberOfPlayers = 0
    On Error Resume Next
    
    frmConfigureGame.Show
    
    
       Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnHost_click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub

Private Sub btnJoin_Click()
    
    On Error GoTo Failed
    Dim indx As Integer
    
    If Mid(gameName, 5, 3) = "XWA" And XWAGamePath = "" Then
        MsgBox "That Gameroom is for X Wing Alliance, and you do not have it installed", vbOKOnly + vbInformation, "Game Not Installed"
    End If
    
    If Mid(gameName, 5, 3) = "BOP" And BOPGamePath = "" Then
        MsgBox "That Gameroom is for X Wing VS Tie Balance of Power, and you do not have it installed", vbOKOnly + vbInformation, "Game Not Installed"
    End If
    
    If Mid(gameName, 5, 3) = "XVT" And XVTGamePath = "" Then
        MsgBox "That Gameroom is for X Wing VS Tie, and you do not have it installed", vbOKOnly + vbInformation, "Game Not Installed"
    End If
    
    
    indx = lstGameRooms.ListIndex
    
    If Val(Mid(lstGameArrayPlayers(indx), 1, 1)) < Val(Mid(lstGameArrayPlayers(indx), 3, 1)) Then
    
    
        btnHost.Enabled = False
        btnJoin.Enabled = False
        btnAway.Enabled = False
        'lstGameRooms.Enabled = False
        
         
        frmJoinGame.Show
    Else
        MsgBox "That Room is full", vbOKOnly, "Sorry"
    End If
    
       Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnJoin_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub

Private Sub btnUID_Click()
    On Error GoTo Failed
        Dim parse As String
        Dim name As String
        
        btnUID.Enabled = False
        txtUID.Enabled = False
        txtPWD.Enabled = False
        
    
        
        name = "(L)" + txtUID.Text

        Socket1.SendLen = Len("####" + name + Chr$(30) + txtPWD.Text + Chr$(3))
        Socket1.SendData = "####" + name + Chr$(30) + txtPWD.Text + Chr$(3)
        
        Timer1.Enabled = True
 
       Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnUID_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    


End Sub

Private Sub lstGameRooms_Click()
    On Error GoTo Failed
    Dim tempstring As String
    
    If lstGameRooms.Text <> "" Then
        tempstring = Replace(lstGameArrayGRPL(lstGameRooms.ListIndex), ", ", vbCrLf)
        AddCustomToolTip lstGameRooms, tempstring, frmClient
    
        
        If playing = False Then
            
            GameHost = lstGameArrayIP(lstGameRooms.ListIndex)
            numberOfPlayers = Val(Mid(lstGameArrayPlayers(lstGameRooms.ListIndex), 3))
            gameName = lstGameArrayNames(lstGameRooms.ListIndex)
            If GameHost <> "" And btnAway.Caption = "Away" Then
                btnJoin.Enabled = True
            Else
                btnJoin.Enabled = False
            End If
        End If
        
    
    End If
       Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "lstGameRooms_click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub


Private Sub lstPlayers_Click()
    On Error GoTo Failed
    
    lstPlayers.ToolTipText = lstArrayIp(lstPlayers.ListIndex)
    
    If lstPlayers.ListIndex = -1 Then Exit Sub
    IMUID = Mid(lstArray(lstPlayers.ListIndex), 4)
 
    If IMUID <> "" Then
    
       btnIM.Enabled = True
       
    
    End If
          Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "lstPlayers_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub
Private Function GetRidofHTML(Str As String)
    'MsgBox Str
    If InStr(1, UCase(frmClient.Caption), "SUPER ADMINISTRATOR") = 0 Then
            Str = Replace(Str, "<", "&lt;")
            Str = Replace(Str, ">", "&gt;")
    End If
    'MsgBox Str
    GetRidofHTML = Str
End Function
Private Sub lstPlayers_DblClick()
    On Error GoTo Failed
   Dim imMsg As String
    
    If lstPlayers.ListIndex = -1 Then Exit Sub
    
    IMUID = Mid(lstArray(lstPlayers.ListIndex), 4)
    If IMUID <> "" And IMUID <> txtUID.Text Then
        imMsg = InputBox("Page User: " + IMUID, "NRSD Errant Venture Comm System")
        
        Socket1.SendLen = Len("IMIM" + txtUID.Text + Chr$(30) + IMUID + Chr$(30) + imMsg + Chr$(3))
        Socket1.SendData = "IMIM" + txtUID.Text + Chr$(30) + IMUID + Chr$(30) + imMsg + Chr$(3)
        newMessage "<font color='#664423'>page to</font> " & IMUID & Chr(198) & "- " & imMsg
        btnIM.Enabled = False
        
    End If
    Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "lstPlayers_DblClick - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub
Private Sub btnIM_Click()
    On Error GoTo Failed
    Dim imMsg As String
    
    If IMUID <> "" And IMUID <> txtUID.Text Then
        imMsg = InputBox("Page User: " + IMUID, "NRSD Errant Venture Comm System")
        
        If imMsg <> "" Then
            imMsg = GetRidofHTML(imMsg)
            Socket1.SendLen = Len("IMIM" + txtUID.Text + Chr$(30) + IMUID + Chr$(30) + imMsg + Chr$(3))
            Socket1.SendData = "IMIM" + txtUID.Text + Chr$(30) + IMUID + Chr$(30) + imMsg + Chr$(3)
            newMessage "<font color='#664423'>page from</font> " & IMUID & Chr(198) & " - " & imMsg
        End If
        btnIM.Enabled = False
    
    End If
   Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnIM_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim strBuffer As String
    Dim first4 As String
    Dim parsed As String
    Dim cnt As Integer
    Dim lstarraytemp() As String
    Dim lstarrayiptemp() As String
    Dim lstarrayLattemp() As String
    Dim B As Integer
    Dim numMsgs As Integer
    Dim msgs(1000) As String
    Dim nextMsgStart As Integer
    Dim wrapLines(100) As String
    Dim w As Integer
    Dim result As String
    Dim game As String
    Dim uidIM As String
    Dim msgIM As String
    Dim replyIM As String
    Dim touidIM
    Dim GRPLGamename As String
    Dim GRPLString As String
    Dim playerip As String
    
    Dim args$()
    
    On Error GoTo Failed
    
    
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
        
        
        
        Select Case first4
        Case "BPWD":
        
            'If Socket1.Connected Then
            '    Socket1.SendLen = Len("@@@@" + txtuid.Text + chr$(3))
            '    Socket1.SendData = "@@@@" + txtuid.Text + chr$(3)
            '    Socket1.Action = SOCKET_CLOSE
            'End If
            'MsgBox "You Have been logged in somewhere else", vbExclamation + vbOKOnly, "Warning"

          booted = True

          Unload Me
          
        Case "IMIM":
            If playing = True Then
                Socket1.SendLen = Len("IMIM" + "Errant Venture Paging System" + Chr$(30) + uidIM + Chr$(30) + "The user you are trying to page is unavallible at the moment, please try again later" + Chr$(3))
                Socket1.SendData = "IMIM" + "Errant Venture Paging System" + Chr$(30) + uidIM + Chr$(30) + "The user you are trying to page is unavallible at the moment, please try again later" + Chr$(3)
            
            Else
                Dim team As Boolean
                team = False
                parsed = Mid(strBuffer, 5)
                args$() = Split(parsed, Chr$(30))
                If UBound(args$) = 3 Then If args$(3) <> "" Then team = True
                
                uidIM = Mid(parsed, 1, InStr(parsed, Chr$(30)) - 1)
                msgIM = Mid(parsed, InStr(InStr(parsed, Chr$(30)) + 1, parsed, Chr$(30)) + 1)
                PlayWave App.path + "/bing.wav"
                If team = True Then
                    args(2) = GetRidofHTML(args(2))
                    newMessage "<font color='#FF0000'>team chat * " & args$(3) & " *</font> " & uidIM & Chr(198) & " - " & args(2)
                Else
                    msgIM = GetRidofHTML(msgIM)
                    newMessage "<font color='#664423'>page from</font> " & uidIM & Chr(198) & " - " & msgIM
                    Do
                        replyIM = InputBox(uidIM + " Says: " + msgIM, "Errant Venture Comm System", " ")
                    Loop Until replyIM <> " "
                    If replyIM <> "" Then
                        replyIM = GetRidofHTML(replyIM)
                        Socket1.SendLen = Len("IMIM" + txtUID.Text + Chr$(30) + uidIM + Chr$(30) + replyIM + Chr$(3))
                        Socket1.SendData = "IMIM" + txtUID.Text + Chr$(30) + uidIM + Chr$(30) + replyIM + Chr$(3)
                        newMessage "<font color='#664423'>page to</font> " & uidIM & Chr(198) & " - " & replyIM
                    End If
                End If
            End If
        Case "VERQ"
            Dim wr$
            wr$ = "VERR" & txtUID.Text & Chr$(30) & "ORG_" & curVersion & "m" & Chr$(3)
            Socket1.SendLen = Len(wr$)
            Socket1.SendData = wr$
        Case "PICR"
            parsed = Mid(strBuffer, 5)
            args$() = Split(parsed, Chr$(30))
            args$(1) = GetRidofHTML(args$(1))
            newMessage args$(0) & " is <img src='" & args$(1) & "' width='32' height='32' align=middle> "
        Case "GRPL"
        
            parsed = Mid(strBuffer, 5)
            
            GRPLGamename = Mid(parsed, 1, InStr(parsed, Chr$(30)) - 1)
            GRPLString = Mid(parsed, InStr(parsed, Chr$(30)) + 1)
            
            For i = 0 To lstGameArrayCount - 1
                If lstGameArrayNames(i) = GRPLGamename Then
                    lstGameArrayGRPL(i) = GRPLString
                End If
            Next i
            
        
        
        
        Case "LLLL":
        
            parse = Mid(strBuffer, 5)
            
            ok = Mid(parse, 1, 1)
            
            If ok = "1" Then
            
              
                If talk Then cboTalk.Visible = True
                
                txtUID.Enabled = False
                txtPWD.Enabled = False
                btnUID.Enabled = False
                
                btnUID.Visible = False
                txtUID.Visible = False
                txtPWD.Visible = False
                txtHost.Visible = False
                btnGo.Visible = False
                
                
                btnAway.Visible = True
                btnIM.Visible = True
                btnIM.Enabled = False
                
                
                txtSend.Visible = True
                btnSend.Visible = True
                btnSend.Default = True
                
                wbReply.Visible = True
                lstPlayers.Visible = True
                 
                lstGameRooms.Visible = True
                btnHost.Visible = True
                btnJoin.Visible = True
                
                WebBrowser1.Navigate2 CurDir + "\lobby.htm"

                txtSend.SetFocus
                
                Text1.Visible = False
                Text2.Visible = False
                Text3.Visible = False
                
                Text4.Visible = True
                Text5.Visible = True
                Text6.Visible = True
                
                
                
                pingtimer.Enabled = True
                 
       
                
                Open CurDir + "\xvtpwdlog.dat" For Output As #1
                Print #1, txtUID.Text
                Print #1, txtPWD.Text
                Close #1
                
            Socket1.SendLen = Len("PICQ" + Chr$(3) + "VERQ" + Chr$(3))
            Socket1.SendData = "PICQ" + Chr$(3) + "VERQ" + Chr$(3)
            
            ElseIf ok = 0 Then
            
                result = MsgBox("Username or Password either does not exist or not registered. " + Chr$(13) + "Press cancel to try again, or Press Ok to register in system", vbExclamation + vbOKCancel, "Login error")
                If result = 1 Then
                    frmRegister.Show
                    Me.Enabled = False
                    
                End If
                
                btnUID.Enabled = True
                txtUID.Enabled = True
                txtPWD.Enabled = True
 
            Else
            
                MsgBox "Already logged in, Logging off other instance", vbExclamation + vbOKOnly, "Already Logged in"
                GoSleepEX 3
                btnUID_Click
                
            End If
            
        Case "RUSR":
            
            parsed = Mid(strBuffer, 5)
            
            If parsed = "0" Then
                MsgBox "User Already exists, Try another", vbOKOnly + vbExclamation, "Validation Error"
                frmRegister.Enabled = True
            Else
                MsgBox "You have successfully registered", vbOKOnly + vbInformation, "Succesful Registration"
                Unload frmRegister
                Me.Enabled = True
            End If
            
            
            
        Case "VVVV":
        
            parsed = Mid(strBuffer, 5)
            
            If parsed <> curVersion Then
            
                MsgBox "this is an old version, please download the latest from http://ccanuk.brinkster.net/xvt/redist.zip", vbCritical + vbOKOnly, "Old Version"
                Shell Chr$(34) + "c:\program files\internet explorer\iexplore.exe" + Chr$(34) + " http://ccanuk.brinkster.net/xvt/redist.zip", vbNormalFocus
                
                'Unload Me
            
            End If
         
        Case "####":
            
            If cboMute.Value <> 1 And pingAll = False And playing = False Then PlayWave App.path + "/joinlobby.wav"
            parsed = Mid(strBuffer, 5)
        
        
            playerip = Mid(parsed, InStr(parsed, Chr$(30)) + 1)
            parsed = Mid(parsed, 1, InStr(parsed, Chr$(30)) - 1)
        
        
            lstPlayers.AddItem parsed
            lstArray(lstArrayCount) = parsed
            lstArrayIp(lstArrayCount) = playerip
            
            lstArrayCount = lstArrayCount + 1
            
            newMessage Mid(parsed, 4) + " Enters"
            
            btnIM.Enabled = False
            
    
        
        Case "@@@@":
            
            If cboMute.Value <> 1 And playing = False Then PlayWave App.path + "/leavelobby.wav"
            parsed = Mid(strBuffer, 5)
            
            cnt = 0
                    
            lstPlayers.Clear
            
            
            ReDim lstarraytemp(lstArrayCount)
            ReDim lstarrayiptemp(lstArrayCount)
            ReDim lstarrayLattemp(lstArrayCount)
            
            For i = 0 To lstArrayCount - 1
            
                If parsed <> Mid(lstArray(i), 4) Then
                    lstPlayers.AddItem lstArray(i) + lstArrayLat(i)
                    lstarraytemp(cnt) = lstArray(i)
                    lstarrayiptemp(cnt) = lstArrayIp(i)
                    lstarrayLattemp(cnt) = lstArrayLat(i)
                    cnt = cnt + 1
                End If
                
            
            Next i
            
            lstArrayCount = cnt
            For i = 0 To lstArrayCount
                lstArray(i) = lstarraytemp(i)
                lstArrayIp(i) = lstarrayiptemp(i)
                lstArrayLat(i) = lstarrayLattemp(i)
            Next i
        
            newMessage parsed + " Leaves"
        
            pingIndex = 0
        
           Case "%%%%":
        
            Dim lstGameArrayIPtemp() As String
            Dim lstGameArrayPlayerstemp() As String
            Dim lstGameArrayNamestemp() As String
            Dim lstGameArrayGametemp() As String
            Dim lstGameArrayGRPLtemp() As String
        
            parsed = Mid(strBuffer, 5)
            
            cnt = 0
                    
            lstGameRooms.Clear
            'gameName = ""
            btnJoin.Enabled = False
            
            
            ReDim lstGameArrayIPtemp(lstGameArrayCount)
            ReDim lstGameArrayPlayerstemp(lstGameArrayCount)
            ReDim lstGameArrayNamestemp(lstGameArrayCount)
            ReDim lstGameArrayGametemp(lstGameArrayCount)
            ReDim lstGameArrayGRPLtemp(lstGameArrayCount)
            
            
            For i = 0 To lstGameArrayCount - 1
            
                If parsed <> lstGameArrayNames(i) Then
                    lstGameRooms.AddItem (Str(i + 1) + " " + lstGameArrayPlayers(i) + " " + lstGameArrayGame(i) + " " + Mid(lstGameArrayNames(i), 1, 30))

                    lstGameArrayIPtemp(cnt) = lstGameArrayIP(i)
                    lstGameArrayPlayerstemp(cnt) = lstGameArrayPlayers(i)
                    lstGameArrayNamestemp(cnt) = lstGameArrayNames(i)
                    lstGameArrayGametemp(cnt) = lstGameArrayGame(i)
                    lstGameArrayGRPLtemp(cnt) = lstGameArrayGRPL(i)
                    cnt = cnt + 1
    
                End If
                
            
            Next i
            
            
            lstGameArrayCount = cnt
            For i = 0 To lstGameArrayCount - 1
                    lstGameArrayIP(i) = lstGameArrayIPtemp(i)
                    lstGameArrayPlayers(i) = lstGameArrayPlayerstemp(i)
                    lstGameArrayNames(i) = lstGameArrayNamestemp(i)
                    lstGameArrayGame(i) = lstGameArrayGametemp(i)
                    lstGameArrayGRPL(i) = lstGameArrayGRPLtemp(i)
            Next i
            
        
        Case "++++":
            parsed = Mid(strBuffer, 5)
            
            lstGameArrayPlayers(lstGameArrayCount) = Mid(parsed, 1, 3)
            lstGameArrayGame(lstGameArrayCount) = Mid(parsed, 5, 3)
            lstGameArrayNames(lstGameArrayCount) = Mid(parsed, 9, 30)
            lstGameArrayIP(lstGameArrayCount) = Mid(parsed, 39)
            
            lstGameArrayCount = lstGameArrayCount + 1
    
            lstGameRooms.AddItem (Str(i + 1) + " " + Mid(parsed, 1, 38))
            
        
        Case "^^^^":
        
        
            
    
            
            
        
            parsed = Mid(strBuffer, 5)
            first3 = Mid(parsed, 1, 3)
            
            parsed = Mid(parsed, 4)
                    
                lstPlayers.Clear
                For i = 0 To lstArrayCount - 1
                    If parsed = Mid(lstArray(i), 4) Then
                        
                        
                        lstArray(i) = first3 + parsed
                        
                        
                    End If
                    lstPlayers.AddItem lstArray(i) + lstArrayLat(i)
                Next i
        Case "^++^":
        
            first3 = Mid(strBuffer, 5, 3)
            game = Mid(strBuffer, 9, 3)
            parsed = Mid(strBuffer, 13)
            
            lstGameRooms.Clear
            For i = 0 To lstGameArrayCount - 1
            
                If lstGameArrayNames(i) = parsed Then
                    lstGameArrayPlayers(i) = first3
                    lstGameArrayGame(i) = game
                    
                End If
                lstGameRooms.AddItem Str(i + 1) + " " + lstGameArrayPlayers(i) + " " + lstGameArrayGame(i) + " " + lstGameArrayNames(i)
                
            
            Next i
        
        
        Case "%++%":
        
            ' tell the server what the Mac Address is
            Socket1.SendLen = Len("MCAD" + MacAddy + Chr$(3))
            Socket1.SendData = "MCAD" + MacAddy + Chr$(3)
            
            exposedIP = Mid(strBuffer, 5)
            
            
            Open CurDir + "\OR.dat" For Append As #3
            Close #3
            Open CurDir + "\OR.dat" For Input As #3
            Do Until EOF(3)
                Line Input #3, exposedIP
                Socket1.SendLen = Len("ORIP" + exposedIP + Chr$(3))
                Socket1.SendData = "ORIP" + exposedIP + Chr$(3)

            Loop
            Close #3
            
                       
        Case "MMMM":
            
            If cboMute.Value <> 1 And pingAll = False And playing = False Then PlayWave App.path + "/message.wav"
        
            parsed = Mid(strBuffer, 5)
            
            If talk Then
                On Error Resume Next
                If cboTalk.Value = 1 Then
                    frmTalk.DirectSS1.Sayit = parsed
                End If
            End If
            parsed = GetRidofHTML(parsed)
             newMessage parsed
        
        Case "ADMN"
        
            parsed = Mid(strBuffer, 5)
            
            'newMessage "Your Status is " + parsed
            frmClient.Caption = frmClient.Caption + " " + parsed
        
            'here is where we get to the new rights if you have an elevated status
            
            If UCase(parsed) = "SUPERADMIN" Then
            
                cmdPunt.Visible = True
                cmdBan.Visible = True
                cmdunBan.Visible = True
                cmdModerator.Visible = True
                cmdNotModerator.Visible = True
                cmdAdmin.Visible = True
                cmdNotAdmin.Visible = True
                cmdSuper.Visible = True
                cmdNotSuper.Visible = True
                
            ElseIf UCase(parsed) = "ADMIN" Then
                cmdPunt.Visible = True
                cmdBan.Visible = True
                cmdunBan.Visible = True
                cmdModerator.Visible = True
                cmdNotModerator.Visible = True
                'cmdAdmin.Visible = True
            
            ElseIf UCase(parsed) = "MODERATOR" Then
                cmdPunt.Visible = True
                cmdBan.Visible = True
                cmdunBan.Visible = True
            End If
                
            
        Case "BNED"
        
            parsed = Mid(strBuffer, 5)
            
            BannedUser(BannedUserCounter) = parsed
            BannedUserCounter = BannedUserCounter + 1
        
        Case "BNND"
        
            frmViewBanned.Show
            BannedUserCounter = 0
            
        Case "FFFF"
        
            parsed = Mid(strBuffer, 5)
            
            FilterCount = FilterCount + 1
            
            Filter(FilterCount) = parsed
            
        
            
        End Select
       
    Next B
       Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "socket1_read - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub

Private Sub newMessage(msg As String)
    On Error GoTo Failed

Dim cursestart
    If cboFilterOff.Value <> 1 Then
        For i = 0 To FilterCount
            msg = Replace(msg, Filter(i), "Smurf", , , vbTextCompare)
        Next i
    End If
    'If InStr(1, UCase(frmClient.Caption), "SUPER ADMINISTRATOR") = 0 Then
      '      msg = Replace(msg, "<", "")
     '       msg = Replace(msg, ">", "")
    'End If
    If InStr(rtfReply, msg) And InStr(msg, "<img") Then
        'MsgBox "DOUBLE BARREL!"
        Exit Sub
    End If

    If Len(rtfReply) > 32000 Then rtfReply = Mid(rtfReply, Len(rtfReply) - 32000)

    If InStr(msg, Chr(198)) <> 0 Then
        
        rtfReply = rtfReply + "<B>"
        If Mid(msg, 1, InStr(msg, Chr(198)) - 1) = txtUID.Text Then
            rtfReply = rtfReply + "<font color=#00FF33>"
        Else
            rtfReply = rtfReply + "<font color=#3300FF>"
        End If
        
        rtfReply = rtfReply + Mid(msg, 1, InStr(msg, Chr(198)) - 1)
        rtfReply = rtfReply + "</B>"
        rtfReply = rtfReply + "<font color=#C0C0C0>"
        rtfReply = rtfReply + Mid(msg, InStr(msg, Chr(198)) + 1) + "<br>"
    
    Else
        
        rtfReply = rtfReply + "<font color=#C0C0C0>"
        rtfReply = rtfReply + Mid(msg, InStr(msg, Chr(198)) + 1) + "<br>"
    End If
    

    Open CurDir + "\chat.htm" For Output As #2
    
    Print #2, "<HTML><BODY BGCOLOR=Black><font size=-1 face=arial>" + rtfReply + "<a name=ender></a></body></html>"
    
    Close #2
    
    If playing = False Then wbReply.Navigate2 CurDir + "\chat.htm#ender"
    

   Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "newmessage - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    


End Sub


Private Sub btnSend_Click()

    On Error GoTo Failed
        
        Dim parse As String
        
        If txtSend <> "" Then
        
            parse = "MMMM" + txtUID.Text + Chr(198) & " - " + txtSend.Text
         
        
            Socket1.SendLen = Len(parse + Chr$(3))
            Socket1.SendData = parse + Chr$(3)
            KeyAscii = 0: txtSend.Text = ""
        End If
        
    Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnSend_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub

Sub Form_Unload(Cancel As Integer)
    On Error GoTo Failed
    
    Timer1.Enabled = False
    
     
    On Error Resume Next
    
    Unload frmHostGame
    Unload frmJoinGame
    
    
    If Socket1.Connected Then
        If btnUID.Enabled = True Then
            Socket1.SendLen = Len("MMMM" + txtUID.Text + " failed to log in" + Chr$(3))
            Socket1.SendData = "MMMM" + txtUID.Text + " failed to log in" + Chr$(3)
        
        Else
            Socket1.SendLen = Len("@@@@" + txtUID.Text + Chr$(3))
            Socket1.SendData = "@@@@" + txtUID.Text + Chr$(3)
        End If
        Socket1.Action = SOCKET_CLOSE
    End If
    
    If booted = True Then
        'MsgBox "Connection Terminated - either you have logged in somewhere else or a moderator booted you", vbInformation, "Errant Venture Main Computer"
    End If
    End
    
       Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "from_unload - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub






Private Sub Timer1_Timer()
    On Error GoTo Failed

    Dim i As Integer
    
    If Socket1.Connected = False Then
        'Socket1.Action = SOCKET_CONNECT
        Socket1.Connect
        
        newMessage "Lost Connection - Attempting to reconnect to the server"
    End If
       Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "timer1_timer - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub

Private Sub txtHost_KeyPress(KeyAscii As Integer)
    On Error GoTo Failed
    If KeyAscii = Asc(Chr$(3)) Then KeyAscii = 0
    
    For i = 0 To FilterCount

        If cboFilterOff.Value <> 1 Then
        
            msg = Replace(txtHost.Text, Filter(i), "Smurf", , , vbTextCompare)
        
        End If

    Next i
    
    
    
   Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "txtHost_Keypress - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub


Private Sub txtSend_KeyPress(KeyAscii As Integer)
    On Error GoTo Failed
  If KeyAscii = Asc(Chr$(3)) Then KeyAscii = 0
    
   Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "txtSend_KeyPress - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    
    
End Sub


Private Sub txtUID_Change()
    On Error GoTo Failed
    If Len(txtUID.Text) > 2 Then
        btnUID.Enabled = True
    End If
    
    For i = 0 To frmClient.FilterCount

        If cboFilterOff.Value <> 1 Then
        
            txtUID.Text = Replace(txtUID.Text, Filter(i), "Smurf", , , vbTextCompare)
        
        End If

    Next i
   Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "txtUID_change - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub

Private Sub txtUID_KeyPress(KeyAscii As Integer)
    On Error GoTo Failed
  If KeyAscii = Asc(Chr$(3)) Then KeyAscii = 0
   Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "txtUID_KeyPress - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub
Public Sub PlayWave(path As String)
    On Error GoTo Failed
    MMControl1.FileName = path
    
    ' Open the MCI WaveAudio device.
    
    MMControl1.Command = "Close"
    MMControl1.Command = "Back"
    MMControl1.Command = "Open"
    MMControl1.Command = "Play"
    
       Exit Sub

Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "PlayWave - " + Str(Err.Number) + " - " + Err.Description + " - frmclient"
    Close #4
    Resume Next

    

End Sub

