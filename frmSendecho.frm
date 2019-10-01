VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Begin VB.Form frmServer 
   Caption         =   "XvT Server"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin SocketWrenchCtrl.Socket Socket2 
      Index           =   0
      Left            =   3120
      Top             =   240
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
   Begin VB.VScrollBar VScroll1 
      Height          =   375
      LargeChange     =   10
      Left            =   4320
      Max             =   90
      Min             =   10
      TabIndex        =   5
      Top             =   3120
      Value           =   10
      Width           =   375
   End
   Begin VB.TextBox txtGameRoomTimeout 
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "10"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Timer timercheckgamerooms 
      Interval        =   2000
      Left            =   2400
      Top             =   1320
   End
   Begin VB.CommandButton btnReloadPWL 
      Caption         =   "Reload Password list"
      Height          =   735
      Left            =   2400
      TabIndex        =   3
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Timer Timercheckconnections 
      Interval        =   2000
      Left            =   5400
      Top             =   1320
   End
   Begin VB.ListBox lstgamerooms 
      Height          =   3375
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.ListBox lstPlayers 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox lstReply 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   9375
   End
   Begin VB.Label Label1 
      Caption         =   "Game Room Keep Alive Timeout"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "frmserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this is an array that holds the last 100 messages
Private lstReplyarray(100) As String
' this is an array that holds the names of the players
Private lstArray(10000) As String
'this is an array that holds the socket index that the corresponding player is connected with
Private lstArraySocketIndex(10000) As Integer
'this is the current number of players
Private lstArrayCount As Integer

' this is an array of IP addresses for game hosts
Private lstGameArrayIP(10000) As String
' this is an array of player counts for gamerooms - ie "1/4", "3/8"
Private lstGameArrayPlayers(10000) As String
' this is an array of game names
Private lstGameArrayNames(10000) As String
' this is an array of what game the room is supporting (XVT, XWA...)
Private lstGameArrayGame(10000) As String
' this is an array of sockets that the host is connected to the server with
Private lstGameArraySocketIndex(10000) As Integer
' this is an array of GameRoomPlayerLists... this holds the tool text tip string
Private lstGameArrayGRPL(10000) As String
' this is an array of dates that is updated by the pinging of the gameroom host to the server
Private lstGameArrayKeepAlive(10000) As Date
' this is stores the current number of hosted gamerooms
Private lstGameArrayCount As Integer


' this is a serverwide variable that stores how many seconds a gameroom can go
'without pinging before it's removed from the list
Private KeepAliveTimeout As Integer

' this is an array of passwords and userid's that the server loads on startup
Private uidpwdArray(10000) As String

'this is a variable that holds the current version of the server
Private curVersion As String

'not sure if this is used
Private LastSocket As Integer

Private Sub Form_Load()
    'executed on the load of the server
    On Error GoTo ender
    
    Dim i As Integer
    i = 0
    
    'set the current version
    curVersion = "10.4"
    ' set the caption on the server
    frmserver.Caption = "Errant Venture Server - Version " + curVersion
    
    'set the gameroom timeout
    KeepAliveTimeout = txtGameRoomTimeout.Text
    
    'open the password file
    Open "c:\pwdfile.dat" For Append As #1
    Close #1
    Open "c:\pwdfile.dat" For Input As #1
    'load the contents of the file into the array
    For i = 0 To 10000
        'if reaches the end before 10000 records then exit for loop
        If EOF(1) Then Exit For
        Line Input #1, uidpwdArray(i)
    
    Next i
    
    ' close the file
    Close #1
    
    ' set up the listening socket
    Socket2(0).AddressFamily = AF_INET
    Socket2(0).Protocol = IPPROTO_IP
    Socket2(0).SocketType = SOCK_STREAM
    Socket2(0).Blocking = False
    Socket2(0).LocalPort = 2020
    Socket2(0).Action = SOCKET_LISTEN
    LastSocket = 0

    ' set the list of players to nothing
    lstArrayCount = 0
    
    ' exit sub
    Exit Sub
    
ender:
    ' if code gets here, something broke, log it
    Open "c:\EVerrorlog " + Str(Day(Now())) + "-" + Str(Month(Now())) + "-" + Str(Year(Now())) + ".log" For Append As #1
    Print #1, "++++FormLoad" + " - " + Str(Now())
    Print #1, "++++" + Str(Err.Number) + " - " + Err.Description
    Close #1
End Sub

Private Sub btnReloadPWL_Click()
' this code is so that when there is a change in the pwd file list, ie changed password
' you can reload the list onto the server's memory
On Error GoTo ender
                'open the file
                Open "c:\pwdfile.dat" For Input As #1
                'grab the first 10000 entries
                For i = 0 To 10000
                    
                    If EOF(1) = True Then Exit For
                    Line Input #1, uidpwdArray(i)
                    i = i + 1
                
                Next i
                                
                Close #1
                
ender:
    
    Open "c:\EVerrorlog " + Str(Day(Now())) + " -" + Str(Month(Now())) + " -" + Str(Year(Now())) + ".log" For Append As #1
    Print #1, "++++btnReloadPWL_Click" + " - " + Str(Now())
    Print #1, "++++" + Str(Err.Number) + " - " + Err.Description
    Close #1



End Sub
Private Sub Socket2_Accept(index As Integer, SocketId As Integer)
    'this sub is for handling connections from clients
On Error GoTo ender
    
    ' loop through the number of sockets
    Dim i As Integer
    For i = 1 To LastSocket
        ' if it's not connected, use it
        If Not Socket2(i).Connected Then Exit For
    Next i
    ' if the socket is a new not a reused one
    If i > LastSocket Then
        'add a new socket
        LastSocket = LastSocket + 1: i = LastSocket
        'load the socket into memory
        Load Socket2(i)
    End If
    'set the socket up to connect
    Socket2(i).AddressFamily = AF_INET
    Socket2(i).Protocol = IPPROTO_IP
    Socket2(i).SocketType = SOCK_STREAM
    Socket2(i).Binary = True
    Socket2(i).BufferSize = 1024
    Socket2(i).Blocking = False
    Socket2(i).Accept = SocketId
    
    
    ' send to the socket a message, telling the client what IP is exposed to the internet
    parsed = "%++%" + Socket2(i).PeerAddress + "�"
    
    Socket2(i).SendLen = Len(parsed)
    Socket2(i).SendData = parsed
    
    'tell the client what version of the server is running
    Socket2(i).SendLen = Len("VVVV" + curVersion + "�")
    Socket2(i).SendData = "VVVV" + curVersion + "�"
    
    Exit Sub
ender:
    Open "c:\EVerrorlog " + Str(Day(Now())) + " -" + Str(Month(Now())) + " -" + Str(Year(Now())) + ".log" For Append As #1
    Print #1, "++++Socket2_accept" + " - " + Str(Now())
    Print #1, "++++" + Str(Err.Number) + " - " + Err.Description
    Close #1

    
End Sub
Private Sub Socket2_Read(index As Integer, DataLength As Integer, IsUrgent As Integer)
    ' this deals with messages sent from the clients
    Dim strdata As String
    Dim first4 As String
    Dim parsed As String
    Dim cnt As Integer
    Dim B As Integer
    Dim numMsgs As Integer
    Dim msgs(100) As String
    Dim nextMsgStart As Integer
    Dim Wraplines(10) As String
    Dim w As Integer
    Dim lstarraytemp() As String
    Dim lstarraysocketindextemp() As Integer
    Dim uidpwd As String
    Dim goflag As Boolean
    Dim Game As String
    Dim uidIM As String
    Dim curnow As Date
    
' set up error catching
On Error GoTo ender
' initialize the recieved message start point
    nextMsgStart = 1
' initialize the number of messages
    numMsgs = 0
' read the recieved data from the buffer
    Socket2(index).Read strdata, DataLength
    
' if the last char is not an ending char, then add the ending char
    If Mid(strdata, Len(strdata)) <> "�" Then
        strdata = strdata + "�"
    End If
    
' loop through the string,
    For B = 1 To Len(strdata)
    'if you find an ending char, then cut out message and put it into the message array
        If Mid(strdata, B, 1) = "�" Then
            msgs(numMsgs) = Mid(strdata, nextMsgStart, B - nextMsgStart)
            ' increment the number of messages
            numMsgs = numMsgs + 1
            ' set the next char to the one after the current end point
            nextMsgStart = B + 1
        End If
    Next B
    
    
' loop through the recieved messages
    For B = 0 To numMsgs - 1
    
        
    '   set this variable to the message looped to
        strdata = msgs(B)
        
        'open the log file, append the current message to it and close the file
        Open "c:\EVerrorlog " + Str(Day(Now())) + " -" + Str(Month(Now())) + " -" + Str(Year(Now())) + ".log" For Append As #1
        Print #1, "    " + Str(Now()) + " - " + strdata
        Close #1
        
        ' get the signalling chars from the message
        first4 = Mid(strdata, 1, 4)
        
        ' determine the action to be taken
        Select Case first4
        
        Case "####": ' add new player
                
                'get the string minus the signalling characters
                parsed = Mid(strdata, 5)
                ' get the username pasword, minus the status indicator (L) or (G) etc
                uidpwd = Mid(parsed, 4)
                
                ' loop though the string looking for the username/password seperator "�"
                For i = 1 To Len(parsed)
                    If Mid(parsed, i, 1) = "�" Then
                        parsed = Mid(parsed, 1, i - 1)
                        'grab the uid, and formulate the new login to send to all connected clients
                        strdata = "####" + parsed + "�" + Socket2(index).PeerAddress
                        Exit For
                    End If
                    
                Next i
                
                ' attempt to log in the player
                If login(uidpwd) Then
                    ' loop though the list of players connected
                    For i = 0 To lstArrayCount - 1
                        'check to see if the username is already logged in somewhere
                        If Mid(UCase(parsed), 4) = Mid(UCase(lstArray(i)), 4) Then
                            ' if it is, send the log out signal to it
                            Socket2(lstArraySocketIndex(i)).SendLen = Len("BPWD" + lstArray(i) + "�")
                            Socket2(lstArraySocketIndex(i)).SendData = "BPWD" + lstArray(i) + "�"
                            ' and tell the connecting client to relogin (client does it automatically)
                            Socket2(index).SendLen = Len("LLLL2�")
                            Socket2(index).SendData = "LLLL2�"
                            
                            ' exit this message processing
                            ' we use this goto to skip the broadcast at the bottom of the select
                            GoTo stopper
                            
                            
                        End If
                            
                    
                    Next i
                    
                    'if not already logged in, then tell the client to log in
                    
                    Socket2(index).SendLen = Len("LLLL1�")
                    Socket2(index).SendData = "LLLL1�"
                    
                    
                        ' tell the client about all the connected clients
                        For i = 0 To lstArrayCount - 1
                            
                            Socket2(index).SendLen = Len("####" + lstArray(i) + "�" + Socket2(lstArraySocketIndex(i)).PeerAddress + "�")
                            Socket2(index).SendData = "####" + lstArray(i) + "�" + Socket2(lstArraySocketIndex(i)).PeerAddress + "�"
                            'GoSleepEX 0.1
                        Next i
                        
                        ' add the user to the local lists and arrays
                        lstPlayers.AddItem parsed
                        lstArray(lstArrayCount) = parsed
                        lstArraySocketIndex(lstArrayCount) = index
                        lstArrayCount = lstArrayCount + 1
                        
                        'GoSleepEX 1
                        
                        ' tell connecting client about avalible game rooms
                        For i = 0 To lstGameArrayCount - 1
                        
                            Socket2(index).SendLen = Len("++++" + lstGameArrayPlayers(i) + " " + lstGameArrayGame(i) + " " + lstGameArrayNames(i) + lstGameArrayIP(i) + "�")
                            Socket2(index).SendData = "++++" + lstGameArrayPlayers(i) + " " + lstGameArrayGame(i) + " " + lstGameArrayNames(i) + lstGameArrayIP(i) + "�"
                            
                            Socket2(index).SendLen = Len("GRPL" + lstGameArrayNames(i) + "�" + lstGameArrayGRPL(i) + "�")
                            Socket2(index).SendData = "GRPL" + lstGameArrayNames(i) + "�" + lstGameArrayGRPL(i) + "�"
                            'GoSleepEX 0.1
                        Next i
                    Else ' bad pwd or uid
                        Socket2(index).SendLen = Len("LLLL0�")
                        Socket2(index).SendData = "LLLL0�"
                        ' we use this goto to skip the broadcast at the bottom of the select
                        GoTo stopper
                    End If
                
        Case "@@@@": ' remove player
            ' get the player to be removed
            parsed = Mid(strdata, 5)
            'reset the counter
            cnt = 0
            'clear the players list
            lstPlayers.Clear
            'redimention the arrays to temporarly hold the players
            ReDim lstarraytemp(lstArrayCount)
            ReDim lstarraysocketindextemp(lstArrayCount)
            
            'loop through the arrays of players
            For i = 0 To lstArrayCount - 1
                'if the player is not the one to be removed
                If parsed <> Mid(lstArray(i), 4) Then
                    'copy the player into the temp array and re add it to the list
                    lstPlayers.AddItem lstArray(i)
                    lstarraytemp(cnt) = lstArray(i)
                    lstarraysocketindextemp(cnt) = lstArraySocketIndex(i)
                    'increment the counter
                    cnt = cnt + 1
                End If
                
            
            Next i
            
            'reset the array count, to the right one
            lstArrayCount = cnt
            'loop through the temp arrays and transfer back into the permanant ones
            For i = 0 To lstArrayCount - 1
                lstArray(i) = lstarraytemp(i)
                lstArraySocketIndex(i) = lstarraysocketindextemp(i)
            Next i
            
        Case "RRRR": ' register player
            ' recieved request to register new user
            
            'get uid and password
            parsed = Mid(strdata, 5)
            'get the username
            parsed = Mid(parsed, 1, InStr(parsed, "�") - 1)
            
            'set flag to go
            goflag = True
            ' go through the list of pwd and UID
            For i = 0 To 9999
                ' check to see if the current uid is the same as the one requested (already exists)
                If InStr(UCase(uidpwdArray(i)), UCase(parsed)) > 0 Then
                    goflag = False ' if so, set the flag to false
                End If
            
            Next i
            ' if the flag is still true (does not already exists)
            If goflag = True Then
            
                'append the uid�pwd to the file
                Open "c:\pwdfile.dat" For Append As #1
                
                Print #1, Mid(strdata, 5)
                
                Close #1
                
                
                'reload the list
                Open "c:\pwdfile.dat" For Input As #1
            
                For i = 0 To 10000
                    
                    If EOF(1) = True Then Exit For
                    Line Input #1, uidpwdArray(i)
                    i = i + 1
                
                Next i
                                
                Close #1
                
                'tell the client that the user has been registerd
                Socket2(index).SendLen = Len("RUSR1�")
                Socket2(index).SendData = "RUSR1�"
                parsed = "RUSR1" + parsed
                ' we use this goto to skip the broadcast at the bottom of the select
                GoTo stopper
            Else
                ' tell the client the user has not been registered
                Socket2(index).SendLen = Len("RUSR0�")
                Socket2(index).SendData = "RUSR0�"
                parsed = "RUSR0" + parsed
                ' we use this goto to skip the broadcast at the bottom of the select
                GoTo stopper
            End If
                                
        
        Case "++++": 'add gameroom
            'recieved message to create listing for new gameroom
            parsed = Mid(strdata, 5)
            ' fill the players
            lstGameArrayPlayers(lstGameArrayCount) = Mid(parsed, 1, 3)
            ' fill the game type
            lstGameArrayGame(lstGameArrayCount) = Mid(parsed, 5, 3)
            ' fill the game name
            lstGameArrayNames(lstGameArrayCount) = Mid(parsed, 9, 30)
            ' fill the game IP
            lstGameArrayIP(lstGameArrayCount) = Mid(parsed, 39)
            ' fill the game socket... socket of the host(used in detecting games hosted by
            ' players no longer here
            lstGameArraySocketIndex(lstGameArrayCount) = index
            'set the last communication time for the game to now
            lstGameArrayKeepAlive(lstGameArrayCount) = Now()
            
            ' increment the number of games
            lstGameArrayCount = lstGameArrayCount + 1
    
    
            'add the new game to the list of games
            lstgamerooms.AddItem (Mid(parsed, 1, 38))
                
            
            
        Case "%%%%": ' recieved signal to remove gameroom
            ' declare temp arrays
            Dim lstGameArrayIPtemp() As String
            Dim lstGameArrayPlayerstemp() As String
            Dim lstGameArrayNamestemp() As String
            Dim lstGameArrayGametemp() As String
            Dim lstGameArrayKeepAlivetemp() As String
            Dim lstGameArrayGRPLtemp() As String
        
            ' get the gameroom name to remove
            parsed = Mid(strdata, 5)
            ' set the counter to nothing
            cnt = 0
            'clear the list of gamerooms
            lstgamerooms.Clear
            
            'set the size of the temp arrays
            ReDim lstGameArrayIPtemp(lstGameArrayCount)
            ReDim lstGameArrayPlayerstemp(lstGameArrayCount)
            ReDim lstGameArrayNamestemp(lstGameArrayCount)
            ReDim lstGameArrayGametemp(lstGameArrayCount)
            ReDim lstGameArrayKeepAlivetemp(lstGameArrayCount)
            ReDim lstGameArrayGRPLtemp(lstGameArrayCount)
            
            ' loop through the list of games
            For i = 0 To lstGameArrayCount - 1
                'if the current game does not equal the one to remove then
                If parsed <> lstGameArrayNames(i) Then
                    're add it to the list
                    lstgamerooms.AddItem (lstGameArrayPlayers(i) + " " + lstGameArrayGame(i) + " " + Mid(lstGameArrayNames(i), 1, 30))
                    'copy the current game data to the temp array
                    lstGameArrayIPtemp(cnt) = lstGameArrayIP(i)
                    lstGameArrayPlayerstemp(cnt) = lstGameArrayPlayers(i)
                    lstGameArrayNamestemp(cnt) = lstGameArrayNames(i)
                    lstGameArrayGametemp(cnt) = lstGameArrayGame(i)
                    lstGameArrayKeepAlivetemp(cnt) = lstGameArrayKeepAlive(i)
                    lstGameArrayGRPLtemp(cnt) = lstGameArrayGRPL(i)
                    'increment the counter
                    cnt = cnt + 1
                End If
                
            
            Next i
            
            'reset the num of gamerooms
            lstGameArrayCount = cnt
            ' loop through the temp arrays, and transfer to the regular storage arrays
            For i = 0 To lstGameArrayCount - 1
                    lstGameArrayIP(i) = lstGameArrayIPtemp(i)
                    lstGameArrayPlayers(i) = lstGameArrayPlayerstemp(i)
                    lstGameArrayNames(i) = lstGameArrayNamestemp(i)
                    lstGameArrayKeepAlive(cnt) = lstGameArrayKeepAlivetemp(i)
                    lstGameArrayGame(i) = lstGameArrayGametemp(i)
                    lstGameArrayGRPL(i) = lstGameArrayGRPL(i)
            Next i
            
        
        
        
        Case "^^^^": ' recieved message change player status (the letter in brackets in the players list)
        
        
            'get the new status and name
            parsed = Mid(strdata, 5)
            'get the new status
            first3 = Mid(parsed, 1, 3)
            'get the name
            parsed = Mid(parsed, 4)
                'clear the list of players
                lstPlayers.Clear
                'loop through the players
                For i = 0 To lstArrayCount - 1
                    ' if the name is equal to the name sent then
                    If parsed = Mid(lstArray(i), 4) Then
                        'change it to the new status and name
                        lstArray(i) = first3 + parsed
                        
                    End If
                    'add item to list
                    lstPlayers.AddItem lstArray(i)
                Next i
            
        Case "IMIM": ' recieved a message to send a page to a player
            ' extract who the message is for
            uidIM = Mid(strdata, InStr(strdata, "�") + 1, InStr(InStr(strdata, "�") + 1, strdata, "�") - InStr(strdata, "�") - 1)
            'loop through the players
            For i = 0 To lstArrayCount - 1
                'if the current player is the one the message is for, send the message to them
                If uidIM = Mid(lstArray(i), 4) Then
                    Socket2(lstArraySocketIndex(i)).SendLen = Len(strdata + "�")
                    Socket2(lstArraySocketIndex(i)).SendData = strdata + "�"
                End If
            
            Next i
            
            'parsed = strdata
            'skip broadcast
            GoTo stopper
                
       
        Case "^++^":
            'recieve message of game room status change
            first3 = Mid(strdata, 5, 3) 'get players
            Game = Mid(strdata, 9, 3) 'get game type
            parsed = Mid(strdata, 13) 'get game name
            
            'clear our the list
            lstgamerooms.Clear
            'loop through the list of games
            For i = 0 To lstGameArrayCount - 1
                'if the current game name is equal to the game to update then
                If lstGameArrayNames(i) = parsed Then
                    'update the players and game type
                    lstGameArrayPlayers(i) = first3
                    lstGameArrayGame(i) = Game
                End If
                'add the current game to the list
                lstgamerooms.AddItem lstGameArrayPlayers(i) + " " + lstGameArrayGame(i) + " " + lstGameArrayNames(i)
            Next i
     
        Case "GRPL":
            'recieved message to update the Game Room Players List
            
            ' get the list
            parsed = Mid(strdata, 5)
            ' get the game name
            GRPLGamename = Mid(parsed, 1, InStr(parsed, "�") - 1)
            ' get the players
            GRPLstring = Mid(parsed, InStr(parsed, "�") + 1)
            ' loop through the games
            For i = 0 To lstGameArrayCount - 1
                ' if the game name to be updated equals the current one, then update it
                If lstGameArrayNames(i) = GRPLGamename Then
                    lstGameArrayGRPL(i) = GRPLstring
                End If
            Next i
      
      
        Case "KALV":
            ' recieved message to update the keepalive time...
            'get game name who is still alive
            parsed = Mid(strdata, 5)
            'loop through the game list
            For i = 0 To lstGameArrayCount - 1
                'if the current game name equals the one to update the keepalive time
               If lstGameArrayNames(i) = parsed Then
                    'then set it to the current time
                    lstGameArrayKeepAlive(i) = Now()
               End If
            
            Next i
            'do not broadcast message
            parsed = strdata
            GoTo stopper
      
        End Select
            
        'start code to broadcast messages
        
        'loop through connections
        For i = 1 To LastSocket
            On Error Resume Next
            ' if the connection is connected
            If Socket2(i).Connected Then
                'send data
                Socket2(i).SendLen = Len(strdata + "�")
                Socket2(i).SendData = strdata + "�"
                        
            End If
        Next i

        'set the messaging string
        parsed = strdata
            
stopper:
                
        ' wrap the message length
        For w = 0 To Int(Len(parsed) / 90)
                
            Wraplines(w) = Mid(parsed, (w * 90) + 1, 90)
                
        Next w
                            
        For w = Int(Len(parsed) / 90) To 0 Step -1
            'insert into listbox
            newMessage Wraplines(w)
        Next w
                
    Next B
    Exit Sub
ender:
    Open "c:\EVerrorlog " + Str(Day(Now())) + " -" + Str(Month(Now())) + " -" + Str(Year(Now())) + ".log" For Append As #1
    Print #1, "++++socket2_read" + " - " + Str(Now())
    Print #1, "++++" + Str(Err.Number) + " - " + Err.Description
    Close #1

    
End Sub

Private Function login(uidpwd As String) As Boolean
Dim i As Integer
On Error GoTo ender
    ' this function is to test if client logging is providing valid uid and pwd
    
    For i = 0 To 9999
        If UCase(uidpwd) = UCase(uidpwdArray(i)) Then
            login = True
            Exit Function
        Else
            login = False
            
        End If
    Next i
    

    Exit Function
ender:
    Open "c:\EVerrorlog " + Str(Day(Now())) + " -" + Str(Month(Now())) + " -" + Str(Year(Now())) + ".log" For Append As #1
    Print #1, "++++Login" + " - " + Str(Now())
    Print #1, "++++" + Str(Err.Number) + " - " + Err.Description
    Close #1

End Function

Private Sub newMessage(msg As String)
On Error GoTo ender
    'this sub is to insert a new message into the list box
    
    lstReply.Clear
    
    For i = 99 To 0 Step -1
    
       lstReplyarray(i + 1) = lstReplyarray(i)
    
    Next i
    
    lstReplyarray(0) = msg
    For i = 0 To 100
       lstReply.AddItem lstReplyarray(i)
    Next i

    Exit Sub
ender:
    Open "c:\EVerrorlog " + Str(Day(Now())) + " -" + Str(Month(Now())) + " -" + Str(Year(Now())) + ".log" For Append As #1
    Print #1, "++++newMessage" + " - " + Str(Now())
    Print #1, "++++" + Str(Err.Number) + " - " + Err.Description
    Close #1

End Sub
Private Sub Socket2_Disconnect(index As Integer)

'this sub is to handle disconnection of sockets
On Error GoTo ender
    Socket2(index).Action = SOCKET_CLOSE
    
    Exit Sub
ender:
    Open "c:\EVerrorlog " + Str(Day(Now())) + " -" + Str(Month(Now())) + " -" + Str(Year(Now())) + ".log" For Append As #1
    Print #1, "++++Disconnect" + " - " + Str(Now())
    Print #1, "++++" + Str(Err.Number) + " - " + Err.Description
    Close #1

End Sub


'Private Sub txtSend_KeyPress(KeyAscii As Integer)
'On Error GoTo ender
'   If KeyAscii = 13 Then
'        Socket1.SendLen = 1024 'Len(txtSend.Text)
'        Socket1.SendData = txtSend.Text
'        KeyAscii = 0: txtSend.Text = ""
'    End If
'    Exit Sub
'ender:
'    Open "c:\EVerrorlog " + Str(Day(Now())) + " -" + Str(Month(Now())) + " -" + Str(Year(Now())) + ".log" For Append As #1
'    Print #1, "++++txtSend_KeyPress" + " - " + Str(Now())
'    Print #1, "++++" + Str(Err.Number) + " - " + Err.Description
'    Close #1

'End Sub

' this is to deal with server being closed
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ender
    ' loop through sockets and disconnect listening ports
    Dim i As Integer
    If Socket2(0).Listening Then Socket2(0).Action = SOCKET_CLOSE
    For i = 1 To LastSocket
        If Socket2(i).Connected Then Socket2(i).Action = SOCKET_CLOSE
    Next i
    End
    Exit Sub
ender:
    Open "c:\EVerrorlog " + Str(Day(Now())) + " -" + Str(Month(Now())) + " -" + Str(Year(Now())) + ".log" For Append As #1
    Print #1, "++++Form_Unload" + " - " + Str(Now())
    Print #1, "++++" + Str(Err.Number) + " - " + Err.Description
    Close #1

End Sub

Private Sub Timercheckconnections_Timer()

' this sub checks for crashed clients and removes them from the list
On Error GoTo ender
    Dim lstarraytemp() As String
    Dim lstarraysocketindextemp() As Integer
    Dim RemoveGame As String
    
    Dim lstGameArrayIPtemp() As String
    Dim lstGameArrayPlayerstemp() As String
    Dim lstGameArrayNamestemp() As String
    Dim lstGameArrayGametemp() As String
    Dim lstGameArrayKeepAlivetemp() As String
    Dim lstGameArrayGRPLtemp() As String
    
    
    For j = 0 To lstArrayCount - 1
        
    ' check for missing player
        If Not Socket2(lstArraySocketIndex(j)).Connected Then
                    
            ' if player is misssing
                        ' if not connected... see if they are hosting a game, and remove that too
                
            For k = 0 To lstGameArrayCount - 1
                
                If lstGameArraySocketIndex(k) = lstArraySocketIndex(j) Then
                    
                                           
                    parsed = lstGameArrayNames(k)
                    
                    cnt = 0
                            
                    lstgamerooms.Clear
                    
                    ReDim lstGameArrayIPtemp(lstGameArrayCount)
                    ReDim lstGameArrayPlayerstemp(lstGameArrayCount)
                    ReDim lstGameArrayNamestemp(lstGameArrayCount)
                    ReDim lstGameArrayGametemp(lstGameArrayCount)
                    ReDim lstGameArrayKeepAlivetemp(lstGameArrayCount)
                    ReDim lstGameArrayGRPLtemp(lstGameArrayCount)
                    
                    
                    For i = 0 To lstGameArrayCount - 1
                    
                        If parsed <> lstGameArrayNames(i) Then
                            
                            lstgamerooms.AddItem (Mid(lstGameArrayPlayers(i) + " " + lstGameArrayNames(i), 1, 26))
                            
                            lstGameArrayIPtemp(cnt) = lstGameArrayIP(i)
                            lstGameArrayPlayerstemp(cnt) = lstGameArrayPlayers(i)
                            lstGameArrayNamestemp(cnt) = lstGameArrayNames(i)
                            lstGameArrayGametemp(cnt) = lstGameArrayGame(i)
                            lstGameArrayKeepAlivetemp(cnt) = lstGameArrayKeepAlive(i)
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
                            lstGameArrayKeepAlive(i) = lstGameArrayKeepAlivetemp(i)
                            lstGameArrayGRPL(i) = lstGameArrayGRPLtemp(i)
                    Next i
                    
                    For i = 1 To LastSocket
                    On Error Resume Next
                    If Socket2(i).Connected Then
                        Socket2(i).SendLen = Len("%%%%" + parsed + "�")
                        Socket2(i).SendData = "%%%%" + parsed + "�"
                        
                    End If
                Next i

                    
                End If
            
            Next k
        
        ' then remove player
                
            parsed = Mid(lstArray(j), 4)
            
          '  For g = 0 To lstArrayCount - 1
            
            
         '   Next g
            
            cnt = 0
                    
            lstPlayers.Clear
            
            ReDim lstarraytemp(lstArrayCount)
            ReDim lstarraysocketindextemp(lstArrayCount)
            
            For i = 0 To lstArrayCount - 1
            
                If parsed <> Mid(lstArray(i), 4) Then
                    lstPlayers.AddItem lstArray(i)
                    lstarraytemp(cnt) = lstArray(i)
                    lstarraysocketindextemp(cnt) = lstArraySocketIndex(i)
                    cnt = cnt + 1
                    
                    If Socket2(lstArraySocketIndex(i)).Connected Then
                    
                        Socket2(lstArraySocketIndex(i)).SendLen = Len("@@@@" + parsed + "�")
                        Socket2(lstArraySocketIndex(i)).SendData = "@@@@" + parsed + "�"
                        
                        
                        
                    End If
                    
                    
                End If
                
            
            Next i
            
            
                
                    
                    
            lstArrayCount = cnt
            For i = 0 To lstArrayCount - 1
                lstArray(i) = lstarraytemp(i)
                lstArraySocketIndex(i) = lstarraysocketindextemp(i)
            Next i
                    
    
        End If
    
    Next j
    
    
    
    Exit Sub
ender:
    Open "c:\EVerrorlog " + Str(Day(Now())) + " -" + Str(Month(Now())) + " -" + Str(Year(Now())) + ".log" For Append As #1
    Print #1, "++++TimerCheckConnections" + " - " + Str(Now())
    Print #1, "++++" + Str(Err.Number) + " - " + Err.Description
    Close #1

End Sub

Private Sub timercheckgamerooms_Timer()

    'this sub is to remove gamerooms when their hosts have disapeared

    Dim lstGameArrayIPtemp() As String
    Dim lstGameArrayPlayerstemp() As String
    Dim lstGameArrayNamestemp() As String
    Dim lstGameArrayGametemp() As String
    Dim lstGameArrayKeepAlivetemp() As String
    Dim lstGameArrayGRPLtemp() As String
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    Dim gameRoomGhost As Boolean
    
    'loop through number of gamerooms
    For k = 0 To lstGameArrayCount - 1
    
        'set the gameroom flag to yes - it's a ghost
'        gameRoomGhost = True
        
        ' loop through the playernames
        'For j = 0 To lstArrayCount - 1
            ' if the player hosting the game is the same as the current player
         '   If Mid(lstGameArrayGRPL(k), 7, Len(Mid(lstArray(j), 4))) = Mid(lstArray(j), 4) Then
                ' the host exists, so the gameroom is not a ghost
          '      gameRoomGhost = False
           '     Exit For
           ' End If
        'Next j
            ' if gameroom is a ghost
            
            If DateDiff("s", lstGameArrayKeepAlive(k), Now()) > KeepAliveTimeout Then
                ' send out the kill signal for the gameroom
                strdata = "%%%%" + lstGameArrayNames(k)
            
                ' loop through all sockets and send to them
                For i = 1 To LastSocket
                    On Error Resume Next
                    If Socket2(i).Connected Then
                        Socket2(i).SendLen = Len(strdata + "�")
                        Socket2(i).SendData = strdata + "�"
                        
                    End If
                Next i

                ' remove the gameroom from the array
                parsed = lstGameArrayNames(k)
                cnt = 0
                lstgamerooms.Clear
                
                ReDim lstGameArrayIPtemp(lstGameArrayCount)
                ReDim lstGameArrayPlayerstemp(lstGameArrayCount)
                ReDim lstGameArrayNamestemp(lstGameArrayCount)
                ReDim lstGameArrayGametemp(lstGameArrayCount)
                ReDim lstGameArrayKeepAlivetemp(lstGameArrayCount)
                ReDim lstGameArrayGRPLtemp(lstGameArrayCount)
                
                'transfer all the valid gamerooms into a set of temp arrays
                For i = 0 To lstGameArrayCount - 1
                    
                    If parsed <> lstGameArrayNames(i) Then
                            
                        lstgamerooms.AddItem (lstGameArrayPlayers(i) + " " + lstGameArrayGame(i) + " " + Mid(lstGameArrayNames(i), 1, 30))
                            
                        lstGameArrayIPtemp(cnt) = lstGameArrayIP(i)
                        lstGameArrayPlayerstemp(cnt) = lstGameArrayPlayers(i)
                        lstGameArrayNamestemp(cnt) = lstGameArrayNames(i)
                        lstGameArrayGametemp(cnt) = lstGameArrayGame(i)
                        lstGameArrayKeepAlivetemp(cnt) = lstGameArrayKeepAlive(i)
                        lstGameArrayGRPLtemp(cnt) = lstGameArrayGRPL(i)
                        cnt = cnt + 1
                    End If
                Next i
                    
                ' adjust the count
                lstGameArrayCount = cnt
                'transfer all the temp gamerooms back to the original arrays
                For i = 0 To lstGameArrayCount - 1
                        lstGameArrayIP(i) = lstGameArrayIPtemp(i)
                        lstGameArrayPlayers(i) = lstGameArrayPlayerstemp(i)
                        lstGameArrayNames(i) = lstGameArrayNamestemp(i)
                        lstGameArrayGame(i) = lstGameArrayGametemp(i)
                        lstGameArrayKeepAlive(i) = lstGameArrayKeepAlivetemp(i)
                        lstGameArrayGRPL(i) = lstGameArrayGRPLtemp(i)
                Next i
           End If
    Next k
End Sub

Private Sub VScroll1_Change()
    txtGameRoomTimeout.Text = VScroll1.Value
    KeepAliveTimeout = txtGameRoomTimeout.Text
    
End Sub
