VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form frmConfigureGame 
   Caption         =   "Configure Game"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   Icon            =   "frmConfigureGame.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   3000
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   800
      ScreenWidth     =   1280
      ScreenHeightDT  =   800
      ScreenWidthDT   =   1280
      FormHeightDT    =   3420
      FormWidthDT     =   4770
      FormScaleHeightDT=   2910
      FormScaleWidthDT=   4650
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   255
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Game Title"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   255
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Number of Players"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Game To Play"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2895
      Begin VB.OptionButton optBOP 
         BackColor       =   &H00000000&
         Caption         =   "Balance Of Power"
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
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optXWA 
         BackColor       =   &H00000000&
         Caption         =   "X-Wing Alliance"
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2655
      End
      Begin VB.OptionButton optXVT 
         BackColor       =   &H00000000&
         Caption         =   "X-Wing vs Tie"
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
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.TextBox txtGameName 
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
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   120
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin VB.VScrollBar scrlPlayers 
      Height          =   375
      Left            =   480
      Max             =   2
      Min             =   8
      TabIndex        =   2
      Top             =   360
      Value           =   4
      Width           =   255
   End
   Begin VB.TextBox txtPlayers 
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   360
      Width           =   375
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5175
      ExtentX         =   9128
      ExtentY         =   5953
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
   Begin VB.Label Label2 
      Caption         =   "Game Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Number Of Players"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmConfigureGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Filter() As String


Private Sub btnCancel_Click()
    On Error GoTo Failed
    Unload Me
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnCancel_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmconfiguregame"
    Close #4
    Resume Next

End Sub

Private Sub Form_Load()
    On Error GoTo Failed
    
    Filter = frmClient.getFilter()
    
    
    frmClient.numberOfPlayers = 4
    
    WebBrowser1.Navigate2 CurDir + "\configure.htm"
    
    
    If frmClient.XWAGamePath = "" Then
        optXWA.Visible = False
    Else
        frmClient.game = "XWA"
        optXWA.Value = True
    End If
    
    If frmClient.BOPGamePath = "" Then
        optBop.Visible = False
    Else
        frmClient.game = "BOP"
        optBop.Value = True
    End If
    
    If frmClient.XVTGamePath = "" Then
        optXVT.Visible = False
    Else
        frmClient.game = "XVT"
        optXVT.Value = True
    End If
   
    If frmClient.game = "" Then
    
        MsgBox "You can't host a game because you do not have any supported games installed", vbCritical + vbOKOnly, "No Games installed"
        Unload Me
    End If
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "Form_Load - " + Str(Err.Number) + " - " + Err.Description + " - frmconfiguregame"
    Close #4
    Resume Next
    
End Sub

Private Sub optBOP_Click()
    On Error GoTo Failed
    frmClient.game = "BOP"
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "optBOP_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmconfiguregame"
    Close #4
    Resume Next

End Sub

Private Sub optXVT_Click()
    On Error GoTo Failed
    frmClient.game = "XVT"
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "optXVT_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmconfiguregame"
    Close #4
    Resume Next

End Sub

Private Sub optXWA_Click()
    On Error GoTo Failed
    frmClient.game = "XWA"
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "optXWA_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmconfiguregame"
    Close #4
    Resume Next


End Sub

Private Sub scrlPlayers_Change()
    On Error GoTo Failed
    txtPlayers.Text = scrlPlayers.Value
            Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "scrlPlayers_Change - " + Str(Err.Number) + " - " + Err.Description + " - frmconfiguregame"
    Close #4
    Resume Next


End Sub
Private Sub btnOK_Click()
    On Error GoTo Failed
    Dim ip As String
    ip = frmClient.exposedIP
    Dim exists As Boolean
    exists = False
    


   If txtGameName.Text <> "" Then
        
        For i = 1 To 30 - Len(txtGameName.Text)
            txtGameName.Text = txtGameName.Text + " "
        Next i
        
        
        
        If frmClient.does_exist(txtGameName.Text) = False Then
        
            frmClient.gameName = txtGameName.Text
            
            
            frmClient.btnHost.Enabled = False
            frmClient.btnJoin.Enabled = False
            frmClient.btnAway.Enabled = False
        
            frmClient.lstGameRooms.Enabled = False
            
            frmClient.Socket1.SendLen = Len("++++" + "1/" + txtPlayers.Text + " " + frmClient.game + " " + frmClient.gameName + ip + Chr$(3))
            frmClient.Socket1.SendData = "++++" + "1/" + txtPlayers.Text + " " + frmClient.game + " " + frmClient.gameName + ip + Chr$(3)
    
            frmClient.Socket1.SendLen = Len("GRPL" + frmClient.gameName + Chr$(30) + "Host: " + frmClient.txtuid.Text + Chr$(3))
            frmClient.Socket1.SendData = "GRPL" + frmClient.gameName + Chr$(30) + "Host: " + frmClient.txtuid.Text + Chr$(3)
           
            frmHostGame.Show
        
            Unload Me
        Else
            MsgBox "That Game Name already exists", vbOKOnly, "Validation Error"
        End If
       
   Else
        MsgBox "Enter a Game Name", vbOKOnly, "Validation Error"
   End If
        Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnOK_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmconfiguregame"
    Close #4
    Resume Next


End Sub
    
    
    
      

Private Sub txtGameName_Change()
    On Error GoTo Failed
    For i = 0 To frmClient.FilterCount

        If frmClient.cboFilterOff.Value <> 1 Then
        
            txtGameName.Text = Replace(txtGameName.Text, Filter(i), "Smurf", , , vbTextCompare)
        
        End If

    Next i
        Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "txtGameNameChange - " + Str(Err.Number) + " - " + Err.Description + " - frmconfiguregame"
    Close #4
    Resume Next

End Sub

Private Sub txtPlayers_Change()
    On Error GoTo Failed
    frmClient.numberOfPlayers = txtPlayers.Text
    
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "txtPlayers_Change - " + Str(Err.Number) + " - " + Err.Description + " - frmconfiguregame"
    Close #4
    Resume Next

End Sub

