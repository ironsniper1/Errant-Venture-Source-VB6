VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form frmViewBanned 
   Caption         =   "Banned Users"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9795
   Icon            =   "frmViewBanned.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   3480
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   800
      ScreenWidth     =   1280
      ScreenHeightDT  =   800
      ScreenWidthDT   =   1280
      FormHeightDT    =   4185
      FormWidthDT     =   9915
      FormScaleHeightDT=   3675
      FormScaleWidthDT=   9795
   End
   Begin VB.CommandButton cmdunban 
      Caption         =   "User Has Turned From the Dark side"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   2895
   End
   Begin VB.ListBox lstBanned 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "frmViewBanned"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private selectedUser As String
Private BannedUser() As String

Private Sub cmdunBan_Click()
    On Error GoTo Failed
    If selectedUser <> "" Then
    
        frmClient.Socket1.SendLen = Len("RMBN" + selectedUser + Chr$(3))
        frmClient.Socket1.SendData = "RMBN" + selectedUser + Chr$(3)
    End If
    cmdunban.Enabled = False
     Unload frmViewBanned
        Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "cmdunBan_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmViewBanned"
    Close #4
    Resume Next

End Sub

Private Sub Form_Load()
    On Error GoTo Failed

    BannedUser = frmClient.getBannedUsers()

    For i = 0 To frmClient.BannedUserCounter
        lstBanned.AddItem BannedUser(i)
        Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "Form_Load - " + Str(Err.Number) + " - " + Err.Description + " - frmViewBanned"
    Close #4
    Resume Next
    Next i

End Sub

Private Sub lstBanned_Click()
    On Error GoTo Failed
    
    Dim selectedData As String
    
    selectedData = BannedUser(lstBanned.ListIndex)
    
    selectedUser = Mid(selectedData, 1, InStr(1, selectedData, Chr$(30)) - 1)
    
    cmdunban.Enabled = True
        Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "lstBanned_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmViewBanned"
    Close #4
    Resume Next
    
End Sub
