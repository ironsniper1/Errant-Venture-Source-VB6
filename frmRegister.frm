VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form frmRegister 
   Caption         =   "Register"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   1080
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   800
      ScreenWidth     =   1280
      ScreenHeightDT  =   800
      ScreenWidthDT   =   1280
      FormHeightDT    =   3705
      FormWidthDT     =   4800
      FormScaleHeightDT=   3195
      FormScaleWidthDT=   4680
   End
   Begin VB.TextBox txtpwd2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2160
      Width           =   4455
   End
   Begin VB.CommandButton btnRegister 
      Caption         =   "Register"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtpwd 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox txtuid 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm Password"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "UserID including Club (ie TRA_Stresser)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnRegister_Click()
    On Error GoTo Failed
    Dim reg As Boolean
    
    reg = True
    

    If txtpwd.Text <> txtpwd2.Text Then
        MsgBox "Passwords don't match, Try again", vbExclamation + vbOKOnly, "Validation Error"
        reg = False
    End If

    If Len(txtpwd.Text) < 7 Then
        MsgBox "Passwords must be 8 characters or longer", vbExclamation + vbOKOnly, "Validation Error"
        reg = False
    End If
    
    'If Mid(txtuid.Text, 3, 1) <> "_" And Mid(txtuid.Text, 4, 1) <> "_" And Mid(txtuid.Text, 5, 1) <> "_" Then
    '    MsgBox "include club info as above" + Chr$(13) + "If you are not a member of a club, enter Newb_ at the start of your userid", vbExclamation + vbOKOnly, "Validation Error"
    '    reg = False
    'End If

    If reg Then
    
        frmClient.Socket1.SendLen = Len("RRRR" + txtuid.Text + Chr$(30) + txtpwd.Text + Chr$(3))
        frmClient.Socket1.SendData = "RRRR" + txtuid.Text + Chr$(30) + txtpwd.Text + Chr$(3)
        
        Me.Enabled = False
        
    End If
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "btnRegister_Click - " + Str(Err.Number) + " - " + Err.Description + " - frmRegister"
    Close #4
    Resume Next


End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Failed
    
    Unload Me
    frmClient.Enabled = True
    
        Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "form_UnLoad - " + Str(Err.Number) + " - " + Err.Description + " - frmRegister"
    Close #4
    Resume Next

End Sub
