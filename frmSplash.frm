VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8610
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   574
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmddeny 
      Caption         =   "I want to show I'm a rebel by flaunting the authority of an online gamming community, or in otherwords I'm a Smurfhead"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   8040
      Width           =   4095
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "I Agree to abide by the TOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   3975
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   4800
      Width           =   8535
      ExtentX         =   15055
      ExtentY         =   5530
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
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      ExtentX         =   15055
      ExtentY         =   8281
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
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit





Private Sub cmdAccept_Click()
    On Error GoTo Failed
    Unload Me
    frmClient.Show
        Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "timer1_timer - " + Str(Err.Number) + " - " + Err.Description + " - frmsplash"
    Close #4
    Resume Next

End Sub

Private Sub cmddeny_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo Failed
    
    WebBrowser1.Navigate2 CurDir + "\splash.htm"
    WebBrowser2.Navigate2 CurDir + "\TOS.htm"
    
    'Load frmClient
    Me.Caption = "Welcome to the NRSD Errant Venture - Version 10.93"
    
    Exit Sub
Failed:
    Open CurDir + "\errorlog.txt" For Append As #4
    Print #4, "Foram_Load - " + Str(Err.Number) + " - " + Err.Description + " - frmsplash"
    Close #4
    Resume Next
    
End Sub




