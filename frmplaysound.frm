VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmplaysound 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin MCI.MMControl MMControl1 
      Height          =   2055
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   3625
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   "C:\Documents and Settings\Administrator\Desktop\xvt\bing.WAV"
   End
End
Attribute VB_Name = "frmplaysound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
   MMControl1.Command = "Prev"
    MMControl1.Command = "Play"
End Sub

Private Sub Form_Load()

    Dim temp As String
    temp = Command
    If temp = "" Then temp = "bing.wav"
    
    ' Set properties needed by MCI to open.
    MMControl1.Notify = False
    MMControl1.Wait = False
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.FileName = App.Path & "\" & temp
    
    ' Open the MCI WaveAudio device.
    MMControl1.Command = "Open"
    MMControl1.Command = "Play"

    'Shell (App.Path + "\" + temp)
    'PlayASound (App.Path + "\" + Command)
    'Unload Me
End Sub
