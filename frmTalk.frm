VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.Form frmTalk 
   Caption         =   "Text To Speech"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   2625
   StartUpPosition =   3  'Windows Default
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS DirectSS1 
      Height          =   1815
      Left            =   0
      OleObjectBlob   =   "frmTalk.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmTalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
