VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Open "c:\EVerrorlog  11 - 12 - 2004.log" For Input As #1
Open "c:\today.log" For Output As #2

Do Until EOF(1)

    Line Input #1, temp

    If InStr(temp, "MMMM") <> 0 Then
        Print #2, temp
    End If
Loop

    
Close #2
Close #1
End Sub
