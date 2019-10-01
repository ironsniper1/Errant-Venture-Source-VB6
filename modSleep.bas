Attribute VB_Name = "modSleep"
' modSleep - Use API which allows other processes to continue
' 1998/05/07 Copyright 1998, Larry Rebich
    
    Option Explicit
    DefLng A-Z

    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    '

Public Sub GoSleep(lSeconds As Long)        'pass seconds as long [original function]
    Const clMillPerSec As Long = 1000       'milliseconds per second
    Sleep lSeconds * clMillPerSec           'convert seconds to milliseconds then call sleep
End Sub

Public Sub GoSleepEX(rSeconds As Single)    'pass seconds as single to allow decimal
    Const clMillPerSec As Long = 1000       'milliseconds per second
    Dim lSeconds As Long
    lSeconds = rSeconds * clMillPerSec      'convert to long
    Sleep lSeconds                          'call sleep
End Sub


