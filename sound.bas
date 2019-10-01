Attribute VB_Name = "sound"
'Written for form or class module, change declarations to public
'for .bas modules

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Const SND_ASYNC = &H1 'continue executing code even
'if sound isn't finished
Const SND_FILENAME = &H20000 '  name is a file name
Const SND_SYNC = &H0 'suspend execution until sound is finished
Const SND_NODEFAULT = &H2 'if file name is not found, don't play
'default sound
Const SND_LOOP = &H8 'loop the sound until next call to the
'function
Const SND_NOSTOP = &H10   'don't stop any currently playing sound
Const SND_NOWAIT = &H2000  'return immediately if driver is busy

'PURPOSE: Plays a .WAV file
'PARAMETER: .WAV to play
'RETURNS: True if Successful, false otherwise
Public Function PlayASound(SoundFile As String) As Boolean

Dim bSuccess As Boolean

'Flags indicate that sound is a file, to play asynchrounously,
'not to stop any currently playing sound, and not to play the
'default sound (and return false) if file
'is not found.  See declarations for other options.

bSuccess = PlaySound(SoundFile, vbNull, SND_FILENAME _
+ SND_SYNC + SND_NOSTOP + SND_NODEFAULT)

PlayASound = bSuccess

End Function

