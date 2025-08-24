Attribute VB_Name = "Sonido"
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uflags As Long) As Long
' Public Type uflags
'SND_ALIAS As Long 'Play a Windows sound (such as SystemStart, Asterisk, etc.).
'SND_ASYNC As Long 'Continue program execution immediately after starting to play the sound.
'SND_FILENAME As Long 'Play the specified filename.
'SND_LOOP As Long   'Play the sound repeatedly until sndPlaySound is called again with lpszSoundName = "". SND_ASYNC must also be set.
'SND_NODEFAULT As Long 'Do not play the Windows default sound if the specified sound cannot be found.
'SND_NOSTOP As Long 'Do not stop playing any currently playing sounds.
'SND_NOWAIT As Long ' Do not wait if the sound driver is busy.
'SND_SYNC As Long ' Wait until the sound has finished playing before continuing program execution.
'End Type
Const SND_ALIAS = &H10000
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
Const SND_LOOP = &H8
Const SND_NODEFAULT = &H2
Const SND_NOSTOP = &H10
Const SND_NOWAIT = &H2000
Const SND_SYNC = &H0

