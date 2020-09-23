Attribute VB_Name = "Module1"
' Pause Module.
Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Pause(HowLong As Long) ' HowLong interval i ms.
 Dim lngEnd As Long
 
lngEnd = GetTickCount() + HowLong
Do

DoEvents

Loop Until GetTickCount() >= lngEnd
End Sub
