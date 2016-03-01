Attribute VB_Name = "DelayHelper"
Option Explicit

Public Sub Delay(mmSec As Long)
On Error GoTo ShowError
    Dim START As Single
    START = Timer
    While (Timer - START) < (mmSec / 1000#)
        DoEvents
    Wend
    Exit Sub
ShowError:
    MsgBox Err.Source & "------" & Err.Description
    Exit Sub
End Sub


Public Sub DelayMS(mmSec As Long)
On Error GoTo ShowError
    Dim START As Single
    
    START = Timer
    While (Timer - START) < (mmSec / 1000#)
        DoEvents
    Wend
    
    Exit Sub

ShowError:
    MsgBox Err.Source & "------" & Err.Description
    Exit Sub
End Sub

