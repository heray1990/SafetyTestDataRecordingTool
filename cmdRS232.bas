Attribute VB_Name = "Module4"
Option Explicit

'The end code of command for Chroma 19032 is vbLF or vbCrLf.
'About the detail of command for 19032, please see its user manual.

Dim cmdBuf As String

Public Sub SAFE_STAR()
    cmdBuf = "SAFE:STAR"
    Form1.MSComm1.Output = cmdBuf & vbCrLf
    Log_Info cmdBuf
End Sub

Public Sub SAFE_STOP()
    cmdBuf = "SAFE:STOP"
    Form1.MSComm1.Output = cmdBuf & vbCrLf
    Log_Info cmdBuf
End Sub

Public Sub SAFE_RES_AREP(flag As String)
    cmdBuf = "SAFE:RES:AREP" & Space(1) & flag
    Form1.MSComm1.Output = cmdBuf & vbCrLf
    Log_Info cmdBuf
End Sub

Public Sub ASK_SAFE_SNUM()
    cmdBuf = "SAFE:SNUM?"
    Form1.MSComm1.Output = cmdBuf & vbCrLf
    Log_Info cmdBuf
End Sub
