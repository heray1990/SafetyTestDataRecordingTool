Attribute VB_Name = "Module4"
Option Explicit

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
