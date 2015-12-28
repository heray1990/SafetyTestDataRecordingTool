Attribute VB_Name = "Module4"
Option Explicit

'The end code of command for Chroma 19032 is vbLF or vbCrLf.
'About the detail of command for 19032, please see its user manual.

Dim cmdBuf As String

Public Sub SAFE_STAR()
    cmdBuf = "SAFE:STAR"
    cmdIdentifyNum = 1
    
    Form1.MSComm1.Output = cmdBuf & vbCrLf
    Log_Info cmdBuf
End Sub

Public Sub SAFE_STOP()
    cmdBuf = "SAFE:STOP"
    cmdIdentifyNum = 2
    
    Form1.MSComm1.Output = cmdBuf & vbCrLf
    Log_Info cmdBuf
End Sub

Public Sub SAFE_RES_AREP(flag As String)
    cmdBuf = "SAFE:RES:AREP" & Space(1) & flag
    cmdIdentifyNum = 3
    
    Form1.MSComm1.Output = cmdBuf & vbCrLf
    Log_Info cmdBuf
End Sub

Public Sub SAFE_RES_AREP_ITEM(items As String)
    cmdBuf = "SAFE:RES:AREP:ITEM" & Space(1) & items
    cmdIdentifyNum = 4
    
    Form1.MSComm1.Output = cmdBuf & vbCrLf
    Log_Info cmdBuf
End Sub

Public Sub ASK_STEP_SNUM()
    cmdBuf = "SAFE:SNUM?"
    cmdIdentifyNum = 5
    
    Form1.MSComm1.Output = cmdBuf & vbCrLf
    Log_Info cmdBuf
End Sub

Public Sub ASK_ALL_STEP_NAME()
    cmdBuf = "SAFE:RES:ALL:MODE?"
    cmdIdentifyNum = 6
    
    Form1.MSComm1.Output = cmdBuf & vbCrLf
    Log_Info cmdBuf
End Sub

Public Sub ASK_ALL_STEP_SPEC(bufStr As String)
    cmdBuf = bufStr
    cmdIdentifyNum = 7
    
    Form1.MSComm1.Output = cmdBuf & vbCrLf
    Log_Info cmdBuf
End Sub
