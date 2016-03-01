Attribute VB_Name = "CmdHelper"
Option Explicit

'The end code of command for Chroma 19032 is vbLF or vbCrLf.
'About the detail of command for 19032, please see its user manual.

Dim mSendDataBuf As String

Private Sub SendCmd()
    Form1.MSComm1.Output = mSendDataBuf & vbCrLf
    DelayMS 500
    
    Log_Info mSendDataBuf
End Sub

Public Sub SAFE_STAR()
    mSendDataBuf = "SAFE:STAR"
    cmdIdentifyNum = 1
    
    SendCmd
End Sub

Public Sub SAFE_STOP()
    mSendDataBuf = "SAFE:STOP"
    cmdIdentifyNum = 2
    
    SendCmd
End Sub

Public Sub SAFE_RES_AREP(flag As String)
    mSendDataBuf = "SAFE:RES:AREP" & Space(1) & flag
    cmdIdentifyNum = 3
    
    SendCmd
End Sub

Public Sub SAFE_RES_AREP_ITEM(items As String)
    mSendDataBuf = "SAFE:RES:AREP:ITEM" & Space(1) & items
    cmdIdentifyNum = 4
    
    SendCmd
End Sub

Public Sub ASK_STEP_SNUM()
    mSendDataBuf = "SAFE:SNUM?"
    cmdIdentifyNum = 5
    
    SendCmd
End Sub

Public Sub ASK_ALL_STEP_NAME()
    mSendDataBuf = "SAFE:RES:ALL:MODE?"
    cmdIdentifyNum = 6
    
    SendCmd
End Sub

Public Sub ASK_ALL_STEP_SPEC(bufStr As String)
    mSendDataBuf = bufStr
    cmdIdentifyNum = 7
    
    SendCmd
End Sub
