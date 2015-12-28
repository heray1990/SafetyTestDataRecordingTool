Attribute VB_Name = "Module3"
Option Explicit


Public strCurrentModelName As String
Public setTVCurrentComBaud As Long
Public strDataVersion As String
Public strSerialNo As String

Public setTVCurrentComID As Integer
Public setData As Integer
Public setDay As Integer
Public stepNum As Integer

Public cmdIdentifyNum As Integer

Public Sub Log_Clear()
    Form1.txtReceive.Text = ""
End Sub

Public Sub Log_Info(strLog As String)
    Form1.txtReceive.Text = Form1.txtReceive.Text & strLog & vbCrLf
End Sub
