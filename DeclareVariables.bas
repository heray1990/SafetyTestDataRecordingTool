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

Public Const SNColNum As Integer = 1
Public Const stepNoColNum As Integer = 2
Public Const GB_ItmColNum As Integer = 3
Public Const GB_RmColNum As Integer = 4
Public Const GB_LowColNum As Integer = 5
Public Const GB_HighColNum As Integer = 6
Public Const AC_VtmColNum As Integer = 7
Public Const AC_ImColNum As Integer = 8
Public Const AC_LowColNum As Integer = 9
Public Const AC_HighColNum As Integer = 10
Public Const DC_VtmColNum As Integer = 11
Public Const DC_ImColNum As Integer = 12
Public Const DC_LowColNum As Integer = 13
Public Const DC_HighColNum As Integer = 14
Public Const IR_VtmColNum As Integer = 15
Public Const IR_RmColNum As Integer = 16
Public Const IR_LowColNum As Integer = 17
Public Const IR_HighColNum As Integer = 18
Public Const LC_VtmColNum As Integer = 19
Public Const LC_ImColNum As Integer = 20
Public Const LC_LowColNum As Integer = 21
Public Const LC_HighColNum As Integer = 22
Public Const OSC_VtmColNum As Integer = 23
Public Const OSC_CColNum As Integer = 24
Public Const OSC_OpenColNum As Integer = 25
Public Const OSC_ShortColNum As Integer = 26
Public Const Judge_StepColNum As Integer = 27
Public Const Judge_TotalColNum As Integer = 28
Public Const dateAndTimeColNum As Integer = 29

Public Sub Log_Clear()
    Form1.txtReceive.Text = ""
End Sub

Public Sub Log_Info(strLog As String)
    Form1.txtReceive.Text = Form1.txtReceive.Text & strLog & vbCrLf
End Sub
