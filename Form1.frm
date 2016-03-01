VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "∞≤πÊ≤‚ ‘ ˝æ›±£¥Êπ§æﬂ"
   ClientHeight    =   7215
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   6480
      Top             =   360
   End
   Begin VB.Frame Frame2 
      Caption         =   "≤‚ ‘Ω·π˚"
      Height          =   1005
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Width           =   5295
      Begin VB.Label lbResult 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Checking"
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5025
      End
   End
   Begin VB.TextBox txtReceive 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4395
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1680
      Width           =   5265
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ãı¬Î"
      Height          =   820
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5295
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Œ¢»Ì—≈∫⁄"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   0
         Text            =   "123456789"
         Top             =   240
         Width           =   5025
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5640
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lbModelName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5295
   End
   Begin VB.Menu tbSetComPort 
      Caption         =   "…Ë÷√¥Æø⁄"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strCommInput As String
Dim lastRowNum As Integer
Dim resArray() As String
Dim stepArray() As String
Dim stepSpecArray() As String
Dim cmdBufStr As String
Dim tmpStr As String
Dim isAllPass As Boolean
'i represent row while j represent column
Dim i, j, cnt As Integer


Private Sub Form_Load()
    subInitComPort
    subInitInterface

    lbModelName = gstrCurProjName
    txtInput.Locked = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrExit
  
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
  
    End
Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub subInitComPort()
On Error GoTo ErrExit
 
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

    MSComm1.CommPort = setTVCurrentComID
    MSComm1.Settings = setTVCurrentComBaud & ",N,8,1"
    MSComm1.InputLen = 0
        
    MSComm1.InBufferCount = 0
    MSComm1.OutBufferCount = 0
    MSComm1.InputMode = comInputModeText
        
    MSComm1.NullDiscard = False
    MSComm1.DTREnable = False
    MSComm1.EOFEnable = False
    MSComm1.RTSEnable = False
    MSComm1.SThreshold = 1
    MSComm1.RThreshold = 1
    MSComm1.InBufferSize = 1024
    MSComm1.OutBufferSize = 512
    Exit Sub

ErrExit:
        MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub subInitInterface()
    txtInput.Text = ""
    Log_Clear
End Sub

Private Sub subInitBeforeRunning()
    strSerialNo = UCase$(txtInput.Text)
    
    lbResult.Caption = "Checking"
    lbResult.BackColor = &HFFFFFF
    Log_Clear
    txtReceive.ForeColor = &H80000008
    
    cmdBufStr = ""
    tmpStr = ""
    isAllPass = True
    txtInput.Locked = True
    txtReceive.SetFocus
End Sub

Private Sub subInitAfterRunning()
    txtInput.Text = ""
    txtInput.SetFocus
    
    txtInput.Locked = False
End Sub

Private Sub subMainProcesser()
On Error GoTo ErrExit
    subInitBeforeRunning
    
    SAFE_STOP
    
    SAFE_RES_AREP "ON"
    
    SAFE_RES_AREP_ITEM "STAT,MODE,OMET,MMET"
    
    ASK_STEP_SNUM
    
    ASK_ALL_STEP_NAME
    
    ASK_ALL_STEP_SPEC cmdBufStr
    
    SAFE_STAR
    
    'subInitAfterRunning
ErrExit:
    'MsgBox Err.Description, vbCritical, Err.Source
    'MsgBox "subMainProcesser Error"
End Sub

Private Sub MSComm1_OnComm()
On Error GoTo Err
    Select Case MSComm1.CommEvent
        Case comEvReceive
            DelayMS 500
            strCommInput = MSComm1.Input
            Call textReceive
        'Case comEvSend
    End Select
Err:
    'MsgBox "MSComm1_OnComm Error"
End Sub

Private Sub textReceive()
On Error GoTo Err
    If Trim(strCommInput) <> "" And Trim(strCommInput) <> vbCr _
        And Trim(strCommInput) <> vbLf And Trim(strCommInput) <> vbCrLf Then
        Log_Info strCommInput
        
        Select Case cmdIdentifyNum
            Case 1
                initExcelObj
                
                resArray = Split(Trim(strCommInput), ",")
                
                For i = 1 To stepNum
                    'Delete LF and CR after the last element resArray(3 + (i - 1) * 4).
                    If i = stepNum Then
                        resArray(3 + (i - 1) * 4) = Replace(resArray(3 + (i - 1) * 4), Chr(13), "")
                        resArray(3 + (i - 1) * 4) = Replace(resArray(3 + (i - 1) * 4), Chr(10), "")
                    End If

                    Select Case resArray((i - 1) * 4)
                        Case "GB"
                            sht.Cells(i + lastRowNum, GB_ItmColNum) = resArray(1 + (i - 1) * 4)
                            sht.Cells(i + lastRowNum, GB_RmColNum) = resArray(2 + (i - 1) * 4)
                        Case "AC"
                            sht.Cells(i + lastRowNum, AC_VtmColNum) = resArray(1 + (i - 1) * 4)
                            sht.Cells(i + lastRowNum, AC_ImColNum) = resArray(2 + (i - 1) * 4)
                        Case "DC"
                            sht.Cells(i + lastRowNum, DC_VtmColNum) = resArray(1 + (i - 1) * 4)
                            sht.Cells(i + lastRowNum, DC_ImColNum) = resArray(2 + (i - 1) * 4)
                        Case "IR"
                            sht.Cells(i + lastRowNum, IR_VtmColNum) = resArray(1 + (i - 1) * 4)
                            sht.Cells(i + lastRowNum, IR_RmColNum) = resArray(2 + (i - 1) * 4)
                        Case "LC"
                            sht.Cells(i + lastRowNum, LC_VtmColNum) = resArray(1 + (i - 1) * 4)
                            sht.Cells(i + lastRowNum, LC_ImColNum) = resArray(2 + (i - 1) * 4)
                        Case "OSC"
                            sht.Cells(i + lastRowNum, OSC_VtmColNum) = resArray(1 + (i - 1) * 4)
                            sht.Cells(i + lastRowNum, OSC_CColNum) = resArray(2 + (i - 1) * 4)
                        Case Else
                            Log_Info "Others"
                    End Select
                    
                    'The result of each step.
                    If resArray(3 + (i - 1) * 4) = "116" Then
                        sht.Cells(i + lastRowNum, Judge_StepColNum) = "PASS"
                        isAllPass = isAllPass And True
                    Else
                        sht.Cells(i + lastRowNum, Judge_StepColNum) = "FAIL"
                        isAllPass = False
                    End If
                Next i
                
                If isAllPass = True Then
                    sht.Cells(1 + lastRowNum, Judge_TotalColNum) = "PASS"
                    Log_Info "----PASS----"
                    deInitExcelObj
                    GoTo PASS
                Else
                    'sht.Cells(1 + lastRowNum, Judge_TotalColNum) = "FAIL"
                    Log_Info "----FAIL----"
                    tmpStr = CStr(1 + lastRowNum) & ":" & CStr(stepNum + lastRowNum)
                    Log_Info tmpStr
                    sht.Rows(tmpStr).Delete Shift:=xlUp
                    deInitExcelObj
                    GoTo FAIL
                End If
            Case 5
                initExcelObj
    
                stepNum = Val(Mid(strCommInput, 2))
                'Get the last row number of an existing sheet.
                lastRowNum = sht.UsedRange.Rows.Count

                Log_Info "Step number is " & CStr(stepNum) & ". Last row number is " & CStr(lastRowNum)
                
                With sht.Range(sht.Cells(lastRowNum + 1, 1), sht.Cells(lastRowNum + stepNum, 1))
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Merge
                End With
                sht.Cells(lastRowNum + 1, SNColNum) = strSerialNo
                
                cnt = 1
                For i = lastRowNum + 1 To lastRowNum + stepNum
                    sht.Cells(i, stepNoColNum) = cnt
                    sht.Cells(i, stepNoColNum).HorizontalAlignment = xlCenter
                    sht.Cells(i, stepNoColNum).VerticalAlignment = xlCenter
                    cnt = cnt + 1
                Next i
                
                'Column "Total"
                With sht.Range(sht.Cells(lastRowNum + 1, Judge_TotalColNum), sht.Cells(lastRowNum + stepNum, Judge_TotalColNum))
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Merge
                End With
                
                'Column "Date & Time"
                With sht.Range(sht.Cells(lastRowNum + 1, dateAndTimeColNum), sht.Cells(lastRowNum + stepNum, dateAndTimeColNum))
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Merge
                End With
                sht.Cells(lastRowNum + 1, dateAndTimeColNum) = Date & vbCrLf & Time
                
                deInitExcelObj
            Case 6
                stepArray = Split(Trim(strCommInput), ",")
                cmdBufStr = ""
                
                For i = 1 To stepNum
                    Log_Info "stepArray(" & CStr(i - 1) & ") = " & stepArray(i - 1)
                    If i = stepNum Then
                        stepArray(i - 1) = Replace(stepArray(i - 1), Chr(13), "")
                        stepArray(i - 1) = Replace(stepArray(i - 1), Chr(10), "")
                    End If
                    
                    Select Case stepArray(i - 1)
                        Case "GB"
                            cmdBufStr = cmdBufStr & "SAFE:STEP" & Str(i) & ":GB:LIM:LOW?;" & vbCrLf & "SAFE:STEP" & Str(i) & ":GB:LIM?"
                            If Not i = stepNum Then
                                cmdBufStr = cmdBufStr & ";" & vbCrLf
                            End If
                        Case "AC"
                            cmdBufStr = cmdBufStr & "SAFE:STEP" & Str(i) & ":AC:LIM:LOW?;" & vbCrLf & "SAFE:STEP" & Str(i) & ":AC:LIM?"
                            If Not i = stepNum Then
                                cmdBufStr = cmdBufStr & ";" & vbCrLf
                            End If
                        Case "DC"
                            cmdBufStr = cmdBufStr & "SAFE:STEP" & Str(i) & ":DC:LIM:LOW?;" & vbCrLf & "SAFE:STEP" & Str(i) & ":DC:LIM?"
                            If Not i = stepNum Then
                                cmdBufStr = cmdBufStr & ";" & vbCrLf
                            End If
                        Case "IR"
                            cmdBufStr = cmdBufStr & "SAFE:STEP" & Str(i) & ":IR:LIM?;" & vbCrLf & "SAFE:STEP" & Str(i) & ":IR:LIM:HIGH?"
                            If Not i = stepNum Then
                                cmdBufStr = cmdBufStr & ";" & vbCrLf
                            End If
                        Case "LC"
                            cmdBufStr = cmdBufStr & "SAFE:STEP" & Str(i) & ":LC:LIM:LOW?;" & vbCrLf & "SAFE:STEP" & Str(i) & ":LC:LIM?"
                            If Not i = stepNum Then
                                cmdBufStr = cmdBufStr & ";" & vbCrLf
                            End If
                        Case "OSC"
                            cmdBufStr = cmdBufStr & "SAFE:STEP" & Str(i) & ":OSC:LIM:OPEN?;" & vbCrLf & "SAFE:STEP" & Str(i) & ":OSC:LIM:SHOR?"
                            If Not i = stepNum Then
                                cmdBufStr = cmdBufStr & ";" & vbCrLf
                            End If
                    End Select
                Next i
            Case 7
                initExcelObj
                
                stepSpecArray = Split(Trim(strCommInput), ";")
                
                For i = 1 To stepNum * 2
                    If i = stepNum * 2 Then
                        stepSpecArray(i - 1) = Replace(stepSpecArray(i - 1), Chr(13), "")
                        stepSpecArray(i - 1) = Replace(stepSpecArray(i - 1), Chr(10), "")
                    End If
                    
                    If stepSpecArray(i - 1) = "+0.000000E+00" Then
                        stepSpecArray(i - 1) = "0"
                    End If
                    
                    Log_Info "stepSpecArray(" & CStr(i - 1) & ") = " & stepSpecArray(i - 1)
                Next i
                
                For i = 1 To stepNum
                    'Log_Info "stepArray(" & CStr(i - 1) & ") = " & stepArray(i - 1)
                    Select Case stepArray(i - 1)
                        Case "GB"
                            sht.Cells(i + lastRowNum, GB_LowColNum) = stepSpecArray((i - 1) * 2)
                            sht.Cells(i + lastRowNum, GB_HighColNum) = stepSpecArray(1 + (i - 1) * 2)
                        Case "AC"
                            sht.Cells(i + lastRowNum, AC_LowColNum) = stepSpecArray((i - 1) * 2)
                            sht.Cells(i + lastRowNum, AC_HighColNum) = stepSpecArray(1 + (i - 1) * 2)
                        Case "DC"
                            sht.Cells(i + lastRowNum, DC_LowColNum) = stepSpecArray((i - 1) * 2)
                            sht.Cells(i + lastRowNum, DC_HighColNum) = stepSpecArray(1 + (i - 1) * 2)
                        Case "IR"
                            sht.Cells(i + lastRowNum, IR_LowColNum) = stepSpecArray((i - 1) * 2)
                            sht.Cells(i + lastRowNum, IR_HighColNum) = stepSpecArray(1 + (i - 1) * 2)
                        Case "LC"
                            sht.Cells(i + lastRowNum, LC_LowColNum) = stepSpecArray((i - 1) * 2)
                            sht.Cells(i + lastRowNum, LC_HighColNum) = stepSpecArray(1 + (i - 1) * 2)
                        Case "OSC"
                            sht.Cells(i + lastRowNum, OSC_OpenColNum) = stepSpecArray((i - 1) * 2)
                            sht.Cells(i + lastRowNum, OSC_ShortColNum) = stepSpecArray(1 + (i - 1) * 2)
                    End Select
                Next i
                
                deInitExcelObj
        End Select
        
        Exit Sub
    Else
        Exit Sub
    End If
    
PASS:
    lbResult.Caption = "PASS"
    lbResult.BackColor = &HFF00&
    Call subInitAfterRunning
    Exit Sub

FAIL:
    lbResult.Caption = "NG"
    lbResult.BackColor = &HFF&
    Call subInitAfterRunning
    Exit Sub

Err:
    Log_Info "Unknown message"

End Sub

Private Sub tbSetComPort_Click()
    frmComPort.Show
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
On Error GoTo ErrExit
    'ASCII = 13 means "Enter" of keyboard.
    If KeyAscii = 13 Then
        If txtInput.Locked = False Then
            If MSComm1.PortOpen = False Then
                MSComm1.PortOpen = True
            End If
            subMainProcesser
        End If
    End If
    Exit Sub

ErrExit:
    'Invalid Port Number
    If Err.Number = 8002 Then
        MsgBox Err.Description, vbCritical, Err.Source
    End If
End Sub
