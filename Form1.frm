VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   Caption         =   "安规测试数据保存工具"
   ClientHeight    =   7215
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   7200
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "writeExcel"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   6480
      Top             =   360
   End
   Begin VB.Frame Frame2 
      Caption         =   "测试结果"
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
            Name            =   "微软雅黑"
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
      Caption         =   "条码"
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
            Name            =   "微软雅黑"
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
         Name            =   "微软雅黑"
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
      Caption         =   "设置串口"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strCommInput As String

Private Sub Command1_Click()
    'i represent row while j represent column
    Dim i, j, cnt, lastRowNum As Integer
   
    initExcelObj
    'Get the last row number of an existing sheet.
    lastRowNum = sht.UsedRange.Rows.Count
    
    sht.Cells(lastRowNum + 1, 1) = "SN Num"
    With sht.Range(sht.Cells(lastRowNum + 1, 1), sht.Cells(lastRowNum + 2, 1))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Merge
    End With
    
    cnt = 1
    For i = lastRowNum + 1 To lastRowNum + 2
        sht.Cells(i, 2) = cnt
        sht.Cells(i, 2).HorizontalAlignment = xlCenter
        sht.Cells(i, 2).VerticalAlignment = xlCenter
        cnt = cnt + 1
    Next i
    
    cnt = 100
    For i = lastRowNum + 1 To (lastRowNum + 2)
        For j = 3 To 29
            sht.Cells(i, j) = cnt
            cnt = cnt + 1
        Next j
    Next i
    
    'sht.Range(sht.Cells(21, 1), sht.Cells(21, 3)).Merge
    
    deInitExcelObj
End Sub

Private Sub Form_Load()
    setTVCurrentComBaud = 9600
    subInitComPort
    subInitInterface

    lbModelName = strCurrentModelName
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
    sqlstring = "select * from CommonTable where Mark='ATS'"
    Executesql (sqlstring)

    If rs.EOF = False Then
        setTVCurrentComID = rs("ComID")
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
    End
    End If

    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""

    ComInit
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
End Sub

Private Sub subInitAfterRunning()
    txtInput.Text = ""
    txtInput.SetFocus
End Sub

Private Sub subMainProcesser()
On Error GoTo ErrExit
    subInitBeforeRunning
    
    SAFE_STOP
    
    SAFE_RES_AREP "OFF"
    DelayMS 500
    
    ASK_STEP_SNUM
    DelayMS 500
    
    SAFE_STAR
    DelayMS 500
    
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
    'i represent row while j represent column
    Dim i, j, cnt, lastRowNum As Integer

On Error GoTo Err
    If Trim(strCommInput) <> "" And Trim(strCommInput) <> vbCr _
        And Trim(strCommInput) <> vbLf And Trim(strCommInput) <> vbCrLf Then
        Log_Info strCommInput
        'If Trim(strCommInput) = """PASS""" & vbCrLf Or Trim(strCommInput) = """PASS""" & vbLf Then
        '    Log_Info "Pass"
        '    GoTo PASS
        'Else
        '    Log_Info "Fail"
        '    GoTo FAIL
        'End If
        Select Case cmdIdentifyNum
            Case 4
                stepNum = Val(Mid(strCommInput, 2))
                
                initExcelObj
                'Get the last row number of an existing sheet.
                lastRowNum = sht.UsedRange.Rows.Count
                
                sht.Cells(lastRowNum + 1, 1) = strSerialNo
                With sht.Range(sht.Cells(lastRowNum + 1, 1), sht.Cells(lastRowNum + stepNum, 1))
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Merge
                End With
                
                cnt = 1
                For i = lastRowNum + 1 To lastRowNum + stepNum
                    sht.Cells(i, 2) = cnt
                    sht.Cells(i, 2).HorizontalAlignment = xlCenter
                    sht.Cells(i, 2).VerticalAlignment = xlCenter
                    cnt = cnt + 1
                Next i
                
                deInitExcelObj
        End Select
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
    Form2.Show
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    'ASCII = 13 means "Enter" of keyboard.
    If KeyAscii = 13 Then
        subMainProcesser
    End If
End Sub
