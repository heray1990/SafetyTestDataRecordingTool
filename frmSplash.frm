VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2190
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cmbModelName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Sample1"
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Version "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "安规测试数据保存工具"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3300
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "请选择机型:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ss As Boolean

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

On Error GoTo ErrExit
    ss = False

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
 
    sqlstring = "select * from settingTable"
    Executesql (sqlstring)

    If rs.EOF = False Then
        rs.MoveFirst
        cmbModelName.Clear

        Do While Not rs.EOF
            cmbModelName.AddItem rs.Fields("Mark")
            rs.MoveNext
        Loop
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
        End
    End If
    
    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""
    
    sqlstring = "select * from CommonTable where Mark='ATS'"
    Executesql (sqlstring)
    
    If rs.EOF = False Then
        strCurrentModelName = rs("CurrentModelName")
        strDataVersion = rs("DataVersion")
        setTVCurrentComID = rs("ComID")
        setData = rs("Date")
        setDay = rs("Day")
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
    End If
    
    Set cn = Nothing
    Set rs = Nothing

    sqlstring = ""
    cmbModelName = strCurrentModelName

    If setData <> Day(Date) Then
        sqlstring = "select * from CommonTable where Mark='ATS'"
        Executesql (sqlstring)
        rs.Fields(4) = Day(Date)
        rs.Fields(5) = setDay + 1
        rs.Update

        Set cn = Nothing
        Set rs = Nothing
        sqlstring = ""
    End If
    Exit Sub
    
ErrExit:
       MsgBox Err.Description, vbCritical, Err.Source
       
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrExit

    strCurrentModelName = cmbModelName
    sqlstring = ""
    sqlstring = "update CommonTable set CurrentModelName='" & strCurrentModelName & "' where Mark='ATS'"
    Executesql (sqlstring)
    
    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""
    
    sqlstring = "select * from settingTable where Mark='" & strCurrentModelName & "'"
    Executesql (sqlstring)

    setTVCurrentComBaud = rs("ComBaud")

    Set rs = Nothing
    Set cn = Nothing
    sqlstring = ""

    Form1.Show

    Exit Sub
    
ErrExit:
    MsgBox ("The Licence Key is Wrong.")
    
End Sub



Private Function ATS() As Boolean

    Dim path As String
    Dim a As String
    Dim b As String
    Dim c As String
    Dim d As String
    Dim oldkey As String
    Dim i%
    Dim key As Single
    Dim fso As New FileSystemObject
    Dim Hdid
    Dim hardwareid As String
    
    Set fso = CreateObject("scripting.filesystemobject")
    Set Hdid = fso.GetDrive("C:")
    hardwareid = Hex(Hdid.SerialNumber)
    path = App.path

On Error GoTo SSS

kk:
    ATS = False

    Open ("C:\source.dll") For Input As #1
    Input #1, b
    Close #1
    Open ("C:\sys.dat") For Input As #2
    Input #2, c
    Close #2

    key = Val(b)

    If key < 283155 And key > 282950 And c = "3" + hardwareid Then
        If Month(Date) > 4 And Month(Date) < 7 And Year(Date) = 2014 Then
            ATS = True
        Else
            a = Str$(key + 100)
        End If
  
        a = Str$(key + 1)
        Open ("C:\source.dll") For Output As #8
        Print #8, a
        Close #8
    
        Exit Function
    Else
        If i > 1 Then
            MsgBox ("Please apply for a licensed APP.")
            Unload frmSplash
            End
        End If
  
        GoTo SSS
        End
        Exit Function
    End If

SSS:
    For i = 1 To 3

    a = InputBox("" & vbNewLine & "     Please Input The Licence Key.", "LICENCE")
    
    If a = "2829558" Then
        a = Str$(Val(key) + Val(Left$(a, 6)))
        Open ("C:\source.dll") For Output As #5
        Print #5, a
        Close #5

        i = i + 1
        GoTo kk
        Exit Function
    Else
        If a = "DIPHD@23456" Then MsgBox (hardwareid)
            
        MsgBox ("The Licence Key is Wrong.")
    End If

    Next i
    End
    Exit Function

End Function
