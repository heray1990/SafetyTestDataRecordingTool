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
   StartUpPosition =   2  'CenterScreen
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
    Dim strProjectName As Variant
    
    cmbModelName.Clear
    
    For Each strProjectName In GetProjectList
        cmbModelName.AddItem strProjectName
    Next strProjectName

    setTVCurrentComBaud = LoadComBaud
    setTVCurrentComID = LoadComId
    cmbModelName.Text = GetCurProjectName
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Exit Sub
    
ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrExit
    gstrCurProjName = cmbModelName.Text
    SetCurProjectName gstrCurProjName

    Form1.Show
    Exit Sub
    
ErrExit:
    MsgBox ("The Licence Key is Wrong.")
    
End Sub
