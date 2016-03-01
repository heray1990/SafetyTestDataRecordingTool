Attribute VB_Name = "SettingXMLHelper"
'**********************************************
' Handling setting.xml for the application
'**********************************************

Option Explicit

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Const gstrDelimiterForProjName As String = "-"

' Return current project's name.
Public Function GetCurProjectName() As String
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(App.Path & "\setting.xml")
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        Dim objNode As MSXML2.IXMLDOMNode
        
        Set objNode = xmlDoc.selectSingleNode("/setting/current_project")
        
        If objNode Is Nothing Then
            MsgBox "There is not <current_project> node in setting.xml."
            GetCurProjectName = "???"
        Else
            GetCurProjectName = objNode.Text
        End If
    End If
End Function

' Save current project's name.
Public Sub SetCurProjectName(strCurProjectName As String)
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(App.Path & "\setting.xml")
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        Dim objNode As MSXML2.IXMLDOMNode
        
        Set objNode = xmlDoc.selectSingleNode("/setting/current_project")
        objNode.Text = strCurProjectName
        
        xmlDoc.Save App.Path & "\setting.xml"
    End If
End Sub

' Return the list of projects' name.
Public Function GetProjectList() As Collection
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    Dim colProjectList As Collection
    
    Set colProjectList = New Collection
    
    success = xmlDoc.Load(App.Path & "\setting.xml")
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        Dim objNodeList As MSXML2.IXMLDOMNodeList
        
        Set objNodeList = xmlDoc.selectNodes("/setting/project_list/project")
        
        If Not objNodeList Is Nothing Then
            Dim objNode As MSXML2.IXMLDOMNode
            Dim brand, model As String
            
            For Each objNode In objNodeList
                brand = Trim(objNode.selectSingleNode("@brand").Text)
                model = Trim(objNode.selectSingleNode("@model").Text)
                colProjectList.Add brand & gstrDelimiterForProjName & model
            Next objNode
        End If
    End If
    
    Set GetProjectList = colProjectList
    Set colProjectList = Nothing
End Function

Public Function LoadComBaud() As String
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean

    If Not CBool(PathFileExists(App.Path & "\setting.xml")) Then
        MsgBox "Cannot open " & App.Path & "\setting.xml" & " file."
        End
    End If
    
    success = xmlDoc.Load(App.Path & "\setting.xml")
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        LoadComBaud = xmlDoc.selectSingleNode("/setting/common").selectSingleNode("@baud").Text
    End If
End Function

Public Sub SaveComBaud(strComBaud As String)
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(App.Path & "\setting.xml")
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        xmlDoc.selectSingleNode("/setting/common").selectSingleNode("@baud").Text = strComBaud
        xmlDoc.Save App.Path & "\setting.xml"
    End If
End Sub

Public Function LoadComId() As Integer
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean

    If Not CBool(PathFileExists(App.Path & "\setting.xml")) Then
        MsgBox "Cannot open " & App.Path & "\setting.xml" & " file."
        End
    End If
    
    success = xmlDoc.Load(App.Path & "\setting.xml")
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        LoadComId = Val(xmlDoc.selectSingleNode("/setting/common").selectSingleNode("@id").Text)
    End If
End Function

Public Sub SaveComId(strComId As String)
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(App.Path & "\setting.xml")
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        xmlDoc.selectSingleNode("/setting/common").selectSingleNode("@id").Text = strComId
        xmlDoc.Save App.Path & "\setting.xml"
    End If
End Sub
