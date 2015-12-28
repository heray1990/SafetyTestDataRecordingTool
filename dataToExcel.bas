Attribute VB_Name = "Module5"
Option Explicit

Public exl As Object
Public wb As Object
Public sht As Object

Public Function initExcelObj()
    Set exl = CreateObject("Excel.Application")
    Set wb = exl.Workbooks.Open(App.path & "\data.xls")
    Set sht = wb.ActiveSheet
End Function

Public Function deInitExcelObj()
    exl.ActiveWorkbook.Save
    exl.ActiveWorkbook.Close
    exl.Quit
    
    Set sht = Nothing
    Set wb = Nothing
    Set exl = Nothing
End Function
