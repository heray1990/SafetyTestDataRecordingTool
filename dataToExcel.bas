Attribute VB_Name = "dataToExcel"
Option Explicit

Public exl As Object
Public wb As Object
Public sht As Object

Public Function executeExcel()
    Set exl = CreateObject("Excel.Application")
    Set wb = exl.Workbooks.Open(App.path & "\data.xls")
    Set sht = wb.ActiveSheet
End Function

