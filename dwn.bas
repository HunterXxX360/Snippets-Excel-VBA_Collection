Option Explicit

Function Dwn(sheetName As String, columnKey As Integer, Optional rowKey As Integer = 1) As Integer
    With ThisWorkbook.Sheets(sheetName)
        If .Cells(rowKey + 1, columnKey).Value = "" Then
            Dwn = rowKey
        Else
            Dwn = .Cells(rowKey, columnKey).End(xlDown).Row
        End If
    End With
End Function
