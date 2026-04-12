Attribute VB_Name = "CommonUtilities"
Option Explicit

Public Sub PutDummyMessageInActiveCell()
    ActiveCell.Value = "Dummy message from CommonUtilities"
    ActiveCell.Offset(1, 0).Value = "Checked at " & Format$(Now, "yyyy/mm/dd hh:mm")
End Sub

Public Sub BuildDemoTable()
    With ActiveSheet
        .Range("A1").Value = "No"
        .Range("B1").Value = "Label"
        .Range("C1").Value = "Status"
        .Range("A2").Value = 1
        .Range("B2").Value = "Alpha"
        .Range("C2").Value = "Ready"
        .Range("A3").Value = 2
        .Range("B3").Value = "Beta"
        .Range("C3").Value = "Waiting"
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C3").Borders.LineStyle = xlContinuous
        .Columns("A:C").AutoFit
    End With
End Sub

Public Sub AddDemoRow()
    Dim targetSheet As Worksheet
    Dim nextRow As Long

    Set targetSheet = ActiveSheet
    nextRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row + 1

    If nextRow < 2 Then
        nextRow = 2
    End If

    targetSheet.Cells(nextRow, 1).Value = nextRow - 1
    targetSheet.Cells(nextRow, 2).Value = "Row-" & Format$(nextRow - 1, "00")
    targetSheet.Cells(nextRow, 3).Value = "Added at " & Format$(Now, "hh:mm")
End Sub
