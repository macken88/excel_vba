Attribute VB_Name = "BookB_TestMacros"
Option Explicit

Public Sub CreateSampleHeaderRow()
    Dim targetSheet As Worksheet

    Set targetSheet = ActiveSheet

    targetSheet.Range("A1").Value = "Item"
    targetSheet.Range("B1").Value = "Value"
    targetSheet.Range("C1").Value = "UpdatedAt"
    targetSheet.Range("A1:C1").Font.Bold = True
    targetSheet.Range("A1:C1").Interior.Color = RGB(221, 235, 247)
End Sub

Public Sub FillSampleRow()
    Dim nextRow As Long
    Dim targetSheet As Worksheet

    Set targetSheet = ActiveSheet
    nextRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row + 1

    If nextRow < 2 Then
        nextRow = 2
    End If

    targetSheet.Cells(nextRow, 1).Value = "Sample"
    targetSheet.Cells(nextRow, 2).Value = 1
    targetSheet.Cells(nextRow, 3).Value = Now
    targetSheet.Cells(nextRow, 3).NumberFormatLocal = "yyyy/mm/dd hh:mm"
End Sub
